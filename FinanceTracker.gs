/**
 * Company Finance Tracker — Income & Expense → Google Sheet → Dashboard
 *
 * Creates ONE form for logging every income and expense entry.
 * Dashboard shows: today's cash flow, weekly P&L, monthly P&L,
 * category breakdown, running balance, charts, daily email summary.
 *
 * HOW TO USE (one time, ~3 minutes)
 *   1. Open https://script.google.com → New project → delete sample code, paste this file.
 *   2. Edit CONFIG below: company name, who can log, categories, opening balance.
 *   3. Project Settings (gear icon) → Time zone → "(GMT+05:30) India Standard Time".
 *   4. Pick function "setup" → Run → allow the permissions Google asks for.
 *   5. Form link + dashboard link appear in Execution log and are emailed to you.
 *      Share the FORM link with anyone who logs transactions.
 *
 * Run setup() only ONCE — every run creates brand-new form + sheet.
 * Do NOT reorder or delete questions in the form later.
 * Adding a new question at the END is safe.
 */

const FINANCE_CONFIG = {
  companyName: 'My Company',           // appears in form title, sheet name, emails
  openingBalance: 0,                   // ₹ balance before this tracker starts (change in Settings tab anytime)
   loggedBy: ['Rahul', 'Accounts', 'Prince'],    // who can submit entries — must match exactly
  incomeCategories: [
    'Sales / Revenue',
    'Service charges',
    'Commission received',
    'Loan received',
    'Investment received',
    'Refund received',
    'Other income',
  ],
  expenseCategories: [
    'Salaries & wages',
    'Rent & office',
    'Utilities (electricity / internet)',
    'Travel & transport',
    'Marketing & advertising',
    'Stock / raw material purchase',
    'Equipment & assets',
    'Taxes & GST',
    'Bank charges & interest',
    'Software & subscriptions',
    'Maintenance & repairs',
    'Miscellaneous',
  ],
  dailyEmailHour: 21,   // 9 PM daily summary email to you. Set null to turn off.
};

// Column of each answer in the Finance "Responses" sheet (A = timestamp, then form questions in order)
const F_RESP = 'Responses';
const F_COL = {
  loggedBy: 'B', date: 'C', type: 'D', category: 'E',
  amount: 'F', payMode: 'G', party: 'H', ref: 'I', description: 'J',
};

// MONEY and colLetter are already declared in SalesDashboard.gs — reuse them here.
const fcol = c => `${F_RESP}!$${c}$2:$${c}`;

function sumIf_(type, from, to) {
  return `SUMIFS(${fcol(F_COL.amount)},${fcol(F_COL.type)},${type},${fcol(F_COL.date)},">="&${from},${fcol(F_COL.date)},"<="&${to})`;
}

function countTx_(from, to) {
  return `COUNTIFS(${fcol(F_COL.date)},">="&${from},${fcol(F_COL.date)},"<="&${to})`;
}

// =====================================================================================
//  SETUP — run ONCE
// =====================================================================================
function setup() {
  // ---- 1. Google Form ----
  const form = FormApp.create(`${FINANCE_CONFIG.companyName} — Income & Expense Log`);
  form.setDescription(
    'Log EVERY income and expense entry here — one entry per transaction.\n' +
    'Income categories start with "Sales / Revenue…"; expense categories start with "Salaries…".'
  );
  form.setConfirmationMessage('Saved ✅  Entry recorded.');
  form.setAllowResponseEdits(true);

  const allCats = [...FINANCE_CONFIG.incomeCategories, '─── Expense ───', ...FINANCE_CONFIG.expenseCategories];

  form.addListItem().setTitle('Logged by').setChoiceValues(FINANCE_CONFIG.loggedBy).setRequired(true);               // B
  form.addDateItem().setTitle('Transaction date').setRequired(true);                                          // C
  form.addMultipleChoiceItem().setTitle('Type').setChoiceValues(['Income', 'Expense']).setRequired(true);    // D
  form.addListItem().setTitle('Category').setChoiceValues(allCats).setRequired(true);                        // E
  form.addTextItem().setTitle('Amount (₹)').setRequired(true)
    .setValidation(FormApp.createTextValidation()
      .setHelpText('Numbers only — no ₹ sign').requireNumberGreaterThanOrEqualTo(0).build());                // F
  form.addMultipleChoiceItem().setTitle('Payment mode')
    .setChoiceValues(['Cash', 'UPI', 'Bank transfer', 'Cheque', 'Credit card']).setRequired(true);           // G
  form.addTextItem().setTitle('Paid to / Received from').setRequired(true)
    .setHelpText('Person or company name');                                                                   // H
  form.addTextItem().setTitle('Invoice / Receipt / Reference number')
    .setHelpText('Leave blank if none');                                                                      // I
  form.addParagraphTextItem().setTitle('Description / Narration')
    .setHelpText('What is this for?');                                                                        // J

  // ---- 2. Spreadsheet linked to the form ----
  const ss = SpreadsheetApp.create(`${FINANCE_CONFIG.companyName} — Finance Dashboard`);
  const ssId = ss.getId();
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ssId);
  const resp = findResponseSheet_(ssId, form.getId());
  resp.setName(F_RESP);
  resp.setFrozenRows(1);
  resp.getRange('C:C').setNumberFormat('dd-mmm-yyyy');
  resp.getRange('F:F').setNumberFormat(MONEY);

  const dash = ss.getSheets().find(s => !formUrlOf_(s));
  dash.setName('Dashboard');

  // ---- 3. Settings tab ----
  const set = ss.insertSheet('Settings');
  set.getRange('A1').setValue('Company Settings').setFontWeight('bold').setFontSize(13);
  set.getRange(2, 1, 2, 2).setValues([
    ['Company name', FINANCE_CONFIG.companyName],
    ['Opening balance (₹)', FINANCE_CONFIG.openingBalance],
  ]);
  set.getRange('B3').setBackground('#fff59d').setNumberFormat(MONEY);
  ss.setNamedRange('OpeningBalance', set.getRange('B3'));
  set.getRange('A5').setValue('Income categories (for reference)').setFontWeight('bold');
  set.getRange(6, 1, FINANCE_CONFIG.incomeCategories.length, 1)
    .setValues(FINANCE_CONFIG.incomeCategories.map(c => [c])).setFontColor('#2e7d32');
  set.getRange('C5').setValue('Expense categories (for reference)').setFontWeight('bold');
  set.getRange(6, 3, FINANCE_CONFIG.expenseCategories.length, 1)
    .setValues(FINANCE_CONFIG.expenseCategories.map(c => [c])).setFontColor('#b71c1c');
  set.setColumnWidth(1, 230); set.setColumnWidth(2, 160); set.setColumnWidth(3, 230);

  // ---- 4. Build dashboard ----
  dash.getRange('A1').setValue(`${FINANCE_CONFIG.companyName} — Finance Dashboard`)
    .setFontSize(18).setFontWeight('bold').setFontColor('#1a237e');
  dash.getRange('A2')
    .setValue('Updates automatically from the form. Change the yellow dates to view a different week or month.')
    .setFontColor('#666666');

  // Date pickers
  dash.getRange('A4:F5').setValues([
    ['Week (pick any date)', '=TODAY()', '', 'Mon → Sun', '=B4-WEEKDAY(B4,3)', '=E4+6'],
    ['Month (pick any date)', '=TODAY()', '', '1st → last', '=DATE(YEAR(B5),MONTH(B5),1)', '=EOMONTH(B5,0)'],
  ]);
  dash.getRange('A4:A5').setFontWeight('bold');
  dash.getRange('B4:B5').setBackground('#fff59d').setNumberFormat('dd-mmm-yyyy')
    .setDataValidation(SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).build());
  dash.getRange('E4:F5').setNumberFormat('ddd dd-mmm-yyyy');

  let row = 7;
  const sec = title => {
    dash.getRange(row, 1).setValue(title).setFontWeight('bold').setFontSize(13).setFontColor('#1a237e');
    row++;
  };
  const hdr = (headers, color) => {
    dash.getRange(row, 1, 1, headers.length).setValues([headers])
      .setFontWeight('bold').setBackground(color || '#1a237e').setFontColor('#ffffff').setWrap(true);
    row++;
  };

  // ---- TODAY ----
  const T = 'TODAY()';
  sec('TODAY');
  hdr(['Income', 'Expense', 'Net (cash today)', 'Transactions', 'Running balance (all time)']);
  const todayDataRow = row;
  const runBal = `OpeningBalance+SUMIFS(${fcol(F_COL.amount)},${fcol(F_COL.type)},"Income")-SUMIFS(${fcol(F_COL.amount)},${fcol(F_COL.type)},"Expense")`;
  dash.getRange(row, 1, 1, 5).setValues([[
    `=${sumIf_('"Income"', T, T)}`,
    `=${sumIf_('"Expense"', T, T)}`,
    `=A${row}-B${row}`,
    `=${countTx_(T, T)}`,
    `=${runBal}`,
  ]]);
  dash.getRange(row, 1, 1, 3).setNumberFormat(MONEY);
  dash.getRange(row, 5, 1, 1).setNumberFormat(MONEY);
  row += 3;

  // ---- WEEKLY ----
  const W = '$E$4', WE = '$F$4';
  sec('THIS WEEK (selected week, Mon–Sun)');
  hdr(['Income', 'Expense', 'Net (profit / loss)', 'Margin %', 'Transactions']);
  const weekDataRow = row;
  dash.getRange(row, 1, 1, 5).setValues([[
    `=${sumIf_('"Income"', W, WE)}`,
    `=${sumIf_('"Expense"', W, WE)}`,
    `=A${row}-B${row}`,
    `=IF(A${row}=0,0,(A${row}-B${row})/A${row})`,
    `=${countTx_(W, WE)}`,
  ]]);
  dash.getRange(row, 1, 1, 3).setNumberFormat(MONEY);
  dash.getRange(row, 4, 1, 1).setNumberFormat('0%');
  row += 3;

  // ---- MONTHLY ----
  const M = '$E$5', ME = '$F$5';
  sec('THIS MONTH (selected month)');
  hdr(['Income', 'Expense', 'Net (profit / loss)', 'Margin %', 'Transactions', 'Avg per transaction']);
  const monthDataRow = row;
  dash.getRange(row, 1, 1, 6).setValues([[
    `=${sumIf_('"Income"', M, ME)}`,
    `=${sumIf_('"Expense"', M, ME)}`,
    `=A${row}-B${row}`,
    `=IF(A${row}=0,0,(A${row}-B${row})/A${row})`,
    `=${countTx_(M, ME)}`,
    `=IF(E${row}=0,0,(A${row}+B${row})/E${row})`,
  ]]);
  dash.getRange(row, 1, 1, 3).setNumberFormat(MONEY);
  dash.getRange(row, 4, 1, 1).setNumberFormat('0%');
  dash.getRange(row, 6, 1, 1).setNumberFormat(MONEY);
  row += 3;

  // ---- INCOME BY CATEGORY ----
  sec('INCOME BY CATEGORY (selected month)');
  hdr(['Category', 'Amount (₹)', '% of total income'], '#2e7d32');
  const incFirst = row;
  FINANCE_CONFIG.incomeCategories.forEach(cat => {
    dash.getRange(row, 1, 1, 3).setValues([[
      cat,
      `=SUMIFS(${fcol(F_COL.amount)},${fcol(F_COL.category)},"${cat}",${fcol(F_COL.type)},"Income",${fcol(F_COL.date)},">="&${M},${fcol(F_COL.date)},"<="&${ME})`,
      `=IF($A$${monthDataRow}=0,0,B${row}/$A$${monthDataRow})`,
    ]]);
    row++;
  });
  const incLast = row - 1;
  dash.getRange(row, 1, 1, 2).setValues([['TOTAL', `=SUM(B${incFirst}:B${incLast})`]])
    .setFontWeight('bold').setBackground('#eeeeee');
  dash.getRange(incFirst, 2, FINANCE_CONFIG.incomeCategories.length + 1, 1).setNumberFormat(MONEY);
  dash.getRange(incFirst, 3, FINANCE_CONFIG.incomeCategories.length, 1).setNumberFormat('0%');
  row += 3;

  // ---- EXPENSE BY CATEGORY ----
  sec('EXPENSE BY CATEGORY (selected month)');
  hdr(['Category', 'Amount (₹)', '% of total expense'], '#b71c1c');
  const expFirst = row;
  FINANCE_CONFIG.expenseCategories.forEach(cat => {
    dash.getRange(row, 1, 1, 3).setValues([[
      cat,
      `=SUMIFS(${fcol(F_COL.amount)},${fcol(F_COL.category)},"${cat}",${fcol(F_COL.type)},"Expense",${fcol(F_COL.date)},">="&${M},${fcol(F_COL.date)},"<="&${ME})`,
      `=IF($B$${monthDataRow}=0,0,B${row}/$B$${monthDataRow})`,
    ]]);
    row++;
  });
  const expLast = row - 1;
  dash.getRange(row, 1, 1, 2).setValues([['TOTAL', `=SUM(B${expFirst}:B${expLast})`]])
    .setFontWeight('bold').setBackground('#eeeeee');
  dash.getRange(expFirst, 2, FINANCE_CONFIG.expenseCategories.length + 1, 1).setNumberFormat(MONEY);
  dash.getRange(expFirst, 3, FINANCE_CONFIG.expenseCategories.length, 1).setNumberFormat('0%');
  row += 3;

  // ---- RECENT TRANSACTIONS ----
  sec('RECENT TRANSACTIONS (newest first)');
  hdr(['Date', 'Type', 'Category', 'Amount', 'Paid to / From', 'Mode', 'Description', 'Ref #', 'Logged by']);
  dash.getRange(row, 1).setFormula(
    `=IFERROR(SORT(FILTER({${fcol(F_COL.date)},${fcol(F_COL.type)},${fcol(F_COL.category)},${fcol(F_COL.amount)},` +
    `${fcol(F_COL.party)},${fcol(F_COL.payMode)},${fcol(F_COL.description)},${fcol(F_COL.ref)},${fcol(F_COL.loggedBy)}},` +
    `${fcol(F_COL.date)}<>""),1,FALSE),"No transactions logged yet")`
  );
  dash.getRange(row, 1, 200, 1).setNumberFormat('dd-mmm-yyyy');
  dash.getRange(row, 4, 200, 1).setNumberFormat(MONEY);
  const recentRow = row;
  row += 23;

  // ---- CHARTS ----
  sec('CHARTS (selected month)');
  const chartRow = row;

  // Trend data (hidden sheet)
  const trend = ss.insertSheet('Trend');
  trend.getRange(1, 1, 1, 4).setValues([['Date', 'Income', 'Expense', 'Net']])
    .setFontWeight('bold').setBackground('#1a237e').setFontColor('#ffffff');
  const trendRows = [];
  for (let d = 0; d < 31; d++) {
    const r = 2 + d;
    trendRows.push([
      `=IF(Dashboard!$E$5+${d}<=Dashboard!$F$5,Dashboard!$E$5+${d},"")`,
      `=IF($A${r}="","",${sumIf_('"Income"', `$A${r}`, `$A${r}`)})`,
      `=IF($A${r}="","",${sumIf_('"Expense"', `$A${r}`, `$A${r}`)})`,
      `=IF($A${r}="","",B${r}-C${r})`,
    ]);
  }
  trend.getRange(2, 1, 31, 4).setValues(trendRows);
  trend.getRange(2, 1, 31, 1).setNumberFormat('dd-mmm');
  trend.getRange(2, 2, 31, 3).setNumberFormat(MONEY);

  // Line chart: income vs expense over month
  dash.insertChart(dash.newChart().setChartType(Charts.ChartType.COMBO)
    .addRange(trend.getRange(1, 1, 32, 4)).setNumHeaders(1)
    .setOption('title', 'Daily income vs expense (₹)')
    .setOption('seriesType', 'bars')
    .setOption('series', { 2: { type: 'line', color: '#1a237e' } })
    .setOption('colors', ['#2e7d32', '#b71c1c', '#1a237e'])
    .setOption('legend', { position: 'bottom' })
    .setOption('width', 620).setOption('height', 320).setPosition(chartRow, 1, 0, 0).build());

  // Pie chart: expense breakdown
  dash.insertChart(dash.newChart().setChartType(Charts.ChartType.PIE)
    .addRange(dash.getRange(expFirst - 1, 1, FINANCE_CONFIG.expenseCategories.length + 1, 2)).setNumHeaders(1)
    .setOption('title', 'Expense by category')
    .setOption('pieHole', 0.4)
    .setOption('legend', { position: 'right' })
    .setOption('width', 500).setOption('height', 320).setPosition(chartRow, 9, 0, 0).build());

  // ---- CONDITIONAL FORMATTING ----
  const green = '#c8e6c9', redBg = '#ffcdd2';
  const netCells = [
    dash.getRange(`C${todayDataRow}`),
    dash.getRange(`C${weekDataRow}`),
    dash.getRange(`C${monthDataRow}`),
  ];
  const balCell = [dash.getRange(`E${todayDataRow}`)];
  dash.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0).setBackground(green).setRanges(netCells).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0).setBackground(redBg).setRanges(netCells).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0).setBackground(redBg).setRanges(balCell).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Income').setFontColor('#2e7d32')
      .setRanges([dash.getRange(`B${recentRow}:B${recentRow + 199}`)]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Expense').setFontColor('#b71c1c')
      .setRanges([dash.getRange(`B${recentRow}:B${recentRow + 199}`)]).build(),
  ]);

  // ---- COLUMN WIDTHS ----
  dash.setColumnWidth(1, 240);
  for (let c = 2; c <= 9; c++) dash.setColumnWidth(c, 130);
  dash.setFrozenRows(1);
  dash.setFrozenColumns(1);

  // ---- TAB ORDER ----
  ss.setActiveSheet(dash); ss.moveActiveSheet(1);
  ss.setActiveSheet(resp); ss.moveActiveSheet(2);
  ss.setActiveSheet(ss.getSheetByName('Settings')); ss.moveActiveSheet(3);
  ss.setActiveSheet(trend); ss.moveActiveSheet(4);
  ss.setActiveSheet(dash);

  // ---- SAVE + TRIGGER ----
  PropertiesService.getScriptProperties().setProperties({
    FINANCE_SS_ID: ssId,
    FINANCE_FORM_URL: form.getPublishedUrl(),
    FINANCE_TODAY_ROW: String(todayDataRow),
    FINANCE_MONTH_ROW: String(monthDataRow),
  });

  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'sendFinanceSummary')
    .forEach(t => ScriptApp.deleteTrigger(t));
  if (FINANCE_CONFIG.dailyEmailHour !== null) {
    ScriptApp.newTrigger('sendFinanceSummary').timeBased().atHour(FINANCE_CONFIG.dailyEmailHour).everyDays(1).create();
  }

  const me = Session.getEffectiveUser().getEmail();
  const msg =
    `TRANSACTION LOG form (share with your team):\n${shortLink_(form)}\n\n` +
    `DASHBOARD (for you):\n${ss.getUrl()}\n\n` +
    `Edit the form (do not reorder or delete questions):\n${form.getEditUrl()}`;
  Logger.log(msg);
  if (Session.getScriptTimeZone() !== 'Asia/Kolkata') {
    Logger.log('⚠ Script time zone is ' + Session.getScriptTimeZone() +
      '. Set Project Settings → Time zone → India Standard Time, then run setupEmailTrigger().');
  }
  MailApp.sendEmail(me, `${FINANCE_CONFIG.companyName} finance tracker is ready`, msg);
}

// =====================================================================================
//  DAILY EMAIL — 9 PM summary (Finance)
// =====================================================================================
function sendFinanceSummary() {
  const p = PropertiesService.getScriptProperties();
  const ss = SpreadsheetApp.openById(p.getProperty('FINANCE_SS_ID'));
  const dash = ss.getSheetByName('Dashboard');
  dash.getRange('B4').setFormula('=TODAY()');
  dash.getRange('B5').setFormula('=TODAY()');
  SpreadsheetApp.flush();

  const tRow = Number(p.getProperty('FINANCE_TODAY_ROW'));
  const mRow = Number(p.getProperty('FINANCE_MONTH_ROW'));
  const date = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'dd MMM yyyy (EEE)');

  const table = rows => '<table style="border-collapse:collapse;font-family:Arial;font-size:13px">' +
    rows.map((r, i) => '<tr>' + r.map(c => {
      const bg = i === 0 ? '#1a237e' : '#ffffff';
      const fg = i === 0 ? '#ffffff' : '#000000';
      return `<td style="border:1px solid #ccc;padding:6px 10px;background:${bg};color:${fg}">${c}</td>`;
    }).join('') + '</tr>').join('') + '</table>';

  const todayVals = dash.getRange(tRow, 1, 1, 5).getDisplayValues();
  const monthVals = dash.getRange(mRow, 1, 1, 6).getDisplayValues();

  const html =
    `<h2 style="font-family:Arial;color:#1a237e">Daily finance summary — ${date}</h2>` +
    '<h3 style="font-family:Arial">Today</h3>' +
    table([['Income', 'Expense', 'Net', 'Transactions', 'Running balance'], ...todayVals]) +
    '<h3 style="font-family:Arial">This month so far</h3>' +
    table([['Income', 'Expense', 'Net', 'Margin %', 'Transactions', 'Avg'], ...monthVals]) +
    `<p style="font-family:Arial;font-size:12px"><a href="${ss.getUrl()}">Open full dashboard</a></p>`;

  MailApp.sendEmail({
    to: Session.getEffectiveUser().getEmail(),
    subject: `Finance summary — ${date}`,
    htmlBody: html,
  });
}

/** Re-create the finance daily email trigger after changing FINANCE_CONFIG.dailyEmailHour. */
function setupFinanceEmailTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'sendFinanceSummary')
    .forEach(t => ScriptApp.deleteTrigger(t));
  if (FINANCE_CONFIG.dailyEmailHour !== null) {
    ScriptApp.newTrigger('sendFinanceSummary').timeBased().atHour(FINANCE_CONFIG.dailyEmailHour).everyDays(1).create();
  }
}

// shortLink_(), formUrlOf_(), findResponseSheet_() are declared in SalesDashboard.gs and shared.
