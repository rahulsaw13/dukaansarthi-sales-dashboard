/**
 * DukaanSarthi — Daily Sales Report  →  Google Sheet  →  Dashboard
 *
 * Creates TWO forms for your salespeople:
 *   1. Daily Report    — once every evening (calls, follow-ups, CRM, day totals)
 *   2. Shop Visit Log  — once for EVERY shop visited / demo given (shop details, needs, interest,
 *                        objections, outcome, next follow-up)
 *
 * HOW TO USE (one time, ~5 minutes)
 *   1. Open https://script.google.com  →  "New project"  →  delete the sample code, paste this whole file.
 *   2. Edit CONFIG below: salesperson names + monthly salary.
 *   3. Project Settings (gear icon, left side)  →  Time zone  →  "(GMT+05:30) India Standard Time".
 *   4. At the top, pick the function "setup"  →  Run  →  allow the permissions Google asks for.
 *   5. The form links and dashboard link appear in the Execution log and are emailed to you.
 *      Send BOTH form links to your salespeople on WhatsApp. Open the DASHBOARD link yourself.
 *   (Already ran an older version of setup()? Don't run it again — run addShopVisitLog() instead.)
 *
 * Run setup() only ONCE — every run creates brand-new forms + sheet.
 * Do NOT reorder or delete questions in the forms later — the dashboard reads answers by column.
 * Adding a new question at the END of a form is safe.
 */

const CONFIG = {
  salespeople: [
    { name: 'Salesperson 1', salary: 15000 },
    { name: 'Salesperson 2', salary: 15000 },
  ],
  // Targets from the Sales KPI & Incentive Policy. You can also change these later in the "Settings" tab.
  targets: {
    VisitsPerFieldDay: 12,
    CallsPerOfficeDay: 40,
    DemosPerWeek: 8,
    LeadsPerWeek: 20,
    SaleLeadCredit: 3,        // 1 sale = 3 interested leads
    TargetMultiple: 3,        // monthly target = 3 × salary
    IncentivePct: 0.04,       // 4% on collected sales above target
    SameDayCloseReward: 250,  // on-spot: sale closed on demo day (₹200–300)
    AnnualPlanReward: 500,    // on-spot: annual / prepaid plan
    FiveLeadDayReward: 150,   // on-spot: 5 interested leads in one day
  },
  dailyEmailHour: 21,         // 9 PM daily summary email to you. Set to null to turn off.
};

const TARGET_LABELS = {
  VisitsPerFieldDay: 'Shop visits per field day',
  CallsPerOfficeDay: 'Calls per office day',
  DemosPerWeek: 'Demos per week',
  LeadsPerWeek: 'Interested leads per week (incl. sale credit)',
  SaleLeadCredit: '1 sale counts as this many leads',
  TargetMultiple: 'Monthly target = salary ×',
  IncentivePct: 'Incentive % above target',
  SameDayCloseReward: 'On-spot ₹: same-day demo close',
  AnnualPlanReward: 'On-spot ₹: annual / prepaid plan',
  FiveLeadDayReward: 'On-spot ₹: 5 leads in one day',
};

// Column of each answer in the "Responses" sheet (A = timestamp, then form questions in order)
const RESP = 'Responses';
const COL = {
  name: 'B', date: 'C', dayType: 'D', visits: 'E', calls: 'F', demos: 'G', leads: 'H',
  fuDue: 'I', fuDone: 'J', sales: 'K', booked: 'L', collected: 'M', annual: 'N',
  sameDay: 'O', crm: 'P', complaints: 'Q',
};

const MONEY = '"₹"#,##0';
const col = c => `${RESP}!$${c}$2:$${c}`;

function sumF(field, who, from, to) {
  return `SUMIFS(${col(COL[field])},${col(COL.name)},${who},${col(COL.date)},">="&${from},${col(COL.date)},"<="&${to})`;
}

// Extra criteria as pairs: countF(who, from, to, 'dayType', '"Field Day"', 'crm', '"Yes"')
function countF(who, from, to, ...pairs) {
  let f = `COUNTIFS(${col(COL.name)},${who},${col(COL.date)},">="&${from},${col(COL.date)},"<="&${to}`;
  for (let i = 0; i < pairs.length; i += 2) f += `,${col(COL[pairs[i]])},${pairs[i + 1]}`;
  return f + ')';
}

// Weekly score out of 100 — same rules as the Dashboard's weekly table, in one formula (used by History)
function scoreFormula(who, from) {
  const to = `(${from}+6)`;
  return `=LET(repDays,${countF(who, from, to, 'dayType', '"<>Leave"')},` +
    `fieldDays,${countF(who, from, to, 'dayType', '"Field Day"')},` +
    `officeDays,${countF(who, from, to, 'dayType', '"Office Day"')},` +
    `leads,${sumF('leads', who, from, to)},sales,${sumF('sales', who, from, to)},` +
    `visits,${sumF('visits', who, from, to)},demos,${sumF('demos', who, from, to)},` +
    `fuDue,${sumF('fuDue', who, from, to)},fuDone,${sumF('fuDone', who, from, to)},` +
    `crmDays,${countF(who, from, to, 'crm', '"Yes"', 'dayType', '"<>Leave"')},calls,${sumF('calls', who, from, to)},` +
    `IF(repDays=0,"",ROUND(` +
    `MIN(40,(leads+SaleLeadCredit*sales)/LeadsPerWeek*40)` +
    `+IF(fieldDays=0,0,MIN(12.5,visits/(fieldDays*VisitsPerFieldDay)*12.5))+MIN(12.5,demos/DemosPerWeek*12.5)` +
    `+IF(fuDue=0,10,MIN(10,fuDone/fuDue*10))+crmDays/repDays*10` +
    `+IF(officeDays=0,0,MIN(15,calls/(officeDays*CallsPerOfficeDay)*15)),0)))`;
}

function colLetter(n) {
  let s = '';
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

function setup() {
  const names = CONFIG.salespeople.map(p => p.name);
  const n = names.length;
  const nameRef = i => `Settings!$A$${3 + i}`;

  // ---------- 1. Google Form ----------
  const form = FormApp.create('DukaanSarthi — Daily Sales Report');
  form.setDescription('Fill this ONCE at the end of every working day (or on leave days, choose "Leave").\n' +
    'Enter only today\'s numbers. Write 0 if none.\n' +
    'Count a lead as "interested" only if: you spoke to the owner, they have a real need, details are in CRM, and a follow-up/demo date is fixed.');
  form.setConfirmationMessage('Thanks! Your report is saved. Made a mistake? Use "Edit your response" instead of submitting again.');
  form.setAllowResponseEdits(true);

  const whole = () => FormApp.createTextValidation().setHelpText('Whole number only — enter 0 if none').requireWholeNumber().build();
  const money = () => FormApp.createTextValidation().setHelpText('Amount in ₹, numbers only — enter 0 if none').requireNumberGreaterThanOrEqualTo(0).build();
  const numQ = (title, help, v) => form.addTextItem().setTitle(title).setHelpText(help).setRequired(true).setValidation(v || whole());

  form.addListItem().setTitle('Sales executive').setChoiceValues(names).setRequired(true);                          // B
  form.addDateItem().setTitle('Report date').setRequired(true);                                                      // C
  form.addMultipleChoiceItem().setTitle('Day type').setChoiceValues(['Field Day', 'Office Day', 'Leave']).setRequired(true); // D
  numQ('Shop / business visits', 'Field day target: 12–15. Enter 0 on office days.');                                // E
  numQ('Calls made', 'Office day target: 40–50.');                                                                  // F
  numQ('Product demos given', 'Weekly target: 8.');                                                                  // G
  numQ('Interested (qualified) leads', 'Only leads that meet ALL the interested-lead rules.');                       // H
  numQ('Follow-ups due today', 'How many follow-ups were scheduled for today.');                                     // I
  numQ('Follow-ups done today', 'How many of those you actually did.');                                              // J
  numQ('Sales closed (number of clients)', '');                                                                      // K
  numQ('Sales value booked (₹)', 'Total value of sales closed today.', money());                                     // L
  numQ('Payment collected today (₹)', 'Money actually received today. Incentive is paid only on collected money.', money()); // M
  numQ('Annual / prepaid plans sold', '');                                                                           // N
  numQ('Sales closed on the same day as the demo', '');                                                              // O
  form.addMultipleChoiceItem().setTitle('CRM / daily report updated today?').setChoiceValues(['Yes', 'No']).setRequired(true); // P
  numQ('Client complaints or cancellations today', '');                                                              // Q
  form.addParagraphTextItem().setTitle('Lead details')
    .setHelpText('One lead per line:  Shop name – Owner name – Phone – Follow-up/Demo date');                         // R
  form.addParagraphTextItem().setTitle('Tomorrow\'s plan').setRequired(true);                                        // S
  form.addParagraphTextItem().setTitle('Remarks / problems faced');                                                  // T

  // ---------- 2. Spreadsheet linked to the form ----------
  const created = SpreadsheetApp.create('DukaanSarthi — Sales Dashboard');
  const ssId = created.getId();
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ssId);
  const resp = findResponseSheet_(ssId, form.getId());
  const ss = SpreadsheetApp.openById(ssId);
  ss.setSpreadsheetTimeZone('Asia/Kolkata');
  resp.setName(RESP);
  resp.setFrozenRows(1);
  resp.getRange('C:C').setNumberFormat('dd-mmm-yyyy');
  resp.getRange('L:M').setNumberFormat(MONEY);

  const dash = ss.getSheets().find(s => !formUrlOf_(s));
  dash.setName('Dashboard');

  // ---------- 3. Settings ----------
  const set = ss.insertSheet('Settings');
  set.getRange('A1').setValue('Salespeople (names must match the form dropdown exactly)').setFontWeight('bold');
  set.getRange(2, 1, 1, 2).setValues([['Name', 'Monthly salary (₹)']]).setFontWeight('bold').setBackground('#e8eaf6');
  set.getRange(3, 1, n, 2).setValues(CONFIG.salespeople.map(p => [p.name, p.salary]));
  set.getRange(3, 2, n, 1).setNumberFormat(MONEY);
  const keys = Object.keys(CONFIG.targets);
  set.getRange('D1').setValue('Targets & rewards (change anytime — dashboard updates)').setFontWeight('bold');
  set.getRange(2, 4, 1, 2).setValues([['Item', 'Value']]).setFontWeight('bold').setBackground('#e8eaf6');
  keys.forEach((k, i) => {
    const r = 3 + i;
    set.getRange(r, 4, 1, 2).setValues([[TARGET_LABELS[k], CONFIG.targets[k]]]);
    ss.setNamedRange(k, set.getRange(r, 5));
  });
  set.getRange(3 + keys.indexOf('IncentivePct'), 5).setNumberFormat('0%');
  set.getRange(3, 5, keys.length, 1).setBackground('#fff59d');
  set.getRange(3, 2, n, 1).setBackground('#fff59d');
  set.setColumnWidth(1, 160); set.setColumnWidth(2, 140); set.setColumnWidth(4, 300);

  // ---------- 4. Dashboard ----------
  dash.getRange('A1').setValue('DukaanSarthi — Sales Dashboard').setFontSize(18).setFontWeight('bold').setFontColor('#1a237e');
  dash.getRange('A2').setValue('Updates automatically from the Google Form. Change the yellow dates to look at another week or month.').setFontColor('#666666');
  dash.getRange('A4:F5').setValues([
    ['Week (pick any date)', '=TODAY()', '', 'Mon → Sun', '=B4-WEEKDAY(B4,3)', '=E4+6'],
    ['Month (pick any date)', '=TODAY()', '', '1st → last', '=DATE(YEAR(B5),MONTH(B5),1)', '=EOMONTH(B5,0)'],
  ]);
  dash.getRange('A4:A5').setFontWeight('bold');
  dash.getRange('B4:B5').setBackground('#fff59d').setNumberFormat('dd-mmm-yyyy')
    .setDataValidation(SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).build());
  dash.getRange('E4:F5').setNumberFormat('ddd dd-mmm-yyyy');

  let row = 7;
  const section = (title, headers, makeRow, totalRow) => {
    dash.getRange(row, 1).setValue(title).setFontWeight('bold').setFontSize(13).setFontColor('#1a237e');
    row++;
    const hdr = row;
    dash.getRange(row, 1, 1, headers.length).setValues([headers])
      .setFontWeight('bold').setBackground('#1a237e').setFontColor('#ffffff').setWrap(true).setVerticalAlignment('middle');
    row++;
    const first = row;
    dash.getRange(first, 1, n, headers.length).setValues(names.map((_, i) => makeRow(i, first + i)));
    row += n;
    const last = row - 1;
    let total = false;
    if (totalRow) {
      dash.getRange(row, 1, 1, headers.length).setValues([totalRow(first, last)]).setFontWeight('bold').setBackground('#eeeeee');
      row++;
      total = true;
    }
    row += 2;
    return { hdr, first, last, cols: headers.length, total };
  };

  // --- Today ---
  const T = 'TODAY()';
  const today = section('TODAY',
    ['Salesperson', 'Report today', 'Day type', 'Visits', 'Calls', 'Demos', 'Leads', 'Sales', 'Collected', 'Duplicate check'],
    (i) => {
      const w = nameRef(i);
      return [
        `=${w}`,
        `=IF(${countF(w, T, T)}>0,"✅ Submitted","❌ Not submitted")`,
        `=IFERROR(INDEX(FILTER(${col(COL.dayType)},${col(COL.name)}=${w},${col(COL.date)}=TODAY()),1),"-")`,
        `=${sumF('visits', w, T, T)}`, `=${sumF('calls', w, T, T)}`, `=${sumF('demos', w, T, T)}`,
        `=${sumF('leads', w, T, T)}`, `=${sumF('sales', w, T, T)}`, `=${sumF('collected', w, T, T)}`,
        `=IF(${countF(w, T, T)}>1,"⚠ "&${countF(w, T, T)}&" entries — delete extras in Responses","")`,
      ];
    });
  dash.getRange(today.first, 9, n, 1).setNumberFormat(MONEY);

  // --- Weekly scorecard ---
  const W = '$E$4', WE = '$F$4';
  const week = section('WEEKLY SCORECARD (selected week, Mon–Sun)',
    ['Salesperson', 'Reports', 'Field days', 'Office days', 'Leads', 'Sales', 'Visits', 'Demos',
      'Follow-ups due', 'Follow-ups done', 'CRM updated (days)', 'Calls',
      'Leads & sales /40', 'Visits & demos /25', 'Follow-up & CRM /20', 'Calls /15',
      'SCORE /100', 'Result', 'Next week', 'Last 3 weeks'],
    (i, r) => {
      const w = nameRef(i);
      const hc = colLetter(2 + i);
      return [
        `=${w}`,
        `=${countF(w, W, WE, 'dayType', '"<>Leave"')}`,
        `=${countF(w, W, WE, 'dayType', '"Field Day"')}`,
        `=${countF(w, W, WE, 'dayType', '"Office Day"')}`,
        `=${sumF('leads', w, W, WE)}`, `=${sumF('sales', w, W, WE)}`,
        `=${sumF('visits', w, W, WE)}`, `=${sumF('demos', w, W, WE)}`,
        `=${sumF('fuDue', w, W, WE)}`, `=${sumF('fuDone', w, W, WE)}`,
        `=${countF(w, W, WE, 'crm', '"Yes"', 'dayType', '"<>Leave"')}`, `=${sumF('calls', w, W, WE)}`,
        `=MIN(40,(E${r}+SaleLeadCredit*F${r})/LeadsPerWeek*40)`,
        `=IF(C${r}=0,0,MIN(12.5,G${r}/(C${r}*VisitsPerFieldDay)*12.5))+MIN(12.5,H${r}/DemosPerWeek*12.5)`,
        `=IF(B${r}=0,0,IF(I${r}=0,10,MIN(10,J${r}/I${r}*10))+K${r}/B${r}*10)`,
        `=IF(D${r}=0,0,MIN(15,L${r}/(D${r}*CallsPerOfficeDay)*15))`,
        `=IF(B${r}=0,"No reports",ROUND(SUM(M${r}:P${r}),0))`,
        `=IF(B${r}=0,"",IF(Q${r}>=85,"Excellent — on-spot bonus",IF(Q${r}>=70,"Good","Below 70")))`,
        `=IF(B${r}=0,"",IF(Q${r}>=70,"5-day (3 field + 2 office)","6-day (4 field + 2 office)"))`,
        `=IF(AND(ISNUMBER(History!${hc}2),History!${hc}2<70,ISNUMBER(History!${hc}3),History!${hc}3<70,ISNUMBER(History!${hc}4),History!${hc}4<70),"⚠ PIP review","OK")`,
      ];
    });
  dash.getRange(week.first, 13, n, 4).setNumberFormat('0.0');

  // --- Monthly target & incentive ---
  const M = '$E$5', ME = '$F$5';
  const monthHeaders = ['Salesperson', 'Salary', 'Target (3× salary)', 'Collected', '% of target', 'Above target',
    'Incentive (4%)', 'On-spot rewards (est.)', 'Booked', 'Sales', 'Leads', 'Demos', 'Lead → Demo', 'Demo → Sale',
    'Visits', 'Calls', 'Annual plans', 'Same-day closes', 'Complaints', 'Field days', 'Office days', 'Leave days'];
  const month = section('MONTHLY TARGET & INCENTIVE (selected month)', monthHeaders,
    (i, r) => {
      const w = nameRef(i);
      return [
        `=${w}`, `=Settings!$B$${3 + i}`, `=B${r}*TargetMultiple`, `=${sumF('collected', w, M, ME)}`,
        `=IF(C${r}=0,0,D${r}/C${r})`, `=MAX(0,D${r}-C${r})`, `=F${r}*IncentivePct`,
        `=R${r}*SameDayCloseReward+Q${r}*AnnualPlanReward+${countF(w, M, ME, 'leads', '">=5"')}*FiveLeadDayReward`,
        `=${sumF('booked', w, M, ME)}`, `=${sumF('sales', w, M, ME)}`, `=${sumF('leads', w, M, ME)}`, `=${sumF('demos', w, M, ME)}`,
        `=IF(K${r}=0,0,L${r}/K${r})`, `=IF(L${r}=0,0,J${r}/L${r})`,
        `=${sumF('visits', w, M, ME)}`, `=${sumF('calls', w, M, ME)}`, `=${sumF('annual', w, M, ME)}`,
        `=${sumF('sameDay', w, M, ME)}`, `=${sumF('complaints', w, M, ME)}`,
        `=${countF(w, M, ME, 'dayType', '"Field Day"')}`, `=${countF(w, M, ME, 'dayType', '"Office Day"')}`,
        `=${countF(w, M, ME, 'dayType', '"Leave"')}`,
      ];
    },
    (first, last) => monthHeaders.map((_, c) => {
      const L = colLetter(c + 1), r = last + 1;
      if (c === 0) return 'TEAM TOTAL';
      if (L === 'E') return `=IF(C${r}=0,0,D${r}/C${r})`;
      if (L === 'M') return `=IF(K${r}=0,0,L${r}/K${r})`;
      if (L === 'N') return `=IF(L${r}=0,0,J${r}/L${r})`;
      return `=SUM(${L}${first}:${L}${last})`;
    }));
  const mRows = n + 1;
  ['B', 'C', 'D', 'F', 'G', 'H', 'I'].forEach(L => dash.getRange(`${L}${month.first}:${L}${month.first + mRows - 1}`).setNumberFormat(MONEY));
  ['E', 'M', 'N'].forEach(L => dash.getRange(`${L}${month.first}:${L}${month.first + mRows - 1}`).setNumberFormat('0%'));

  // ---------- 5. History (last 12 weeks of scores) ----------
  const hist = ss.insertSheet('History');
  hist.getRange(1, 1).setValue('Week starting (Mon)');
  hist.getRange(1, 2, 1, n).setFormulas([names.map((_, i) => `=${nameRef(i)}`)]);
  const histRows = [];
  for (let k = 0; k < 12; k++) {
    const r = 2 + k;
    histRows.push([`=Dashboard!$E$4-7*${k}`].concat(names.map((_, i) => scoreFormula(`${colLetter(2 + i)}$1`, `$A${r}`))));
  }
  hist.getRange(2, 1, 12, n + 1).setValues(histRows);
  hist.getRange(2, 1, 12, 1).setNumberFormat('ddd dd-mmm-yyyy');
  hist.getRange(1, 1, 1, n + 1).setFontWeight('bold').setBackground('#1a237e').setFontColor('#ffffff');
  hist.setColumnWidth(1, 160);
  hist.getRange(14, 1).setValue('Weekly score out of 100 (blank = no reports that week). Row 2 = week selected on Dashboard.').setFontColor('#666666');

  // ---------- 6. Trend data + charts ----------
  const trend = ss.insertSheet('Trend');
  const leadStart = n + 3;
  const trendHdr = ['Date'].concat(names.map((_, i) => `=${nameRef(i)}`), [''], ['Date'], names.map((_, i) => `=${nameRef(i)}`));
  trend.getRange(1, 1, 1, trendHdr.length).setValues([trendHdr]).setFontWeight('bold');
  const trendRows = [];
  for (let d = 0; d < 31; d++) {
    const r = 2 + d;
    const line = [`=IF(Dashboard!$E$5+${d}<=Dashboard!$F$5,Dashboard!$E$5+${d},"")`];
    names.forEach((_, i) => line.push(`=IF($A${r}="","",SUMIFS(${col(COL.collected)},${col(COL.name)},${colLetter(2 + i)}$1,${col(COL.date)},$A${r}))`));
    line.push('');
    line.push(`=$A${r}`);
    names.forEach((_, i) => line.push(`=IF($A${r}="","",SUMIFS(${col(COL.leads)},${col(COL.name)},${colLetter(leadStart + 1 + i)}$1,${col(COL.date)},$A${r}))`));
    trendRows.push(line);
  }
  trend.getRange(2, 1, 31, trendHdr.length).setValues(trendRows);
  trend.getRange(2, 1, 31, 1).setNumberFormat('dd-mmm');
  trend.getRange(2, leadStart, 31, 1).setNumberFormat('dd-mmm');

  dash.getRange(row, 1).setValue('CHARTS (selected month)').setFontWeight('bold').setFontSize(13).setFontColor('#1a237e');
  const chartRow = row + 1;
  dash.insertChart(dash.newChart().setChartType(Charts.ChartType.LINE)
    .addRange(trend.getRange(1, 1, 32, 1 + n)).setNumHeaders(1)
    .setOption('title', 'Payment collected per day (₹)').setOption('legend', { position: 'bottom' })
    .setOption('width', 620).setOption('height', 320).setPosition(chartRow, 1, 0, 0).build());
  dash.insertChart(dash.newChart().setChartType(Charts.ChartType.COLUMN)
    .addRange(trend.getRange(1, leadStart, 32, 1 + n)).setNumHeaders(1)
    .setOption('title', 'Interested leads per day').setOption('legend', { position: 'bottom' })
    .setOption('width', 620).setOption('height', 320).setPosition(chartRow, 9, 0, 0).build());
  dash.insertChart(dash.newChart().setChartType(Charts.ChartType.BAR)
    .addRange(dash.getRange(month.hdr, 1, n + 1, 1))
    .addRange(dash.getRange(month.hdr, 3, n + 1, 2)).setNumHeaders(1)
    .setOption('title', 'Monthly target vs collected (₹)').setOption('legend', { position: 'bottom' })
    .setOption('width', 620).setOption('height', 300).setPosition(chartRow + 17, 1, 0, 0).build());

  // ---------- 7. Formatting ----------
  const green = '#c8e6c9', yellow = '#fff9c4', red = '#ffcdd2';
  const scoreRules = rng => [
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(85).setBackground(green).setRanges([rng]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberBetween(70, 84.999).setBackground(yellow).setRanges([rng]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(70).setBackground(red).setRanges([rng]).build(),
  ];
  const todayStatus = dash.getRange(today.first, 2, n, 1);
  const pctRange = dash.getRange(month.first, 5, mRows, 1);
  dash.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenTextContains('❌').setBackground(red).setRanges([todayStatus]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextContains('✅').setBackground(green).setRanges([todayStatus]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextContains('PIP').setBackground(red).setRanges([dash.getRange(week.first, 20, n, 1)]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextContains('6-day').setBackground(red).setRanges([dash.getRange(week.first, 19, n, 1)]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(1).setBackground(green).setRanges([pctRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberBetween(0.5, 0.9999).setBackground(yellow).setRanges([pctRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0.5).setBackground(red).setRanges([pctRange]).build(),
  ].concat(scoreRules(dash.getRange(week.first, 17, n, 1))));
  hist.setConditionalFormatRules(scoreRules(hist.getRange(2, 2, 12, n)));

  dash.setColumnWidth(1, 170);
  for (let c = 2; c <= 22; c++) dash.setColumnWidth(c, 105);
  dash.setColumnWidth(10, 150); dash.setColumnWidth(18, 170); dash.setColumnWidth(19, 180);
  dash.setRowHeight(today.hdr, 40); dash.setRowHeight(week.hdr, 40); dash.setRowHeight(month.hdr, 40);
  dash.setFrozenColumns(1);

  ss.setActiveSheet(dash); ss.moveActiveSheet(1);
  ss.setActiveSheet(resp); ss.moveActiveSheet(2);
  ss.setActiveSheet(dash);

  // ---------- 8. Save + daily email trigger ----------
  PropertiesService.getScriptProperties().setProperties({
    SS_ID: ssId,
    FORM_URL: form.getPublishedUrl(),
    LAYOUT: JSON.stringify({ today, week, month }),
  });

  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'sendDailySummary')
    .forEach(t => ScriptApp.deleteTrigger(t));
  if (CONFIG.dailyEmailHour !== null) {
    ScriptApp.newTrigger('sendDailySummary').timeBased().atHour(CONFIG.dailyEmailHour).everyDays(1).create();
  }

  const visit = buildVisitLog_();

  const me = Session.getEffectiveUser().getEmail();
  const msg = 'DAILY REPORT form (salespeople fill once every evening):\n' + shortLink_(form) +
    '\n\nSHOP VISIT form (salespeople fill once for EVERY shop visit / demo):\n' + shortLink_(visit) +
    '\n\nDASHBOARD link (for you):\n' + ss.getUrl() +
    '\n\nEdit the forms (careful — do not reorder or delete questions):\n' + form.getEditUrl() + '\n' + visit.getEditUrl();
  Logger.log(msg);
  if (Session.getScriptTimeZone() !== 'Asia/Kolkata') {
    Logger.log('⚠ Script time zone is ' + Session.getScriptTimeZone() +
      '. Set Project Settings → Time zone → India Standard Time, then run setupEmailTrigger() so the email comes at the right hour.');
  }
  MailApp.sendEmail(me, 'DukaanSarthi sales dashboard is ready', msg);
}

function shortLink_(form) {
  try { return form.shortenFormUrl(form.getPublishedUrl()); } catch (e) { return form.getPublishedUrl(); }
}

// getFormUrl() throws on tabs that are not linked to a form, so treat any error as "not linked"
function formUrlOf_(sheet) {
  try { return sheet.getFormUrl() || ''; } catch (e) { return ''; }
}

function findResponseSheet_(id, formId) {
  for (let i = 0; i < 20; i++) {
    SpreadsheetApp.flush();
    const sh = SpreadsheetApp.openById(id).getSheets().find(s => formUrlOf_(s).indexOf(formId) >= 0);
    if (sh) return sh;
    Utilities.sleep(1000);
  }
  throw new Error('Could not find the form responses sheet. Please run setup again.');
}

/** Re-create the daily email trigger (e.g. after changing the time zone or dailyEmailHour). */
function setupEmailTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'sendDailySummary')
    .forEach(t => ScriptApp.deleteTrigger(t));
  if (CONFIG.dailyEmailHour !== null) {
    ScriptApp.newTrigger('sendDailySummary').timeBased().atHour(CONFIG.dailyEmailHour).everyDays(1).create();
  }
}

/** Emails you today's status, this week's score, this month's incentive and follow-ups. Runs daily from the trigger. */
function sendDailySummary() {
  const p = PropertiesService.getScriptProperties();
  const L = JSON.parse(p.getProperty('LAYOUT'));
  const ss = SpreadsheetApp.openById(p.getProperty('SS_ID'));
  const dash = ss.getSheetByName('Dashboard');
  // Point the dashboard back at the current week / month before reading it
  dash.getRange('B4').setFormula('=TODAY()');
  dash.getRange('B5').setFormula('=TODAY()');
  SpreadsheetApp.flush();

  const read = b => dash.getRange(b.hdr, 1, b.last - b.hdr + 1 + (b.total ? 1 : 0), b.cols).getDisplayValues();
  const pick = (rows, idx) => rows.map(r => idx.map(i => r[i]));
  const table = rows => '<table style="border-collapse:collapse;font-family:Arial;font-size:13px">' +
    rows.map((r, i) => '<tr>' + r.map(c => {
      const bg = i === 0 ? '#1a237e' : (/❌|Below 70|6-day|PIP|Overdue|^Hot/.test(c) ? '#ffcdd2' : (/✅|Excellent/.test(c) ? '#c8e6c9' : '#ffffff'));
      const fg = i === 0 ? '#ffffff' : '#000000';
      return `<td style="border:1px solid #ccc;padding:6px 10px;background:${bg};color:${fg}">${c}</td>`;
    }).join('') + '</tr>').join('') + '</table>';

  const date = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'dd MMM yyyy (EEE)');
  const html =
    `<h2 style="font-family:Arial;color:#1a237e">Daily sales report — ${date}</h2>` +
    '<h3 style="font-family:Arial">Today</h3>' + table(read(L.today)) +
    '<h3 style="font-family:Arial">This week</h3>' + table(pick(read(L.week), [0, 1, 4, 5, 6, 7, 11, 16, 17, 18, 19])) +
    '<h3 style="font-family:Arial">This month</h3>' + table(pick(read(L.month), [0, 2, 3, 4, 6, 7, 9, 10])) +
    visitEmailPart_(ss, table) +
    `<p style="font-family:Arial"><a href="${ss.getUrl()}">Open full dashboard</a></p>`;

  MailApp.sendEmail({
    to: Session.getEffectiveUser().getEmail(),
    subject: 'DukaanSarthi daily sales — ' + date,
    htmlBody: html,
  });
}

// =====================================================================================
//  SHOP VISIT / DEMO LOG — salesperson fills ONE entry for EVERY shop they visit
//  Creates: "Visits" (raw answers), "Follow-ups", "Hot & Warm Leads", "All Shops",
//           "Shop Lookup" and "Visit Insights" tabs.
// =====================================================================================

const VIS = 'Visits';
// Column of each answer in the "Visits" sheet (A = timestamp, then questions in order)
const VC = {
  exec: 'B', date: 'C', visitType: 'D', shop: 'E', owner: 'F', mobile: 'G', bizType: 'H', area: 'I',
  location: 'J', metWith: 'K', currentSys: 'L', currentSw: 'M', monthlySales: 'N', staff: 'O',
  pains: 'P', products: 'Q', demo: 'R', interest: 'S', objections: 'T', outcome: 'U', plan: 'V',
  price: 'W', saleAmt: 'X', collected: 'Y', payMode: 'Z', nextDate: 'AA', nextAction: 'AB',
  source: 'AC', notes: 'AD',
};
const vc = c => `${VIS}!$${c}$2:$${c}`;

const VISIT_TYPES = ['First visit', 'Follow-up visit', 'Demo visit', 'Closing / payment collection'];
const BIZ_TYPES = ['Kirana / General store', 'Supermarket / Mart', 'Medical / Pharmacy', 'Hardware / Paint / Sanitary',
  'Electronics / Mobile shop', 'Clothing / Footwear', 'Restaurant / Cafe / Hotel', 'Bakery / Sweets',
  'Cosmetics / Gift / Stationery', 'Wholesale / Distributor'];
const MET_WITH = ['Owner / decision maker', 'Manager', 'Staff only', 'Family member of owner'];
const CURRENT_SYS = ['Manual register / notebook', 'Excel', 'Tally', 'Vyapar', 'Marg', 'Busy', 'Other billing software', 'Nothing'];
const MONTHLY_SALES = ['Below ₹1 lakh', '₹1–5 lakh', '₹5–20 lakh', 'Above ₹20 lakh', 'Not known'];
const STAFF = ['Only owner', '1–2 staff', '3–5 staff', '6+ staff'];
const PAINS = ['Billing is slow', 'Stock / inventory tracking', 'Udhaar (credit) tracking', 'GST & accounts',
  'Wants online orders / website', 'Wants WhatsApp / SMS marketing', 'Staff mistakes / theft',
  'Needs sales reports', 'Home delivery management', 'Multiple branches'];
const PRODUCTS = ['DukaanSarthi billing / ERP', 'Mobile POS app', 'Ecommerce website / online store', 'Delivery app',
  'WhatsApp & SMS marketing'];
const INTEREST = ['Hot - ready to buy', 'Warm - interested, needs follow-up', 'Cold - maybe later', 'Not interested'];
const OBJECTIONS = ['Price too high', 'Already uses other software', 'Not comfortable with computer / mobile',
  'Wants to discuss with family / partner', 'No time right now', 'Does not trust a new company', 'Internet problem',
  'No objection'];
const OUTCOMES = ['Sale closed', 'Demo scheduled', 'Follow-up scheduled', 'Call back later', 'Not interested'];
const PLANS = ['Monthly', 'Annual / prepaid', 'Not discussed'];
const PAY_MODES = ['Cash', 'UPI', 'Bank transfer', 'Cheque', 'Not paid yet'];
const SOURCES = ['Walk-in (cold visit)', 'Referral by existing client', 'Office call', 'Social media / WhatsApp', 'Other'];

// Row is the newest entry for that shop (shops are identified by mobile number)
const V_LATEST = `(COUNTIFS(${vc(VC.mobile)},${vc(VC.mobile)},${VIS}!$A$2:$A,">"&${VIS}!$A$2:$A)=0)`;
const V_OPEN = [`${vc(VC.outcome)}<>"Sale closed"`, `${vc(VC.outcome)}<>"Not interested"`, `${vc(VC.interest)}<>"Not interested"`];

/** Already ran setup() before the shop visit log existed? Run THIS once to add it to the same sheet. */
function addShopVisitLog() {
  const form = buildVisitLog_();
  const msg = 'SHOP VISIT form (salespeople fill once for EVERY shop visit / demo):\n' + shortLink_(form) +
    '\n\nEdit the form (careful — do not reorder or delete questions):\n' + form.getEditUrl();
  Logger.log(msg);
  MailApp.sendEmail(Session.getEffectiveUser().getEmail(), 'DukaanSarthi shop visit log is ready', msg);
}

function buildVisitLog_() {
  const props = PropertiesService.getScriptProperties();
  const ssId = props.getProperty('SS_ID');
  if (!ssId) throw new Error('Run setup() first.');
  if (props.getProperty('VISIT_FORM_URL')) throw new Error('The shop visit log was already added. Not creating it again.');
  const names = CONFIG.salespeople.map(p => p.name);
  const n = names.length;
  const nameRef = i => `Settings!$A$${3 + i}`;

  // ---------- 1. Form ----------
  const form = FormApp.create('DukaanSarthi — Shop Visit / Demo Log');
  form.setDescription('Fill this ONCE FOR EVERY SHOP you visit or give a demo to — while you are at the shop.\n' +
    'Same shop visited again? Fill again with the SAME mobile number — the system joins the history.');
  form.setConfirmationMessage('Saved ✅  Next shop!');
  form.setAllowResponseEdits(true);

  const money = () => FormApp.createTextValidation().setHelpText('Amount in ₹, numbers only').requireNumberGreaterThanOrEqualTo(0).build();
  const mc = (title, choices, required, other) => {
    const it = form.addMultipleChoiceItem().setTitle(title).setChoiceValues(choices).setRequired(!!required);
    if (other) it.showOtherOption(true);
    return it;
  };
  const cb = (title, choices) => form.addCheckboxItem().setTitle(title).setHelpText('Tick all that apply')
    .setChoiceValues(choices).showOtherOption(true);

  form.addListItem().setTitle('Sales executive').setChoiceValues(names).setRequired(true);                       // B
  form.addDateItem().setTitle('Visit date').setRequired(true);                                                    // C
  mc('Visit type', VISIT_TYPES, true);                                                                            // D

  form.addSectionHeaderItem().setTitle('Shop details');
  form.addTextItem().setTitle('Shop name').setRequired(true);                                                     // E
  form.addTextItem().setTitle('Owner / contact person name').setRequired(true);                                   // F
  form.addTextItem().setTitle('Mobile number').setHelpText('10 digits, no +91. Use the SAME number every time for the same shop.')
    .setRequired(true).setValidation(FormApp.createTextValidation().setHelpText('Enter a 10-digit mobile number')
      .requireTextMatchesPattern('^[6-9][0-9]{9}$').build());                                                     // G
  mc('Business type', BIZ_TYPES, true, true);                                                                     // H
  form.addTextItem().setTitle('Area / locality').setRequired(true);                                               // I
  form.addTextItem().setTitle('Location link or address')
    .setHelpText('Google Maps → tap the blue dot → Share → copy link and paste here (proof of visit).');           // J
  mc('Who did you meet?', MET_WITH, true);                                                                        // K
  cb('How do they do billing / accounts now?', CURRENT_SYS);                                                      // L
  form.addTextItem().setTitle('Current software name & what they pay (if known)');                                // M
  mc('Approx. monthly sales of the shop', MONTHLY_SALES);                                                         // N
  mc('Staff size', STAFF);                                                                                        // O

  form.addSectionHeaderItem().setTitle('Needs & interest');
  cb('Problems they told you about', PAINS);                                                                      // P
  cb('Interested in', PRODUCTS);                                                                                  // Q
  mc('Demo given?', ['Yes', 'No'], true);                                                                         // R
  mc('Interest level', INTEREST, true);                                                                           // S
  cb('Objections / reasons for not buying', OBJECTIONS);                                                          // T

  form.addSectionHeaderItem().setTitle('Result');
  mc('Outcome of this visit', OUTCOMES, true);                                                                    // U
  mc('Plan discussed', PLANS);                                                                                    // V
  form.addTextItem().setTitle('Price quoted (₹)').setValidation(money());                                         // W
  form.addTextItem().setTitle('Sale amount (₹) — only if sale closed').setValidation(money());                    // X
  form.addTextItem().setTitle('Amount collected now (₹)').setValidation(money());                                 // Y
  mc('Payment mode', PAY_MODES);                                                                                  // Z
  form.addDateItem().setTitle('Next follow-up / demo date').setHelpText('Leave blank only if sale closed or not interested.'); // AA
  form.addTextItem().setTitle('Next action').setHelpText('e.g. "Show stock module to owner\'s son", "Collect balance ₹3000"'); // AB
  mc('How did you find this shop?', SOURCES);                                                                     // AC
  form.addParagraphTextItem().setTitle('Notes').setHelpText('Anything useful: competitor, budget, best time to call, who decides…'); // AD

  // ---------- 2. Link to the existing spreadsheet ----------
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ssId);
  const vs = findResponseSheet_(ssId, form.getId());
  const ss = SpreadsheetApp.openById(ssId);
  vs.setName(VIS);
  vs.setFrozenRows(1);
  vs.setFrozenColumns(5);
  [VC.date, VC.nextDate].forEach(c => vs.getRange(`${c}:${c}`).setNumberFormat('dd-mmm-yyyy'));
  [VC.price, VC.saleAmt, VC.collected].forEach(c => vs.getRange(`${c}:${c}`).setNumberFormat(MONEY));

  const red = '#ffcdd2', yellow = '#fff9c4', green = '#c8e6c9', orange = '#ffe0b2';
  const title = (sh, t, sub) => {
    sh.getRange('A1').setValue(t).setFontSize(16).setFontWeight('bold').setFontColor('#1a237e');
    sh.getRange('A2').setValue(sub).setFontColor('#666666');
  };
  const header = (sh, row, headers) => sh.getRange(row, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setBackground('#1a237e').setFontColor('#ffffff').setWrap(true);
  const listFormula = (cols, conds, sort, empty) =>
    `=ARRAYFORMULA(IFERROR(SORT(FILTER({${cols.map(c => c[1]).join(',')}},${conds.join(',')}),${sort}),"${empty}"))`;

  // Latest status of every shop (shared by several tabs)
  const SHOP_COLS = [
    ['Last visit', vc(VC.date)], ['Shop', vc(VC.shop)], ['Owner', vc(VC.owner)], ['Mobile', vc(VC.mobile)],
    ['Area', vc(VC.area)], ['Business type', vc(VC.bizType)], ['Visits so far', `COUNTIF(${vc(VC.mobile)},${vc(VC.mobile)})`],
    ['Interest', vc(VC.interest)], ['Last outcome', vc(VC.outcome)], ['Next follow-up', vc(VC.nextDate)],
    ['Next action', vc(VC.nextAction)], ['Interested in', vc(VC.products)], ['Problems', vc(VC.pains)],
    ['Objections', vc(VC.objections)], ['Uses now', vc(VC.currentSys)], ['Software / pays', vc(VC.currentSw)],
    ['Monthly sales', vc(VC.monthlySales)], ['Staff', vc(VC.staff)], ['Price quoted', vc(VC.price)],
    ['Salesperson', vc(VC.exec)], ['Location', vc(VC.location)], ['Notes', vc(VC.notes)],
  ];
  const shopFormats = sh => {
    sh.getRange('A5:A').setNumberFormat('dd-mmm-yyyy');
    sh.getRange('J5:J').setNumberFormat('dd-mmm-yyyy');
    sh.getRange('S5:S').setNumberFormat(MONEY);
    sh.setFrozenRows(4); sh.setFrozenColumns(2);
    sh.setColumnWidth(2, 170); sh.setColumnWidth(11, 200);
  };
  const interestRules = rng => [
    SpreadsheetApp.newConditionalFormatRule().whenTextStartsWith('Hot').setBackground(red).setRanges([rng]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextStartsWith('Warm').setBackground(orange).setRanges([rng]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextStartsWith('Cold').setBackground('#e3f2fd').setRanges([rng]).build(),
  ];

  // ---------- 3. Follow-ups ----------
  const fu = ss.insertSheet('Follow-ups');
  title(fu, 'Follow-ups — overdue, today and next 7 days', 'Latest status of each shop. Closed / not-interested shops are hidden. Call these first!');
  const FU_COLS = [
    ['Status', `IF(${vc(VC.nextDate)}<TODAY(),"Overdue",IF(${vc(VC.nextDate)}=TODAY(),"Today","Upcoming"))`],
    ['Follow-up date', vc(VC.nextDate)], ['Days late', `IF(${vc(VC.nextDate)}<TODAY(),TODAY()-${vc(VC.nextDate)},0)`],
    ['Shop', vc(VC.shop)], ['Owner', vc(VC.owner)], ['Mobile', vc(VC.mobile)], ['Area', vc(VC.area)],
    ['Interest', vc(VC.interest)], ['Last outcome', vc(VC.outcome)], ['Next action', vc(VC.nextAction)],
    ['Salesperson', vc(VC.exec)], ['Last visit', vc(VC.date)], ['Notes', vc(VC.notes)],
  ];
  header(fu, 4, FU_COLS.map(c => c[0]));
  fu.getRange('A5').setFormula(listFormula(FU_COLS,
    [`${vc(VC.shop)}<>""`, V_LATEST, `${vc(VC.nextDate)}<>""`, `${vc(VC.nextDate)}<=TODAY()+7`].concat(V_OPEN),
    '2,TRUE', 'Nothing due in the next 7 days 🎉'));
  fu.getRange('B5:B').setNumberFormat('dd-mmm-yyyy');
  fu.getRange('L5:L').setNumberFormat('dd-mmm-yyyy');
  fu.setFrozenRows(4); fu.setColumnWidth(4, 170); fu.setColumnWidth(10, 220);
  fu.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Overdue').setBackground(red).setRanges([fu.getRange('A5:A')]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Today').setBackground(yellow).setRanges([fu.getRange('A5:A')]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Upcoming').setBackground(green).setRanges([fu.getRange('A5:A')]).build(),
  ].concat(interestRules(fu.getRange('H5:H'))));

  // ---------- 4. Hot & Warm leads ----------
  const hot = ss.insertSheet('Hot & Warm Leads');
  title(hot, 'Hot & Warm leads — not closed yet', 'Hot first, then by follow-up date. These are your next sales.');
  header(hot, 4, SHOP_COLS.map(c => c[0]));
  hot.getRange('A5').setFormula(listFormula(SHOP_COLS,
    [`${vc(VC.shop)}<>""`, V_LATEST, `(LEFT(${vc(VC.interest)},3)="Hot")+(LEFT(${vc(VC.interest)},4)="Warm")`].concat(V_OPEN),
    '8,TRUE,10,TRUE', 'No open hot / warm leads yet'));
  shopFormats(hot);
  hot.setConditionalFormatRules(interestRules(hot.getRange('H5:H')));

  // ---------- 5. All shops ----------
  const all = ss.insertSheet('All Shops');
  title(all, 'All shops — one row per shop, latest status', 'Your full shop database. Use Data → Create a filter to filter by area, business type, salesperson…');
  header(all, 4, SHOP_COLS.map(c => c[0]));
  all.getRange('A5').setFormula(listFormula(SHOP_COLS, [`${vc(VC.shop)}<>""`, V_LATEST], '1,FALSE', 'No shop visits logged yet'));
  shopFormats(all);
  all.setConditionalFormatRules(interestRules(all.getRange('H5:H')));

  // ---------- 6. Shop lookup ----------
  const lk = ss.insertSheet('Shop Lookup');
  title(lk, 'Shop lookup — full visit history of one shop', 'Type a mobile number or part of a shop name in the yellow cell.');
  lk.getRange('A3').setValue('Search:').setFontWeight('bold');
  lk.getRange('B3').setBackground('#fff59d').setBorder(true, true, true, true, false, false);
  lk.getRange('A5').setFormula(`={${VIS}!A1:AD1}`);
  lk.getRange('A5:AD5').setFontWeight('bold').setBackground('#1a237e').setFontColor('#ffffff').setWrap(true);
  lk.getRange('A6').setFormula(`=IF($B$3="","⬆ Type a mobile number or shop name in the yellow cell",` +
    `IFERROR(SORT(FILTER(${VIS}!$A$2:$AD,(${vc(VC.mobile)}&""=$B$3&"")+ISNUMBER(SEARCH($B$3,${vc(VC.shop)}))),1,FALSE),"No shop found"))`);
  lk.getRange('A6:A').setNumberFormat('dd-mmm-yyyy hh:mm');
  lk.getRange('C6:C').setNumberFormat('dd-mmm-yyyy');
  lk.getRange('AA6:AA').setNumberFormat('dd-mmm-yyyy');
  lk.setFrozenRows(5); lk.setColumnWidth(2, 220);

  // ---------- 7. Visit insights ----------
  const ins = ss.insertSheet('Visit Insights');
  title(ins, 'Visit insights', '');
  ins.getRange('A2').setFormula('="Month shown: "&TEXT(Dashboard!$E$5,"mmmm yyyy")&"  (change it with the yellow Month cell on the Dashboard)"');
  const has = `(${vc(VC.shop)}<>"")*${V_LATEST}`;
  const open = V_OPEN.map(c => `(${c})`).join('*');
  const kpis = [
    ['Follow-ups OVERDUE', `=SUMPRODUCT(${has}*(${vc(VC.nextDate)}<>"")*(${vc(VC.nextDate)}<TODAY())*${open})`],
    ['Follow-ups due TODAY', `=SUMPRODUCT(${has}*(${vc(VC.nextDate)}=TODAY())*${open})`],
    ['Open HOT leads', `=SUMPRODUCT(${has}*(LEFT(${vc(VC.interest)},3)="Hot")*${open})`],
    ['Open WARM leads', `=SUMPRODUCT(${has}*(LEFT(${vc(VC.interest)},4)="Warm")*${open})`],
    ['Shops in database', `=COUNTUNIQUE(${vc(VC.mobile)})`],
  ];
  ins.getRange(4, 1, kpis.length, 2).setValues(kpis);
  ins.getRange(4, 1, kpis.length, 1).setFontWeight('bold');
  ins.getRange(4, 2, kpis.length, 1).setFontSize(14).setFontWeight('bold').setHorizontalAlignment('center').setBackground('#e8eaf6');

  const F = 'Dashboard!$E$5', T = 'Dashboard!$F$5';
  const inMonth = `${vc(VC.date)},">="&${F},${vc(VC.date)},"<="&${T}`;
  const vCount = (who, ...pairs) => {
    let f = `COUNTIFS(${who ? `${vc(VC.exec)},${who},` : ''}${inMonth}`;
    for (let i = 0; i < pairs.length; i += 2) f += `,${vc(pairs[i])},${pairs[i + 1]}`;
    return f + ')';
  };
  const vSum = (field, who) => `SUMIFS(${vc(field)},${vc(VC.exec)},${who},${inMonth})`;

  // Per salesperson (selected month)
  let row = 11;
  ins.getRange(row - 1, 1).setValue('BY SALESPERSON (selected month)').setFontWeight('bold').setFontSize(13).setFontColor('#1a237e');
  const pHeaders = ['Salesperson', 'Visits logged', 'Unique shops', 'New shops', 'Demos', 'Hot', 'Warm', 'Cold',
    'Not interested', 'Sales closed', 'Sale amount', 'Collected', 'Demo → Sale', 'Visits in daily report', 'Logged − reported'];
  header(ins, row, pHeaders);
  const pFirst = row + 1;
  const pRows = names.map((_, i) => {
    const w = nameRef(i), r = pFirst + i;
    return [
      `=${w}`, `=${vCount(w)}`,
      `=COUNTUNIQUEIFS(${vc(VC.mobile)},${vc(VC.exec)},${w},${inMonth})`,
      `=${vCount(w, VC.visitType, '"First visit"')}`, `=${vCount(w, VC.demo, '"Yes"')}`,
      `=${vCount(w, VC.interest, '"Hot*"')}`, `=${vCount(w, VC.interest, '"Warm*"')}`, `=${vCount(w, VC.interest, '"Cold*"')}`,
      `=${vCount(w, VC.interest, '"Not interested"')}`, `=${vCount(w, VC.outcome, '"Sale closed"')}`,
      `=${vSum(VC.saleAmt, w)}`, `=${vSum(VC.collected, w)}`, `=IF(E${r}=0,0,J${r}/E${r})`,
      `=${sumF('visits', w, F, T)}`, `=B${r}-N${r}`,
    ];
  });
  ins.getRange(pFirst, 1, n, pHeaders.length).setValues(pRows);
  const pTotal = pFirst + n;
  ins.getRange(pTotal, 1, 1, pHeaders.length).setValues([pHeaders.map((_, c) => {
    const L = colLetter(c + 1);
    if (c === 0) return 'TEAM TOTAL';
    if (L === 'M') return `=IF(E${pTotal}=0,0,J${pTotal}/E${pTotal})`;
    return `=SUM(${L}${pFirst}:${L}${pTotal - 1})`;
  })]).setFontWeight('bold').setBackground('#eeeeee');
  ins.getRange(`K${pFirst}:L${pTotal}`).setNumberFormat(MONEY);
  ins.getRange(`M${pFirst}:M${pTotal}`).setNumberFormat('0%');
  ins.getRange(pTotal + 1, 1).setValue('"Logged − reported" below 0 = visits claimed in the daily report but not logged shop-by-shop. Ask why.')
    .setFontColor('#666666');
  row = pTotal + 4;

  // Breakdown tables (selected month, whole team)
  const breakdowns = [];
  const breakdown = (titleText, choices, field, multi) => {
    ins.getRange(row - 1, 1).setValue(titleText).setFontWeight('bold').setFontSize(13).setFontColor('#1a237e');
    header(ins, row, [titleText.split(' (')[0], 'Visits', 'Demos', 'Sales closed', 'Visit → Sale']);
    const first = row + 1;
    const rows = choices.map((ch, i) => {
      const crit = multi ? `"*${ch}*"` : `"${ch}"`, r = first + i;
      return [ch, `=${vCount(null, field, crit)}`, `=${vCount(null, field, crit, VC.demo, '"Yes"')}`,
        `=${vCount(null, field, crit, VC.outcome, '"Sale closed"')}`, `=IF(B${r}=0,0,D${r}/B${r})`];
    });
    ins.getRange(first, 1, rows.length, 5).setValues(rows);
    ins.getRange(first, 5, rows.length, 1).setNumberFormat('0%');
    breakdowns.push({ titleText, hdr: row, first, last: first + rows.length - 1 });
    row = first + rows.length + 3;
  };
  breakdown('Business type (which shops buy?)', BIZ_TYPES, VC.bizType, false);
  breakdown('Objections (why they don\'t buy)', OBJECTIONS, VC.objections, true);
  breakdown('Problems shops face (what to pitch)', PAINS, VC.pains, true);
  breakdown('Interested in (which product sells)', PRODUCTS, VC.products, true);
  breakdown('Current billing method (competition)', CURRENT_SYS, VC.currentSys, true);
  breakdown('Lead source (where good leads come from)', SOURCES, VC.source, false);
  breakdown('Visit outcome', OUTCOMES, VC.outcome, false);
  breakdown('Shop size by monthly sales', MONTHLY_SALES, VC.monthlySales, false);

  ins.setColumnWidth(1, 260);
  for (let c = 2; c <= 15; c++) ins.setColumnWidth(c, 95);
  ins.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0).setBackground(red)
      .setRanges([ins.getRange(pFirst, 15, n + 1, 1)]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0).setBackground(red).setRanges([ins.getRange('B4:B5')]).build(),
  ]);

  // Charts next to the breakdowns
  const chartAt = (b, type) => ins.insertChart(ins.newChart().setChartType(type)
    .addRange(ins.getRange(b.hdr, 1, b.last - b.hdr + 1, 2)).setNumHeaders(1)
    .setOption('title', b.titleText).setOption('legend', { position: 'none' })
    .setOption('width', 560).setOption('height', 300).setPosition(b.hdr, 7, 0, 0).build());
  chartAt(breakdowns[0], Charts.ChartType.BAR);
  chartAt(breakdowns[1], Charts.ChartType.BAR);
  chartAt(breakdowns[2], Charts.ChartType.BAR);
  chartAt(breakdowns[6], Charts.ChartType.PIE);

  // ---------- 8. Tab order + save ----------
  const order = ['Dashboard', 'Follow-ups', 'Hot & Warm Leads', 'All Shops', 'Shop Lookup', 'Visit Insights',
    RESP, VIS, 'History', 'Settings', 'Trend'];
  order.forEach((name, i) => {
    const sh = ss.getSheetByName(name);
    if (sh) { ss.setActiveSheet(sh); ss.moveActiveSheet(i + 1); }
  });
  ss.setActiveSheet(ss.getSheetByName('Dashboard'));

  props.setProperty('VISIT_FORM_URL', form.getPublishedUrl());
  return form;
}

/** Follow-up + lead part of the 9 PM email. */
function visitEmailPart_(ss, table) {
  if (!PropertiesService.getScriptProperties().getProperty('VISIT_FORM_URL')) return '';
  const ins = ss.getSheetByName('Visit Insights');
  const fu = ss.getSheetByName('Follow-ups');
  if (!ins || !fu) return '';
  const kpi = ins.getRange('A4:B8').getDisplayValues();
  const fuRows = fu.getRange(4, 1, 41, 11).getDisplayValues()
    .filter((r, i) => i === 0 || r[0] === 'Overdue' || r[0] === 'Today')
    .map(r => [r[0], r[1], r[3], r[4], r[5], r[6], r[7], r[9], r[10]]);
  return '<h3 style="font-family:Arial">Shops & follow-ups</h3>' + table([['Item', 'Count']].concat(kpi)) +
    (fuRows.length > 1
      ? '<h3 style="font-family:Arial">Follow-ups overdue / due today</h3>' + table(fuRows)
      : '<p style="font-family:Arial">No follow-ups overdue or due today.</p>');
}
