/**
 * Google Sheets template generator for an example investment bank trade.
 *
 * How to use:
 * 1) Go to https://script.google.com and create a new project.
 * 2) Paste this file into Code.gs.
 * 3) Run createExampleInvestmentBankTradeSheet().
 * 4) Authorize when prompted.
 */
function createExampleInvestmentBankTradeSheet() {
  var ss = SpreadsheetApp.create('Example Investment Bank Trade (Illustrative)');

  var summary = ss.getSheets()[0];
  summary.setName('Trade Summary');

  summary.getRange('A1').setValue('Example Investment Bank Trade (Illustrative)');
  summary.getRange('A1').setFontWeight('bold').setFontSize(14);

  summary.getRange('A2').setValue('Note');
  summary.getRange('B2').setValue('Fictional data for demonstration only (not a recommendation or real market quote).');
  summary.getRange('A2:B2').setFontColor('#374151');

  summary.getRange('A4').setValue('Field');
  summary.getRange('B4').setValue('Value');
  summary.getRange('A4:B4').setFontWeight('bold').setBackground('#E5E7EB');

  var rows = [
    ['Trade ID', 'IRS-2026-000184'],
    ['Trade Date', new Date('2026-03-06')],
    ['Effective Date', new Date('2026-03-10')],
    ['Maturity Date', new Date('2031-03-10')],
    ['Counterparty', 'Example Pension Fund Ltd'],
    ['Booking Entity', 'Example Bank N.A.'],
    ['Product', 'Interest Rate Swap (Fixed vs Floating)'],
    ['Direction', 'Bank Pays Fixed / Receives Floating'],
    ['Notional', 25000000],
    ['Currency', 'USD'],
    ['Fixed Rate', 0.04125],
    ['Fixed Leg Day Count', '30/360'],
    ['Fixed Payment Frequency', 'Semi-Annual'],
    ['Float Index', 'SOFR (Compounded)'],
    ['Float Spread', 0.00035],
    ['Float Day Count', 'ACT/360'],
    ['Float Reset Frequency', 'Quarterly'],
    ['Clearing', 'Cleared (Example CCP)'],
    ['Portfolio / Book', 'Rates_USD_Swaps'],
    ['Trader', 'J. Doe'],
    ['Sales', 'A. Smith'],
    ['Status', 'Confirmed'],
    ['', ''],
    ['Pricing', 'Illustrative mark-to-market fields (placeholders)'],
    ['Discounting Curve', 'USD OIS (SOFR)'],
    ['Projected Floating Curve', 'USD SOFR Term Structure'],
    ['PV (USD)', -312500],
    ['DV01 (USD)', -11850]
  ];

  summary.getRange(5, 1, rows.length, 2).setValues(rows);

  summary.getRange('B5:B7').setNumberFormat('yyyy-mm-dd');
  summary.getRange('B13').setNumberFormat('0.000%');
  summary.getRange('B16').setNumberFormat('0.000%');
  summary.getRange('B10').setNumberFormat('#,##0.00');
  summary.getRange('B27:B28').setNumberFormat('#,##0.00');

  summary.setFrozenRows(4);
  summary.autoResizeColumns(1, 2);

  var cash = ss.insertSheet('Cashflows (Illustrative)');

  cash.getRange('A1').setValue('Illustrative Cashflow Schedule (first 6 periods)');
  cash.getRange('A1').setFontWeight('bold').setFontSize(14);

  var header = [[
    'Period Start',
    'Period End',
    'Pay Date',
    'Fixed Amount (Bank pays)',
    'Float Rate',
    'Spread',
    'Float Amount (Bank receives)'
  ]];

  cash.getRange(3, 1, 1, header[0].length).setValues(header);
  cash.getRange(3, 1, 1, header[0].length).setFontWeight('bold').setBackground('#E5E7EB');

  var flows = [
    [new Date('2026-03-10'), new Date('2026-06-10'), new Date('2026-06-10'), 515625, 0.03980, 0.00035, 514200],
    [new Date('2026-06-10'), new Date('2026-09-10'), new Date('2026-09-10'), 515625, 0.04010, 0.00035, 519850],
    [new Date('2026-09-10'), new Date('2026-12-10'), new Date('2026-12-10'), 515625, 0.04045, 0.00035, 529200],
    [new Date('2026-12-10'), new Date('2027-03-10'), new Date('2027-03-10'), 515625, 0.04070, 0.00035, 535900],
    [new Date('2027-03-10'), new Date('2027-06-10'), new Date('2027-06-10'), 515625, 0.04105, 0.00035, 545350],
    [new Date('2027-06-10'), new Date('2027-09-10'), new Date('2027-09-10'), 515625, 0.04125, 0.00035, 550750]
  ];

  cash.getRange(4, 1, flows.length, header[0].length).setValues(flows);

  cash.getRange('A4:C9').setNumberFormat('yyyy-mm-dd');
  cash.getRange('D4:D9').setNumberFormat('#,##0.00');
  cash.getRange('E4:F9').setNumberFormat('0.000%');
  cash.getRange('G4:G9').setNumberFormat('#,##0.00');

  cash.setFrozenRows(3);
  cash.autoResizeColumns(1, header[0].length);

  Logger.log('Created: ' + ss.getUrl());
}
