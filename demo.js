/* Builds a small fictional workbook in the Setpoint tape layout (Report / Guidelines / Waivers / Overrides / Guarantors). */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.SetpointDemo = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';
  const ser = (y, m, d) => Math.round((Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 864e5);

  function build(XLSX) {
    const cols = ['Property Code', 'Parent/Cross', 'Status', 'Address', 'City', 'State', 'Zipcode', 'Exceptions',
      'Calculated FICO Score', 'Client Calculated FICO Score', 'Credit Report Aged',
      'Holdback Amount ($)', 'Client Holdback Amount ($)', 'Interest Reserve ($)', 'Note Amount ($)', 'Client Note Amount ($)',
      'Origination Date', 'Client Origination Date', 'Origination Date Variance', 'First Payment Date', 'Client First Payment Date',
      'Maturity Date', 'Client Maturity Date', 'Maturity Date Variance', 'Loan Term',
      'Gross Purchase Price ($)', 'Client Gross Purchase Price ($)', 'Purchase Date', 'Client Purchase Date', 'Purchase Date Variance',
      'Valuation', 'Valuation Type', 'Number of Units', 'Client Number of Units',
      'As-Is Appraised Value ($)', 'Client As-Is Appraised Value ($)', 'As-Repaired Appraised Value ($)', 'Client As-Repaired Appraised Value ($)',
      'Appraisal Date', 'Appraisers Cert Is Attached', 'Executed', 'Appraiser License Attached', 'Appraiser E/O Insurance Attached',
      'Repair Budget Amount ($)', 'Cost Incurred to Date ($)', 'Debt Service', '3rd Party Fee ($)', 'Closing Fees', 'Loan Type',
      'Policy Expiration Date', 'Flood Insurance Policy', 'In Flood Zone', 'Flood Zone Code'];
    const base = {
      'Parent/Cross': 'Parent', Status: 'Tied Out', City: 'Springfield', State: 'UT',
      'Calculated FICO Score': 735, 'Client Calculated FICO Score': 735, 'Credit Report Aged': 'N',
      'Holdback Amount ($)': 1200000, 'Client Holdback Amount ($)': 1200000, 'Interest Reserve ($)': 90000,
      'Note Amount ($)': 2400000, 'Client Note Amount ($)': 2400000,
      'Origination Date': ser(2026, 2, 27), 'Client Origination Date': ser(2026, 2, 27), 'Origination Date Variance': 0,
      'First Payment Date': ser(2026, 4, 1), 'Client First Payment Date': ser(2026, 4, 1),
      'Maturity Date': ser(2027, 9, 1), 'Client Maturity Date': ser(2027, 9, 1), 'Maturity Date Variance': 0, 'Loan Term': 18,
      'Gross Purchase Price ($)': 500000, 'Client Gross Purchase Price ($)': 500000,
      'Purchase Date': ser(2022, 8, 4), 'Client Purchase Date': ser(2022, 8, 4), 'Purchase Date Variance': 0,
      Valuation: 'Y', 'Valuation Type': '1004', 'Number of Units': 1, 'Client Number of Units': 1,
      'As-Is Appraised Value ($)': 1500000, 'Client As-Is Appraised Value ($)': 1500000,
      'As-Repaired Appraised Value ($)': 3800000, 'Client As-Repaired Appraised Value ($)': 3800000,
      'Appraisal Date': ser(2026, 2, 11), 'Appraisers Cert Is Attached': 'Y', Executed: 'Y', 'Appraiser License Attached': 'Y', 'Appraiser E/O Insurance Attached': 'Y',
      'Repair Budget Amount ($)': 2300000, 'Cost Incurred to Date ($)': 700000, 'Debt Service': 0, '3rd Party Fee ($)': 0, 'Closing Fees': 0,
      'Policy Expiration Date': ser(2026, 9, 10), 'Flood Insurance Policy': 'N', 'In Flood Zone': 'N', 'Flood Zone Code': 'X',
    };
    const loanA = Object.assign({}, base, { 'Property Code': 'DEMO-001', Address: '100 Example St', Zipcode: 84000, Exceptions: 'Flood Insurance Policy - Missing', 'Loan Type': 'Construction' });
    const loanB = Object.assign({}, base, { 'Property Code': 'DEMO-002', Address: '102 Example St', Zipcode: 84001, Exceptions: '', 'Loan Type': 'SFR 1-4',
      'Note Amount ($)': 2650000, 'Client Note Amount ($)': 2650000, 'Maturity Date Variance': 1, 'Loan Term': 19 });
    const report = [['Demo loan tape']].concat([cols], [loanA, loanB].map((o) => cols.map((c) => (o[c] === undefined ? null : o[c]))));

    const guidelines = [
      [], [],
      ['Max Term', 24], ['FICO', 625], ['Max Loan Amount ', 2500000], ['Max Loan Amount (Construction)', 15000000], [],
      [null, 'SFR 1-4', 'Commercial', 'Construction'],
      ['LTAIV', 0.85, 0.85, 0.85], ['LTARV', 0.75, 0.75, 0.75], ['Initial LTC', 0.8, 0.9, 0.9],
      ['Land Initial LTC', 0.7, 0.6], ['Total LTC', 0.9, 0.9, 0.9],
    ];
    const waivers = [['Property Code', 'Guideline', 'Waived At', 'Comment'], ['DEMO-002', 'Initial LTC', '2026-04-30 10:54:47 PM', 'Per lender: exception given for LTC']];
    const overrides = [['Property Code', 'Override Field', 'Override Value', 'Comment']];
    const guarantors = [['Asset', 'Name', 'Guarantor Type', 'Equifax', 'Experian', 'TransUnion', 'Report Date'],
      ['DEMO-001', 'Pat Example', 'Primary', 780, 771, 790, '08/02/2025'], ['DEMO-002', 'Pat Example', 'Primary', 780, 771, 790, '08/02/2025']];

    const wb = XLSX.utils.book_new();
    [['Report', report], ['Guidelines', guidelines], ['Waivers', waivers], ['Overrides', overrides], ['Guarantors', guarantors]].forEach(([n, rows]) => {
      const ws = XLSX.utils.aoa_to_sheet(rows.map((r) => (Array.isArray(r) ? r : cols.map((c) => r[c]))));
      XLSX.utils.book_append_sheet(wb, ws, n);
    });
    return wb;
  }
  return { build };
});
