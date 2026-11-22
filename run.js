// Run with: node test/run.js [optional/path/to/workbook.xlsx]
const assert = require('assert');
const fs = require('fs');
const XLSX = require('../vendor/xlsx.full.min.js');
const Engine = require('../js/engine.js');
const Report = require('../js/report.js');
const Demo = require('../js/demo.js');

let n = 0;
const t = (name, fn) => { fn(); n++; console.log('ok -', name); };

t('DAYS360 matches Excel (US method)', () => {
  const s = (y, m, d) => Math.round((Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 864e5);
  assert.strictEqual(Engine.days360(s(2022, 8, 4), s(2026, 2, 27)), 1283);
  assert.strictEqual(Engine.days360(s(2026, 4, 1), s(2027, 9, 1)), 510);
  assert.strictEqual(Engine.days360(s(2024, 1, 31), s(2024, 3, 31)), 60);
});

const demo = Engine.analyze(XLSX, Demo.build(XLSX), { today: '2026-11-21' });
const [a, b] = demo.loans;

t('demo: construction loan -> bucket, checks, one missing doc, grade B', () => {
  assert.strictEqual(a.bucket, 'Construction');
  assert.ok(a.checks.every((c) => c.result === 'PASS'));
  assert.strictEqual(a.missingDocs.count, 1);
  assert.deepStrictEqual([a.grades.credit, a.grades.valuation, a.grades.overall], ['B', 'A', 'B']);
});
t('demo: SFR loan -> loan amount fails, LTC waived, grade C', () => {
  assert.strictEqual(b.bucket, 'SFR 1-4');
  const r = Object.fromEntries(b.checks.map((c) => [c.name, c.result]));
  assert.strictEqual(r['Loan Amount'], 'FAIL');
  assert.strictEqual(r['Initial LTC'], 'Waived');
  assert.strictEqual(b.grades.overall, 'C');
});
t('demo: maturity variance and flood observations are raised with sources', () => {
  assert.ok(b.observations.some((o) => /Maturity Date Variance/.test(o.title) && o.refs.length));
  assert.ok(a.observations.some((o) => /Flood policy/.test(o.title)));
});
t('missing-doc counting is case-insensitive', () => assert.strictEqual(a.missingDocs.count, 1));
t('every check cites sources', () => a.checks.forEach((c) => assert.ok(c.valueRefs.length && c.limitRef, c.name)));

t('report renders, escapes HTML and includes sources + email', () => {
  const wb = Demo.build(XLSX);
  wb.Sheets.Report.D3.v = '<script>alert(1)</script>';
  const an = Engine.analyze(XLSX, wb, { today: '2026-11-21' });
  const html = Report.standalone(an, { title: 'T <b>', preparedBy: 'Me' });
  assert.ok(!html.includes('<script>alert(1)'));
  assert.ok(html.includes('Source index') && html.includes('Report!') && html.includes('Draft response'));
  assert.ok(Report.sourcesCsv(an).split('\n').length > 20);
});

t('expectation flag works', () => {
  assert.strictEqual(demo.meetsExpectation, false);
  assert.strictEqual(Engine.analyze(XLSX, Demo.build(XLSX), { expectedGrade: 'B', today: '2026-11-21' }).meetsExpectation, false);
});

t('clear error when the Report tab is missing', () => {
  const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([[1]]), 'Other');
  assert.throws(() => Engine.analyze(XLSX, wb), /Report/);
});

const sample = process.argv[2];
if (sample && fs.existsSync(sample)) {
  const wb = XLSX.read(fs.readFileSync(sample), { type: 'buffer', cellFormula: true });
  const an = Engine.analyze(XLSX, wb, { today: '2026-11-21' });
  an.loans.forEach((l) => console.log(' ', l.code, l.bucket, JSON.stringify(l.grades), JSON.stringify(l.counts)));
  t('sample workbook analysed (' + sample + ')', () => assert.ok(an.loans.length >= 1));
}
console.log('\n' + n + ' tests passed');
