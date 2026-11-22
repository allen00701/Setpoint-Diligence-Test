/*
 * Turns an analysis (from engine.js) into a shareable, self-contained HTML report,
 * an email draft, and a sources CSV. No dependencies.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.SetpointReport = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const esc = (s) => String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const money = (v) => (typeof v === 'number' ? '$' + Math.round(v).toLocaleString('en-US') : esc(v));
  const pct = (v) => (typeof v === 'number' ? (v * 100).toFixed(1) + '%' : esc(v));
  const fmtDate = (s) => { const d = new Date(Date.UTC(1899, 11, 30) + s * 864e5); return (d.getUTCMonth() + 1) + '/' + d.getUTCDate() + '/' + d.getUTCFullYear(); };

  const CSS = `
.sp-report{font:14px/1.5 -apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;color:#1b1f24;max-width:1000px;margin:0 auto}
.sp-report h1{font-size:24px;margin:0 0 4px}.sp-report h2{font-size:18px;margin:28px 0 8px;padding-bottom:4px;border-bottom:2px solid #1f3864}
.sp-report h3{font-size:15px;margin:18px 0 6px}.sp-meta{color:#57606a;font-size:13px;margin-bottom:14px}
.sp-report table{border-collapse:collapse;width:100%;margin:6px 0 12px;font-size:13px}
.sp-report th{background:#1f3864;color:#fff;text-align:left;padding:6px 8px;font-weight:600}
.sp-report td{border-bottom:1px solid #d8dee4;padding:6px 8px;vertical-align:top}
.sp-report td.num{text-align:right;white-space:nowrap;font-variant-numeric:tabular-nums}
.sp-grade{display:inline-block;min-width:26px;text-align:center;font-weight:700;border-radius:6px;padding:2px 8px;color:#fff}
.sp-grade.A{background:#1a7f37}.sp-grade.B{background:#9a6700}.sp-grade.C{background:#bc4c00}.sp-grade.D{background:#cf222e}.sp-grade.Cross{background:#6e7781}
.sp-pill{display:inline-block;border-radius:10px;padding:1px 9px;font-size:12px;font-weight:600}
.sp-pill.PASS{background:#dafbe1;color:#116329}.sp-pill.FAIL{background:#ffebe9;color:#a40e26}.sp-pill.PENDING{background:#fff8c5;color:#7d4e00}.sp-pill.Waived{background:#ddf4ff;color:#0550ae}
.sp-src{display:inline-block;font:11px/1.4 ui-monospace,SFMono-Regular,Menlo,Consolas,monospace;background:#eef1f4;border:1px solid #d0d7de;border-radius:4px;padding:0 5px;margin:1px 2px 1px 0;color:#32383f;white-space:nowrap}
.sp-card{border:1px solid #d0d7de;border-radius:8px;padding:12px 16px;margin:10px 0;background:#fff}
.sp-sev{font-weight:700;text-transform:uppercase;font-size:11px;letter-spacing:.04em;margin-right:6px}
.sp-sev.issue{color:#cf222e}.sp-sev.watch{color:#9a6700}.sp-sev.info{color:#0969da}
.sp-callout{background:#f6f8fa;border-left:4px solid #1f3864;padding:10px 14px;margin:10px 0;border-radius:0 6px 6px 0}
.sp-warn{background:#fff8c5;border-left:4px solid #bf8700;padding:8px 12px;margin:8px 0}
.sp-report pre{white-space:pre-wrap;background:#f6f8fa;border:1px solid #d0d7de;border-radius:6px;padding:12px;font:13px/1.5 inherit}
.sp-small{font-size:12px;color:#57606a}.sp-report ul{margin:4px 0 8px 18px;padding:0}
@media print{.sp-report{max-width:none}.sp-card,.sp-report table{page-break-inside:avoid}}
`;

  const grade = (g) => '<span class="sp-grade ' + esc(g) + '">' + esc(g) + '</span>';
  const pill = (r) => '<span class="sp-pill ' + esc(r) + '">' + esc(r) + '</span>';
  const refs = (list) => (list || []).filter(Boolean).filter((x, i, a) => a.indexOf(x) === i).map((r) => '<span class="sp-src">' + esc(r) + '</span>').join(' ');

  function showCheckVal(c, v) {
    if (v === 'PENDING' || v == null) return 'PENDING';
    if (c.name === 'Loan Amount') return money(v);
    if (c.name === 'FICO' || c.name === 'Term') return esc(v);
    return pct(v);
  }

  function buildEmail(a, opts) {
    const L = [];
    const loans = a.loans.filter((l) => !l.cross);
    const gradeList = loans.map((l) => l.code + ' is a ' + l.grades.overall).join(' and ');
    L.push('Hi ' + (opts.clientName || '[Client name]') + ',');
    L.push('');
    L.push('Thank you for the data tape. We completed the review of ' + loans.map((l) => l.code).join(' and ') + '. ' +
      (a.meetsExpectation ? 'Both loans meet your expected overall grade of ' + a.expected + '.' : 'Not every loan reaches the overall grade of ' + a.expected + ' you expected: ' + gradeList + '.'));
    loans.forEach((l) => {
      L.push('');
      L.push(l.code + ' - Overall ' + l.grades.overall + ' (Credit ' + l.grades.credit + ', Valuation ' + l.grades.valuation + ')');
      if (!l.pathToA.length) L.push('  All guideline checks pass, no documents are missing and the valuation is complete.');
      l.pathToA.forEach((p) => L.push('  - ' + p.text));
      const issues = l.observations.filter((o) => o.sev !== 'info');
      if (issues.length) { L.push('  Items to confirm:'); issues.forEach((o) => L.push('  - ' + o.title + '.')); }
    });
    L.push('');
    L.push('Sources for every figure above are listed in the attached report. Please let us know if you would like anything reviewed differently.');
    L.push('');
    L.push('Best regards,');
    L.push(opts.preparedBy || '[Your name]');
    return L.join('\n');
  }

  function sourcesRows(a) {
    const rows = [];
    a.loans.forEach((l) => Object.values(l.sources || {}).forEach((s) => rows.push({ loan: l.code, ref: s.ref, field: s.label, value: s.value })));
    return rows;
  }
  function sourcesCsv(a) {
    const q = (x) => '"' + String(x == null ? '' : x).replace(/"/g, '""') + '"';
    return ['Loan,Source cell,Field,Value'].concat(sourcesRows(a).map((r) => [r.loan, r.ref, r.field, r.value].map(q).join(','))).join('\n');
  }

  function render(a, opts) {
    opts = opts || {};
    const loans = a.loans;
    const rankable = loans.filter((l) => !l.cross);
    const h = [];
    h.push('<div class="sp-report">');
    h.push('<h1>' + esc(opts.title || 'Loan Tape Review Report') + '</h1>');
    h.push('<div class="sp-meta">' + (opts.fileName ? 'Source file: <b>' + esc(opts.fileName) + '</b> &middot; ' : '') + 'Generated ' + esc(new Date(a.generated).toLocaleString()) +
      (opts.preparedBy ? ' &middot; Prepared by ' + esc(opts.preparedBy) : '') + ' &middot; Expected overall grade: <b>' + esc(a.expected) + '</b></div>');
    a.warnings.forEach((w) => h.push('<div class="sp-warn">' + esc(w) + '</div>'));

    /* summary */
    h.push('<h2>1. Summary</h2>');
    h.push('<div class="sp-callout">' + (a.meetsExpectation ? 'All loans meet the expected overall grade of <b>' + esc(a.expected) + '</b>.'
      : 'Not all loans reach the expected overall grade of <b>' + esc(a.expected) + '</b>: ' + rankable.map((l) => esc(l.code) + ' = <b>' + esc(l.grades.overall) + '</b>').join(', ') + '.') + '</div>');
    h.push('<table><tr><th>Loan</th><th>Address</th><th>Bucket</th><th>Credit</th><th>Valuation</th><th>Overall</th><th>Missing docs</th><th>Fails</th><th>Waivers</th></tr>');
    loans.forEach((l) => {
      h.push('<tr><td>' + esc(l.code) + '</td><td>' + esc(l.address) + '</td><td>' + esc(l.bucket || '-') + '</td><td>' + grade(l.grades.credit) + '</td><td>' + grade(l.grades.valuation) + '</td><td>' + grade(l.grades.overall) + '</td>' +
        '<td class="num">' + (l.counts ? l.counts.missing : '-') + '</td><td class="num">' + (l.counts ? l.counts.fails : '-') + '</td><td class="num">' + (l.counts ? l.counts.waived : '-') + '</td></tr>');
    });
    h.push('</table>');

    /* per loan */
    loans.forEach((l, i) => {
      h.push('<h2>' + (i + 2) + '. ' + esc(l.code) + (l.address ? ' - ' + esc(l.address) + (l.city ? ', ' + esc(l.city) : '') + (l.state ? ', ' + esc(l.state) : '') : '') + '</h2>');
      if (l.cross) { h.push('<p>' + l.reasons[0] + '</p>'); return; }
      h.push('<p>Overall grade ' + grade(l.grades.overall) + ' &nbsp; Credit ' + grade(l.grades.credit) + ' &nbsp; Valuation ' + grade(l.grades.valuation) + '</p>');

      h.push('<h3>How the grade was reached</h3><ul>' + l.reasons.map((r) => '<li>' + esc(r) + '</li>').join('') + '</ul>');
      h.push('<h3>What it would take to reach an A</h3>');
      if (!l.pathToA.length) h.push('<p>Nothing: the loan already meets every guideline and has no missing documents.</p>');
      else h.push('<ul>' + l.pathToA.map((p) => '<li>' + esc(p.text) + ' ' + refs(p.refs) + '</li>').join('') + '</ul>');

      h.push('<h3>Calculated metrics</h3><table><tr><th>Metric</th><th>Value</th><th>Calculation</th><th>Sources</th></tr>');
      Object.values(l.metrics).forEach((m) => {
        h.push('<tr><td>' + esc(m.label) + '</td><td class="num">' + esc(m.display) + (m.overridden ? ' <span class="sp-pill Waived">override</span>' : '') + '</td><td>' + esc(m.formula) + '</td><td>' + refs(m.refs.concat(m.overridden ? [m.overridden.ref] : [])) + '</td></tr>');
      });
      h.push('</table>');

      h.push('<h3>Guideline checks</h3><table><tr><th>Guideline</th><th>Loan value</th><th>Limit</th><th>Result</th><th>Rule</th><th>Sources</th></tr>');
      l.checks.forEach((c) => {
        h.push('<tr><td>' + esc(c.name) + '</td><td class="num">' + showCheckVal(c, c.value) + '</td><td class="num">' + (c.min ? '&ge; ' : '&le; ') + showCheckVal(c, c.limit) + '</td><td>' + pill(c.result) + '</td><td>' + esc(c.rule) + '</td><td>' + refs([].concat(c.valueRefs, c.limitRef, c.waiver)) + '</td></tr>');
      });
      h.push('</table>');
      l.waiversUsed.forEach((w) => h.push('<div class="sp-card"><b>Waiver applied: ' + esc(w.guideline) + '</b> ' + refs([w.ref]) + '<br>' + esc(w.comment || '(no comment)') + '</div>'));

      h.push('<h3>Missing documents</h3>');
      if (!l.missingDocs.count) h.push('<p>None flagged in the Exceptions column.</p>');
      else h.push('<ul>' + l.missingDocs.items.map((m) => '<li>' + esc(m.text) + ' ' + refs([m.ref]) + '</li>').join('') + '</ul>');

      h.push('<h3>Valuation</h3><p>Type: <b>' + esc(l.valuation.type || '-') + '</b> &middot; score ' + l.valuation.score + (l.valuation.deficiencies.length ? '</p><ul>' + l.valuation.deficiencies.map((d) => '<li>' + esc(d.t) + ' ' + refs([d.ref]) + '</li>').join('') + '</ul>' : ' (no deficiencies)</p>'));

      h.push('<h3>Findings and items to confirm</h3>');
      h.push('<p class="sp-small">Client vs tape tie-out: ' + l.tieOut.tied + ' of ' + l.tieOut.compared + ' client-supplied fields agree with the tape.</p>');
      if (!l.observations.length) h.push('<p>No data observations.</p>');
      l.observations.forEach((o) => h.push('<div class="sp-card"><span class="sp-sev ' + esc(o.sev) + '">' + esc(o.sev) + '</span><b>' + esc(o.title) + '</b><br>' + esc(o.detail) + '<br>' + refs(o.refs) + '</div>'));
    });

    /* email */
    h.push('<h2>' + (loans.length + 2) + '. Draft response to the client</h2><pre>' + esc(buildEmail(a, opts)) + '</pre>');

    /* sources */
    const rows = sourcesRows(a);
    h.push('<h2>' + (loans.length + 3) + '. Source index</h2><p class="sp-small">Every cell cited above. Cell references are to the uploaded workbook.</p><table><tr><th>Loan</th><th>Cell</th><th>Field</th><th>Value</th></tr>');
    rows.forEach((r) => h.push('<tr><td>' + esc(r.loan) + '</td><td>' + refs([r.ref]) + '</td><td>' + esc(r.field) + '</td><td>' + (typeof r.value === 'number' && /date/i.test(r.field) ? fmtDate(r.value) : esc(r.value)) + '</td></tr>'));
    h.push('</table>');
    h.push('<h2>Method and assumptions</h2><ul class="sp-small">' +
      '<li>Metrics and grades are recalculated from the Report, Guidelines, Waivers, Overrides and Guarantors tabs using the same rules as the template Summary tab. Nothing is read from cached Summary values except for the cross-check.</li>' +
      '<li>Ratios are compared with the guideline using "less than or equal"; FICO uses "greater than or equal".</li>' +
      '<li>A credit report is treated as aged after ' + a.agedCreditDays + ' days before origination (an assumption; confirm against the client\'s rule).</li>' +
      '<li>All processing happened in the browser; the workbook was not uploaded anywhere.</li></ul>');
    h.push('</div>');
    return h.join('\n');
  }

  function standalone(a, opts) {
    return '<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>' + esc((opts && opts.title) || 'Loan Tape Review Report') + '</title><style>body{margin:24px;background:#fff}' + CSS + '</style></head><body>' + render(a, opts) + '</body></html>';
  }

  return { render, standalone, buildEmail, sourcesCsv, CSS, esc };
});
