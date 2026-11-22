/*
 * Setpoint loan-tape review engine.
 * Re-computes the Summary-tab logic (metrics, guideline checks, missing docs,
 * grades) from the raw Report / Guidelines / Waivers / Overrides tabs and keeps
 * track of the source cell behind every number, so a report can cite its sources.
 *
 * Works in the browser (global SetpointEngine) and in Node (require).
 * The SheetJS library is passed in, so this file has no dependencies.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.SetpointEngine = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  /* ---------- small helpers ---------- */
  const EPOCH = Date.UTC(1899, 11, 30);
  const norm = (s) => String(s == null ? '' : s).toLowerCase().replace(/[^a-z0-9]/g, '');
  const trim = (s) => String(s == null ? '' : s).trim();
  const fromSerial = (s) => {
    const d = new Date(EPOCH + Math.round(s) * 864e5);
    return { y: d.getUTCFullYear(), m: d.getUTCMonth() + 1, d: d.getUTCDate() };
  };
  const toSerial = (y, m, d) => Math.round((Date.UTC(y, m - 1, d) - EPOCH) / 864e5);
  const lastDayOfFeb = (y, m, d) => m === 2 && d === new Date(Date.UTC(y, 2, 0)).getUTCDate();

  // Excel DAYS360, US (NASD) method - the default the template uses.
  function days360(a, b) {
    const s = fromSerial(a), e = fromSerial(b);
    let d1 = s.d, d2 = e.d;
    const sLast = lastDayOfFeb(s.y, s.m, s.d), eLast = lastDayOfFeb(e.y, e.m, e.d);
    if (sLast && eLast) d2 = 30;
    if (sLast) d1 = 30;
    if (d2 === 31 && d1 >= 30) d2 = 30;
    if (d1 === 31) d1 = 30;
    return (e.y - s.y) * 360 + (e.m - s.m) * 30 + (d2 - d1);
  }

  function num(v) {
    if (typeof v === 'number') return isFinite(v) ? v : null;
    if (typeof v === 'string') {
      const t = v.replace(/[$,%\s]/g, '');
      if (t !== '' && !isNaN(Number(t))) return Number(t);
    }
    return null;
  }
  const n0 = (v) => num(v) || 0; // blank cells behave like 0 (as INDEX does in Excel)

  function dateSerial(v) {
    if (typeof v === 'number') return v > 1 ? v : null;
    if (typeof v === 'string') {
      let m = v.match(/^(\d{4})-(\d{2})-(\d{2})/);
      if (m) return toSerial(+m[1], +m[2], +m[3]);
      m = v.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
      if (m) return toSerial(+m[3], +m[1], +m[2]);
    }
    return null;
  }
  const fmtDate = (s) => { if (s == null) return ''; const o = fromSerial(s); return o.m + '/' + o.d + '/' + o.y; };
  const fmtMoney = (v) => (typeof v === 'number' ? '$' + Math.round(v).toLocaleString('en-US') : String(v));
  const fmtPct = (v) => (typeof v === 'number' ? (v * 100).toFixed(1) + '%' : String(v));

  // Value of a SheetJS cell. If a formula has no cached value (workbook saved on
  // manual calc), fall back to parsing =VALUE("yyyy-mm-dd") date formulas.
  function cellValue(cell) {
    if (!cell || cell.t === 'e') return null;
    if (cell.v !== undefined && cell.v !== '') return cell.v;
    if (cell.f) {
      const m = String(cell.f).match(/VALUE\("(\d{4})-(\d{2})-(\d{2})"\)/i);
      if (m) return toSerial(+m[1], +m[2], +m[3]);
    }
    return null;
  }

  const SUMMARY_FIELDS = {
    purchasePrice: 'Purchase Price', days: 'Days Since Purchase', asIs: 'As is Value',
    arv: 'As Repaired Value', holdback: 'Holdback', initialCost: 'Initial Cost',
    costBasis: 'Cost Basis', iltc: 'ILTC', tltc: 'Total LTC', ltaiv: 'LTAIV', ltarv: 'LTARV',
    units: 'Units', loanType: 'Loan Type', bucket: 'Bucket', term: 'Term',
  };
  const DEFAULT_BUCKETS = ['SFR 1-4', 'Commercial', 'Construction'];
  const APPRAISAL_LIST = ['bpo', 'hybridappraisal', 'desktopappraisal', 'other'];
  const AGED_CREDIT_DAYS = 120;

  /* ---------- main entry ---------- */
  function analyze(XLSX, wb, opts) {
    opts = opts || {};
    const now = new Date();
    const today = opts.today ? (dateSerial(opts.today) || opts.today) : toSerial(now.getFullYear(), now.getMonth() + 1, now.getDate());
    const expected = (opts.expectedGrade || 'A').toUpperCase();
    const warnings = [];
    const U = XLSX.utils;

    const sheetByName = (name) => {
      const k = Object.keys(wb.Sheets).find((n) => norm(n) === norm(name));
      return k ? wb.Sheets[k] : null;
    };
    const rangeOf = (ws) => (ws && ws['!ref'] ? U.decode_range(ws['!ref']) : { s: { r: 0, c: 0 }, e: { r: -1, c: -1 } });
    const at = (ws, c, r) => (ws ? ws[U.encode_cell({ c, r: r - 1 })] : undefined); // r is 1-based

    /* --- Report tab: header row 2, data from row 3 --- */
    const reportWs = sheetByName('Report');
    if (!reportWs) throw new Error('No "Report" tab found. Upload the Setpoint loan-tape workbook.');
    const rr = rangeOf(reportWs);
    const headers = {};
    for (let c = 0; c <= rr.e.c; c++) {
      const h = trim(cellValue(at(reportWs, c, 2)));
      if (h) (headers[h] = headers[h] || []).push(c);
    }
    const missingFields = new Set();
    const get = (field, row, occ) => {
      const cols = headers[field];
      if (!cols || cols[occ || 0] === undefined) { missingFields.add(field); return { v: null, ref: null, field }; }
      const c = cols[occ || 0];
      return { v: cellValue(at(reportWs, c, row)), ref: 'Report!' + U.encode_col(c) + row, field };
    };

    const byCode = {};
    const order = [];
    for (let r = 3; r <= rr.e.r + 1; r++) {
      const code = trim(get('Property Code', r).v);
      if (!code) continue;
      if (!byCode[code]) { byCode[code] = []; order.push(code); }
      byCode[code].push(r);
    }
    if (!order.length) throw new Error('No property codes found in the Report tab (expected from row 3, column A).');

    /* --- Guidelines tab --- */
    const gWs = sheetByName('Guidelines');
    if (!gWs) throw new Error('No "Guidelines" tab found.');
    const gr = rangeOf(gWs);
    const labels = {};
    let bucketRow = null, buckets = DEFAULT_BUCKETS.slice();
    for (let r = 1; r <= gr.e.r + 1; r++) {
      const lab = trim(cellValue(at(gWs, 0, r)));
      if (lab && !labels[norm(lab)]) labels[norm(lab)] = r;
      if (!bucketRow) {
        const b = [1, 2, 3].map((c) => trim(cellValue(at(gWs, c, r))));
        if (/sfr/i.test(b[0]) && b[1]) { bucketRow = r; buckets = b; }
      }
    }
    const gcell = (label, col) => {
      const r = labels[norm(label)];
      if (!r) { warnings.push('Guidelines tab: row "' + label + '" not found.'); return { v: null, ref: null }; }
      return { v: num(cellValue(at(gWs, col, r))), ref: 'Guidelines!' + U.encode_col(col) + r };
    };
    const guide = {
      maxTerm: gcell('Max Term', 1), fico: gcell('FICO', 1),
      maxLoan: gcell('Max Loan Amount', 1), maxLoanCon: gcell('Max Loan Amount (Construction)', 1),
      buckets,
      ratio: (label, bucket) => {
        const i = buckets.findIndex((b) => norm(b) === norm(bucket));
        return i < 0 ? { v: null, ref: null } : gcell(label, i + 1);
      },
    };

    /* --- Waivers & Overrides --- */
    const readTable = (name, cols) => {
      const ws = sheetByName(name), out = [];
      if (!ws) return out;
      const r = rangeOf(ws);
      for (let row = 2; row <= r.e.r + 1; row++) {
        const o = { row };
        cols.forEach((k, c) => { o[k] = cellValue(at(ws, c, row)); });
        if (trim(o.code)) { o.code = trim(o.code); o.sheet = name; out.push(o); }
      }
      return out;
    };
    const waivers = readTable('Waivers', ['code', 'guideline', 'at', 'comment']);
    const overrides = readTable('Overrides', ['code', 'field', 'value', 'comment']);

    /* --- Guarantors (optional cross-check) --- */
    const guarWs = sheetByName('Guarantors');
    const guarantors = [];
    if (guarWs) {
      const r = rangeOf(guarWs);
      for (let row = 2; row <= r.e.r + 1; row++) {
        const code = trim(cellValue(at(guarWs, 0, row)));
        if (!code) continue;
        const sc = [3, 4, 5].map((c) => num(cellValue(at(guarWs, c, row)))).filter((x) => x != null).sort((a, b) => a - b);
        guarantors.push({ code, name: trim(cellValue(at(guarWs, 1, row))), type: trim(cellValue(at(guarWs, 2, row))),
          scores: sc, median: sc.length === 3 ? sc[1] : null, reportDate: dateSerial(cellValue(at(guarWs, 6, row))),
          row, ref: 'Guarantors!A' + row });
      }
    }

    /* --- Summary tab cached grades (optional cross-check) --- */
    const summaryGrades = {};
    const sumWs = sheetByName('Summary');
    if (sumWs) {
      const sr = rangeOf(sumWs);
      let hdr = null;
      for (let r = 1; r <= Math.min(30, sr.e.r + 1) && !hdr; r++) {
        for (let c = 0; c <= sr.e.c; c++) {
          if (norm(cellValue(at(sumWs, c, r))) === 'overallgrade') { hdr = { r, overall: c }; break; }
        }
      }
      if (hdr) {
        const find = (name) => { for (let c = 0; c <= sr.e.c; c++) if (norm(cellValue(at(sumWs, c, hdr.r))) === norm(name)) return c; return -1; };
        const pc = find('Property Code'), cc = find('Credit Grade'), vc = find('Valuation Grade');
        for (let r = hdr.r + 1; r <= sr.e.r + 1; r++) {
          const code = trim(cellValue(at(sumWs, pc, r)));
          if (!code || code === '0') continue;
          summaryGrades[code] = {
            credit: trim(cellValue(at(sumWs, cc, r))), valuation: trim(cellValue(at(sumWs, vc, r))),
            overall: trim(cellValue(at(sumWs, hdr.overall, r))), ref: 'Summary!' + U.encode_col(hdr.overall) + r,
          };
        }
      }
    }

    /* ---------- per-loan analysis ---------- */
    const loans = order.map((code) => analyzeLoan(code));

    function analyzeLoan(code) {
      const rows = byCode[code];
      const r0 = rows[0];
      const g = (f, occ) => get(f, r0, occ);
      const sumOf = (f) => {
        const parts = rows.map((r) => get(f, r));
        return { v: parts.reduce((a, p) => a + n0(p.v), 0), refs: parts.map((p) => p.ref).filter(Boolean) };
      };
      const loan = { code, rows, metrics: {}, checks: [], observations: [], overridesUsed: [], waiversUsed: [], sources: {} };
      const src = (ref, label, value) => { if (ref) loan.sources[ref] = { ref, label, value }; return ref; };
      const S = (p, label) => { src(p.ref, label || p.field, p.v); return p.ref; };
      const obs = (sev, title, detail, refs) => loan.observations.push({ sev, title, detail, refs: (refs || []).filter(Boolean) });

      loan.address = trim(g('Address').v);
      loan.city = trim(g('City').v); loan.state = trim(g('State').v);
      loan.parentCross = trim(g('Parent/Cross').v);
      loan.cross = /^c/i.test(loan.parentCross);
      loan.status = trim(g('Status').v);
      if (loan.cross) {
        loan.grades = { credit: loan.parentCross, valuation: loan.parentCross, overall: 'Cross' };
        loan.reasons = ['Cross-collateralised child record: graded with its parent loan, not on its own.'];
        loan.pathToA = [];
        return loan;
      }

      const getOverride = (key) => {
        const o = overrides.find((x) => x.code === code && norm(x.field) === norm(SUMMARY_FIELDS[key]));
        return o || null;
      };
      // Build a metric: base calculation, then template-style override, then PENDING on any failure.
      const metric = (key, label, calc, formula, refs, fmt) => {
        let value;
        try { value = calc(); } catch (e) { value = 'PENDING'; }
        if (typeof value === 'number' && !isFinite(value)) value = 'PENDING';
        if (value == null || Number.isNaN(value)) value = 'PENDING';
        const ov = getOverride(key);
        let overridden = null;
        if (ov) {
          overridden = { base: value, ref: 'Overrides!C' + ov.row, comment: ov.comment };
          value = typeof ov.value === 'number' || num(ov.value) == null ? ov.value : num(ov.value);
          loan.overridesUsed.push({ field: SUMMARY_FIELDS[key], value, ref: overridden.ref, comment: trim(ov.comment), base: overridden.base });
          src(overridden.ref, 'Override: ' + SUMMARY_FIELDS[key], value);
        }
        const m = { key, label, value, formula, refs: (refs || []).filter(Boolean), overridden, fmt: fmt || 'num' };
        m.display = display(m.value, m.fmt);
        loan.metrics[key] = m;
        return value;
      };
      const display = (v, fmt) => (v === 'PENDING' || v == null ? 'PENDING'
        : fmt === 'money' ? fmtMoney(v) : fmt === 'pct' ? fmtPct(v) : fmt === 'date' ? fmtDate(v) : String(v));

      // inputs (with their source cells)
      const purchase = g('Gross Purchase Price ($)'), purchaseDate = g('Purchase Date'), origDate = g('Origination Date');
      const asIsIn = g('As-Is Appraised Value ($)'), arvIn = g('As-Repaired Appraised Value ($)');
      const unitsIn = g('Number of Units'), loanTypeIn = g('Loan Type');
      const note = sumOf('Note Amount ($)'), holdbackS = sumOf('Holdback Amount ($)'), irS = sumOf('Interest Reserve ($)');
      const repair = sumOf('Repair Budget Amount ($)'), debt = sumOf('Debt Service'), incurred = sumOf('Cost Incurred to Date ($)');
      const fee3 = sumOf('3rd Party Fee ($)'), closing = sumOf('Closing Fees');
      const firstPay = g('First Payment Date'), maturity = g('Maturity Date');
      [purchase, purchaseDate, origDate, asIsIn, arvIn, unitsIn, firstPay, maturity].forEach((p) => S(p));
      S(loanTypeIn, 'Loan Type (Additional Fields)');
      note.refs.forEach((r) => src(r, 'Note Amount ($)', note.v));

      metric('purchasePrice', 'Purchase Price', () => n0(purchase.v), 'Gross Purchase Price from the purchase agreement', [purchase.ref], 'money');
      metric('days', 'Days Since Purchase', () => {
        const a = dateSerial(purchaseDate.v), b = dateSerial(origDate.v);
        if (a == null || b == null) throw 0; return days360(a, b);
      }, 'DAYS360(Purchase Date, Origination Date)', [purchaseDate.ref, origDate.ref]);
      metric('asIs', 'As-Is Value', () => n0(asIsIn.v), 'As-Is Appraised Value', [asIsIn.ref], 'money');
      metric('arv', 'As-Repaired Value', () => n0(arvIn.v), 'As-Repaired Appraised Value', [arvIn.ref], 'money');
      metric('holdback', 'Holdback', () => holdbackS.v, 'Holdback Amount on the final HUD', holdbackS.refs, 'money');
      metric('units', 'Units', () => n0(unitsIn.v), 'Number of Units from the valuation', [unitsIn.ref]);
      const M = (k) => loan.metrics[k].value;
      const isNum = (v) => typeof v === 'number';

      metric('initialCost', 'Initial Cost', () => {
        const aged = M('days') === 'PENDING' || M('days') > 730 || M('purchasePrice') === 0;
        let v = aged ? M('asIs') : M('purchasePrice') + incurred.v;
        if (M('units') >= 5) v += fee3.v + closing.v;
        return v;
      }, 'If purchase is over 730 days old (or price is 0): As-Is Value; otherwise Purchase Price + Cost Incurred to Date. Add 3rd-party and closing fees when 5+ units.',
      [].concat(loan.metrics.asIs.refs, loan.metrics.purchasePrice.refs, incurred.refs), 'money');
      metric('costBasis', 'Cost Basis', () => M('initialCost') + repair.v + debt.v - incurred.v,
        'Initial Cost + Repair Budget + Debt Service - Cost Incurred to Date', [].concat(repair.refs, debt.refs, incurred.refs), 'money');
      repair.refs.forEach((r) => src(r, 'Repair Budget Amount ($)', repair.v));
      incurred.refs.forEach((r) => src(r, 'Cost Incurred to Date ($)', incurred.v));
      irS.refs.forEach((r) => src(r, 'Interest Reserve ($)', irS.v));
      holdbackS.refs.forEach((r) => src(r, 'Holdback Amount ($)', holdbackS.v));

      metric('iltc', 'Initial LTC', () => (note.v - M('holdback')) / M('initialCost'), '(Note Amount - Holdback) / Initial Cost', note.refs.concat(holdbackS.refs), 'pct');
      metric('tltc', 'Total LTC', () => note.v / M('costBasis'), 'Note Amount / Cost Basis', note.refs, 'pct');
      metric('ltaiv', 'LTAIV', () => (note.v - M('holdback') - irS.v) / M('asIs'), '(Note Amount - Holdback - Interest Reserve) / As-Is Value', note.refs.concat(holdbackS.refs, irS.refs), 'pct');
      metric('ltarv', 'LTARV', () => note.v / M('arv'), 'Note Amount / As-Repaired Value', note.refs, 'pct');

      // Bucket from Loan Type (blank -> PENDING; starts with C -> Construction; 5+ units -> Commercial; else SFR)
      metric('bucket', 'Bucket', () => {
        const lt = trim(loanTypeIn.v);
        if (!lt) throw 0;
        if (/^c/i.test(lt)) return buckets[2];
        return M('units') >= 5 ? buckets[1] : buckets[0];
      }, 'From the Loan Type on the Report tab: "C…" = Construction, otherwise Commercial if 5+ units, else SFR 1-4', [loanTypeIn.ref, unitsIn.ref]);
      loan.bucket = M('bucket');

      metric('term', 'Term (months)', () => {
        const a = dateSerial(firstPay.v), b = dateSerial(maturity.v);
        if (a == null || b == null) throw 0;
        return Math.round(days360(a, b) / 30 + 1);
      }, 'ROUND(DAYS360(First Payment Date, Maturity Date) / 30 + 1)', [firstPay.ref, maturity.ref]);

      /* --- guideline checks --- */
      const findWaiver = (name) => waivers.find((w) => w.code === code && norm(w.guideline) === norm(name));
      const check = (name, valueFn, limitFn, rule, valueRefs) => {
        const w = findWaiver(name);
        let result, value, limit = { v: null, ref: null };
        try { value = valueFn(); limit = limitFn(); } catch (e) { value = 'PENDING'; }
        if (w) { result = 'Waived'; src('Waivers!A' + w.row, 'Waiver: ' + name, trim(w.comment)); loan.waiversUsed.push({ guideline: name, ref: 'Waivers!A' + w.row, comment: trim(w.comment), at: w.at }); }
        else if (value === 'PENDING' || value == null || limit.v == null) result = 'PENDING';
        else result = value <= limit.v ? 'PASS' : 'FAIL';
        if (limit.ref) src(limit.ref, 'Guideline: ' + name, limit.v);
        loan.checks.push({ name, value, limit: limit.v, limitRef: limit.ref, result, rule, valueRefs: (valueRefs || []).filter(Boolean), waiver: w ? 'Waivers!A' + w.row : null });
        return result;
      };
      const fico = g('Calculated FICO Score'); S(fico, 'Calculated FICO Score');
      const ficoOK = () => { const f = num(fico.v); if (!f) throw 0; return f; };
      // FICO is a minimum (>=), so express it as "shortfall" to reuse the <= test.
      (function () {
        const w = findWaiver('FICO'); let result, f = null;
        try { f = ficoOK(); } catch (e) { /* pending */ }
        if (w) { result = 'Waived'; src('Waivers!A' + w.row, 'Waiver: FICO', trim(w.comment)); loan.waiversUsed.push({ guideline: 'FICO', ref: 'Waivers!A' + w.row, comment: trim(w.comment), at: w.at }); }
        else if (f == null || guide.fico.v == null) result = 'PENDING';
        else result = f >= guide.fico.v ? 'PASS' : 'FAIL';
        if (guide.fico.ref) src(guide.fico.ref, 'Guideline: FICO minimum', guide.fico.v);
        loan.checks.push({ name: 'FICO', value: f == null ? 'PENDING' : f, limit: guide.fico.v, limitRef: guide.fico.ref, result, rule: 'Calculated FICO must be at least the minimum', min: true, valueRefs: [fico.ref], waiver: w ? 'Waivers!A' + w.row : null });
      })();
      check('Loan Amount', () => { if (!note.v) throw 0; if (M('bucket') === 'PENDING') throw 0; return note.v; },
        () => (M('bucket') === buckets[2] ? guide.maxLoanCon : guide.maxLoan), 'Note Amount must not exceed the max loan amount for the bucket (Construction has its own max)', note.refs);
      check('Term', () => { const t = M('term'); if (t === 'PENDING' || t === 1) throw 0; return t; }, () => guide.maxTerm, 'Term must not exceed the max term', loan.metrics.term.refs);
      const ratioCheck = (name, key, label) => check(name, () => { const v = M(key); if (!isNum(v)) throw 0; return v; },
        () => { if (M('bucket') === 'PENDING') throw 0; return guide.ratio(label, M('bucket')); }, label + ' must not exceed the threshold for the bucket', loan.metrics[key].refs);
      ratioCheck('Initial LTC', 'iltc', 'Initial LTC');
      ratioCheck('Total LTC', 'tltc', 'Total LTC');
      ratioCheck('LTAIV', 'ltaiv', 'LTAIV');
      ratioCheck('LTARV', 'ltarv', 'LTARV');

      /* --- missing docs --- */
      const missing = [];
      rows.forEach((r) => {
        const ex = get('Exceptions', r);
        const text = trim(ex.v);
        const count = (text.match(/missing/gi) || []).length;
        if (count) { missing.push({ text, count, ref: ex.ref }); src(ex.ref, 'Exceptions', text); }
      });
      loan.missingDocs = { count: missing.reduce((a, m) => a + m.count, 0), items: missing };

      /* --- valuation score / grade --- */
      const vType = g('Valuation Type'), vDoc = g('Valuation'), appDate = g('Appraisal Date');
      const cert = g('Appraisers Cert Is Attached'), exec = g('Executed'), lic = g('Appraiser License Attached'), eo = g('Appraiser E/O Insurance Attached');
      [vType, vDoc, appDate, cert, exec, lic, eo].forEach((p) => S(p));
      const typeStr = trim(vType.v);
      const defs = [];
      if (trim(vDoc.v).toUpperCase() === 'N') defs.push({ t: 'Valuation document is missing', ref: vDoc.ref, pts: 3 });
      else if (!/^[bo]/i.test(typeStr)) {
        if (dateSerial(appDate.v) == null) defs.push({ t: 'No appraisal date', ref: appDate.ref, pts: 1 });
        if (!(n0(asIsIn.v) > 0 || n0(arvIn.v) > 0)) defs.push({ t: 'No appraised value', ref: asIsIn.ref, pts: 1 });
        [[cert, "Appraiser's certification not attached"], [exec, 'Appraisal not executed'], [lic, 'Appraiser license not attached'], [eo, 'Appraiser E&O insurance not attached']]
          .forEach(([p, t]) => { if (trim(p.v).toUpperCase() === 'N') defs.push({ t, ref: p.ref, pts: 1 }); });
      }
      const vScore = defs.reduce((a, d) => a + d.pts, 0);
      loan.valuation = { type: typeStr, score: vScore, deficiencies: defs };

      /* --- grades (same rules as the Summary tab) --- */
      const fails = loan.checks.filter((c) => c.result === 'FAIL' || c.result === 'PENDING').length;
      const waived = loan.checks.filter((c) => c.result === 'Waived').length;
      const issues = loan.missingDocs.count + fails;
      let credit;
      if (issues === 0) credit = waived === 0 ? 'A' : waived <= 2 ? 'B' : 'C';
      else if (issues === 1) credit = waived === 0 ? 'B' : 'C';
      else if (issues === 2) credit = 'C';
      else credit = 'D';
      let valuation;
      if (APPRAISAL_LIST.includes(norm(typeStr)) && vScore === 0) valuation = 'B';
      else valuation = vScore === 0 ? 'A' : vScore <= 2 ? 'B' : 'C';
      const rank = { A: 0, B: 1, C: 2, D: 3 };
      const overall = [credit, valuation].sort((a, b) => rank[b] - rank[a])[0];
      loan.counts = { missing: loan.missingDocs.count, fails, waived };
      loan.grades = { credit, valuation, overall };

      loan.reasons = [];
      loan.reasons.push('Credit grade ' + credit + ': ' + loan.missingDocs.count + ' missing document(s) + ' + fails + ' guideline fail/pending = ' + issues + ' issue(s), with ' + waived + ' waiver(s).' +
        (issues === 0 && waived === 0 ? ' A clean file earns an A.' : issues === 1 ? ' One issue caps the grade at B (C if a waiver is also used).' : issues === 2 ? ' Two issues is a C.' : issues >= 3 ? ' Three or more issues is a D.' : ' Waivers alone cap the grade at B (1-2) or C (3+).'));
      loan.reasons.push('Valuation grade ' + valuation + ': valuation score ' + vScore + (typeStr ? ' on a ' + typeStr + ' valuation' : '') + (vScore === 0 && valuation === 'B' ? ' (BPO / hybrid / desktop / other valuations cap at B)' : '') + '.');
      loan.reasons.push('Overall grade ' + overall + ': the lowest of the credit and valuation grades.');

      /* --- path to A --- */
      loan.pathToA = [];
      loan.checks.forEach((c) => {
        const shown = c.name === 'Loan Amount' ? fmtMoney(c.value) : typeof c.value === 'number' && c.value < 5 && c.name !== 'FICO' && c.name !== 'Term' ? fmtPct(c.value) : c.value;
        const lim = c.name === 'Loan Amount' ? fmtMoney(c.limit) : typeof c.limit === 'number' && c.limit < 5 ? fmtPct(c.limit) : c.limit;
        if (c.result === 'FAIL') loan.pathToA.push({ text: c.name + ' fails (' + shown + ' vs ' + (c.min ? 'minimum ' : 'limit ') + lim + '). Cure the breach or reclassify the loan.', refs: [].concat(c.valueRefs, c.limitRef) });
        if (c.result === 'PENDING') loan.pathToA.push({ text: c.name + ' is PENDING. Complete the missing data so the check can run.', refs: c.valueRefs });
        if (c.result === 'Waived') loan.pathToA.push({ text: c.name + ' is waived. A waiver still counts against the grade; the loan can only reach A if it meets the guideline without one.', refs: [c.waiver] });
      });
      loan.missingDocs.items.forEach((m) => loan.pathToA.push({ text: 'Missing document: "' + m.text + '". Obtain it, or show it is not required, and clear the exception.', refs: [m.ref] }));
      defs.forEach((d) => loan.pathToA.push({ text: 'Valuation: ' + d.t + '.', refs: [d.ref] }));

      /* ---------- observations (data quality / tie-out) ---------- */
      if (!trim(loanTypeIn.v)) obs('issue', 'Loan Type is blank', 'The Loan Type field (Additional Fields) is empty, so the loan has no bucket and every guideline limit is PENDING. It must be populated (for example "Construction").', [loanTypeIn.ref]);
      if (loan.overridesUsed.length) loan.overridesUsed.forEach((o) => obs('info', 'Override applied: ' + o.field, o.field + ' is overridden to ' + o.value + ' (calculated value was ' + o.base + '). ' + (o.comment ? 'Comment: ' + o.comment : ''), [o.ref]));

      // client vs tape tie-out
      const clientHeaders = Object.keys(headers).filter((h) => /^Client /.test(h));
      let tied = 0, compared = 0;
      const varianceMap = {};
      clientHeaders.forEach((ch) => {
        const base = ch.replace(/^Client /, '');
        if (!headers[base]) return;
        const a = g(ch), b = g(base);
        if (a.v == null && b.v == null) return;
        compared++;
        let same;
        const na = num(a.v), nb = num(b.v);
        if (/date/i.test(base)) same = dateSerial(a.v) === dateSerial(b.v);
        else if (na != null && nb != null) same = Math.abs(na - nb) < 0.005;
        else same = trim(a.v).toLowerCase() === trim(b.v).toLowerCase();
        if (same) tied++;
        else obs('issue', 'Client vs tape mismatch: ' + base, 'Client value ' + a.v + ' differs from tape value ' + b.v + '.', [S(a), S(b)]);
        varianceMap[base] = { same, a, b };
      });
      loan.tieOut = { compared, tied };

      // variance flags
      Object.keys(headers).filter((h) => /Variance$/.test(h)).forEach((vh) => {
        const p = g(vh);
        const v = p.v;
        const isZero = v == null || v === '' || num(v) === 0 || (typeof v === 'string' && /^0+$/.test(v));
        if (isZero) return;
        const base = vh.replace(/ Variance$/, '');
        const t = varianceMap[base];
        const extra = t && t.same ? ' The client and tape values actually agree (' + (/date/i.test(base) ? fmtDate(dateSerial(t.b.v)) : t.b.v) + '), so the flag is not explained by the underlying fields; confirm with the client.' : '';
        obs('watch', vh + ' is flagged (' + v + ')', 'The data tape reports a variance of ' + v + ' for ' + base + '.' + extra, [S(p, vh), t ? S(t.a) : null, t ? S(t.b) : null]);
      });

      // note loan term vs calculated term
      const noteTerm = g('Loan Term');
      if (num(noteTerm.v) != null && loan.metrics.term && loan.metrics.term.overridden == null && isNum(loan.metrics.term.value) && num(noteTerm.v) !== loan.metrics.term.value) {
        obs('watch', 'Note loan term differs from calculated term', 'The Note shows a loan term of ' + noteTerm.v + ' months, but first payment to maturity dates give ' + loan.metrics.term.value + ' months. The guideline check uses the calculated term. Consider whether a Term override is appropriate.', [S(noteTerm, 'Loan Term'), firstPay.ref, maturity.ref]);
      }
      if (loan.metrics.term && loan.metrics.term.overridden && num(noteTerm.v) != null && num(noteTerm.v) !== loan.metrics.term.overridden.base)
        obs('info', 'Term uses the Note loan term', 'Calculated term from dates is ' + loan.metrics.term.overridden.base + ' months; the override sets it to ' + loan.metrics.term.value + ', matching the Note loan term of ' + noteTerm.v + '.', [S(noteTerm, 'Loan Term'), loan.metrics.term.overridden.ref]);

      // flood
      const floodDoc = g('Flood Insurance Policy'), inZone = g('In Flood Zone'), zone = g('Flood Zone Code');
      const exAll = missing.map((m) => m.text).join(' ');
      if (/flood/i.test(exAll) && /^n/i.test(trim(inZone.v))) obs('watch', 'Flood policy flagged missing, but property is not in a flood zone', 'The Exceptions column shows a missing flood policy, while the flood cert says the property is not in a flood zone' + (trim(zone.v) ? ' (zone ' + zone.v + ')' : '') + '. A policy may not be required; clearing the exception would remove the missing-document count.', [missing[0].ref, S(inZone, 'In Flood Zone'), S(zone, 'Flood Zone Code'), S(floodDoc, 'Flood Insurance Policy')]);

      // hazard policy expiry
      const hExp = g('Policy Expiration Date', 0);
      const hs = dateSerial(hExp.v), os = dateSerial(origDate.v);
      if (hs != null) {
        S(hExp, 'Hazard Policy Expiration Date');
        if (os != null && hs < os) obs('issue', 'Hazard policy had expired at origination', 'Hazard policy expired ' + fmtDate(hs) + ', before origination on ' + fmtDate(os) + '.', [hExp.ref, origDate.ref]);
        else if (hs < today) obs('watch', 'Hazard policy has expired since origination', 'Hazard policy expired ' + fmtDate(hs) + ' (report date ' + fmtDate(today) + '). It was valid at origination; confirm renewal evidence is on file.', [hExp.ref]);
      }

      // FICO / credit report cross-check with Guarantors tab
      const gs = guarantors.filter((x) => x.code === code);
      const primary = gs.find((x) => /primary/i.test(x.type)) || gs[0];
      if (primary && primary.median != null && num(fico.v) != null && primary.median !== num(fico.v))
        obs('watch', 'FICO on the Report differs from the Guarantors tab', 'Report shows a calculated FICO of ' + fico.v + '; the Guarantors tab bureau scores (' + primary.scores.join(' / ') + ') give a median of ' + primary.median + '. Both ' + (num(fico.v) >= guide.fico.v && primary.median >= guide.fico.v ? 'pass' : 'are not both passing') + ' the minimum of ' + guide.fico.v + '.', [S(fico), primary.ref]);
      if (primary && primary.reportDate != null && os != null) {
        const age = os - primary.reportDate;
        const flag = g('Credit Report Aged');
        if (age > AGED_CREDIT_DAYS && /^n/i.test(trim(flag.v)))
          obs('watch', 'Credit report looks aged but is flagged not aged', 'Credit report dated ' + fmtDate(primary.reportDate) + ' is ' + age + ' days before origination (' + fmtDate(os) + '), over a ' + AGED_CREDIT_DAYS + '-day assumption, yet "Credit Report Aged" is "' + flag.v + '". Confirm against the client\'s aging rule.', [primary.ref.replace('A', 'G'), S(flag, 'Credit Report Aged'), origDate.ref]);
      }

      // cached Summary-tab cross-check
      const cached = summaryGrades[code];
      if (cached && cached.overall && cached.overall !== overall)
        obs('watch', 'Workbook Summary tab shows a different overall grade', 'The uploaded workbook\'s Summary tab shows ' + cached.overall + ', but recalculating from the Report tab gives ' + overall + '. The workbook may be on manual calculation (press F9) or hold stale values.', [cached.ref]);
      loan.summaryCached = cached || null;
      return loan;
    }

    if (missingFields.size) warnings.push('Report tab is missing expected columns: ' + Array.from(missingFields).join(', ') + '.');
    const meetsExpectation = loans.filter((l) => !l.cross).every((l) => l.grades.overall === expected);
    return { generated: new Date().toISOString(), today, expected, loans, warnings, guide: { buckets, maxTerm: guide.maxTerm, fico: guide.fico, maxLoan: guide.maxLoan, maxLoanCon: guide.maxLoanCon }, meetsExpectation, agedCreditDays: AGED_CREDIT_DAYS };
  }

  return { analyze, days360, fmtDate, fmtMoney, fmtPct, dateSerial, norm };
});
