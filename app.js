(function () {
  'use strict';
  const $ = (id) => document.getElementById(id);
  let workbook = null, fileName = '', analysis = null, opts = {};

  const style = document.createElement('style');
  style.textContent = SetpointReport.CSS;
  document.head.appendChild(style);

  function status(msg, err) { const s = $('status'); s.textContent = msg || ''; s.className = 'status' + (err ? ' err' : ''); }
  function readOpts() {
    opts = { title: $('title').value || 'Loan Tape Review Report', preparedBy: $('prepared').value.trim(), clientName: $('client').value.trim(), fileName: fileName };
  }

  function run() {
    if (!workbook) return;
    readOpts();
    try {
      analysis = SetpointEngine.analyze(XLSX, workbook, { expectedGrade: $('expected').value });
      $('report').innerHTML = SetpointReport.render(analysis, opts);
      $('actions').hidden = false;
      const n = analysis.loans.length;
      status(n + ' loan' + (n === 1 ? '' : 's') + ' analysed from ' + (fileName || 'demo data') + '.');
      $('report').scrollIntoView({ behavior: 'smooth', block: 'start' });
    } catch (e) {
      analysis = null; $('report').innerHTML = ''; $('actions').hidden = true;
      status(e.message || String(e), true);
    }
  }

  function loadFile(file) {
    if (!file) return;
    status('Reading ' + file.name + '...');
    const reader = new FileReader();
    reader.onerror = () => status('Could not read that file.', true);
    reader.onload = (ev) => {
      try {
        workbook = XLSX.read(new Uint8Array(ev.target.result), { type: 'array', cellFormula: true });
        fileName = file.name;
        run();
      } catch (e) { status('That does not look like a valid Excel file: ' + e.message, true); }
    };
    reader.readAsArrayBuffer(file);
  }

  const drop = $('drop'), input = $('file');
  drop.addEventListener('click', () => input.click());
  drop.addEventListener('keydown', (e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); input.click(); } });
  input.addEventListener('change', () => loadFile(input.files[0]));
  ['dragenter', 'dragover'].forEach((t) => drop.addEventListener(t, (e) => { e.preventDefault(); drop.classList.add('over'); }));
  ['dragleave', 'drop'].forEach((t) => drop.addEventListener(t, (e) => { e.preventDefault(); drop.classList.remove('over'); }));
  drop.addEventListener('drop', (e) => loadFile(e.dataTransfer.files[0]));

  ['title', 'prepared', 'client', 'expected'].forEach((id) => $(id).addEventListener('change', run));
  $('demo').addEventListener('click', () => { workbook = SetpointDemo.build(XLSX); fileName = ''; run(); });

  function download(name, text, type) {
    const a = document.createElement('a');
    a.href = URL.createObjectURL(new Blob([text], { type: type }));
    a.download = name; document.body.appendChild(a); a.click();
    setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 500);
  }
  const stem = () => (fileName ? fileName.replace(/\.[^.]+$/, '') : 'demo') + '-report';
  $('dl-html').addEventListener('click', () => { readOpts(); download(stem() + '.html', SetpointReport.standalone(analysis, opts), 'text/html'); });
  $('dl-csv').addEventListener('click', () => download(stem() + '-sources.csv', SetpointReport.sourcesCsv(analysis), 'text/csv'));
  $('dl-json').addEventListener('click', () => download(stem() + '.json', JSON.stringify(analysis, null, 2), 'application/json'));
  $('print').addEventListener('click', () => window.print());
  $('copy-email').addEventListener('click', () => {
    readOpts();
    const text = SetpointReport.buildEmail(analysis, opts);
    (navigator.clipboard ? navigator.clipboard.writeText(text) : Promise.reject()).then(
      () => status('Email copied to clipboard.'),
      () => { download(stem() + '-email.txt', text, 'text/plain'); status('Clipboard blocked; downloaded the email as a text file instead.'); });
  });
})();
