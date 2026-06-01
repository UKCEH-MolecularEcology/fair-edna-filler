// ── Mode tabs ──────────────────────────────────────────────────────────────
document.querySelectorAll('.mode-tab').forEach(tab => {
  tab.addEventListener('click', () => {
    document.querySelectorAll('.mode-tab').forEach(t => t.classList.remove('active'));
    document.querySelectorAll('.mode-panel').forEach(p => p.classList.remove('active'));
    tab.classList.add('active');
    document.getElementById('panel-' + tab.dataset.panel).classList.add('active');
  });
});

// ── Upload form ────────────────────────────────────────────────────────────
const form          = document.querySelector('#uploadForm');
const reportInput   = document.querySelector('#report');
const templateInput = document.querySelector('#template');
const reportName    = document.querySelector('#reportName');
const templateName  = document.querySelector('#templateName');
const statusText    = document.querySelector('#statusText');
const progressBar   = document.querySelector('#progress span');
const resultCard    = document.querySelector('#resultCard');
const downloadLink  = document.querySelector('#downloadLink');
const submitBtn     = document.querySelector('.primary-action');

let latestUrl  = '';
let latestName = 'FAIRe_metadata_filled.xlsx';

reportInput.addEventListener('change',   () => { reportName.textContent   = reportInput.files?.[0]?.name   ?? 'Choose a report file'; });
templateInput.addEventListener('change', () => { templateName.textContent = templateInput.files?.[0]?.name ?? 'Choose the FAIRe full template (.xlsx)'; });

function setStatus(msg, pct) {
  statusText.textContent    = msg;
  progressBar.style.width   = pct + '%';
}

form.addEventListener('submit', async e => {
  e.preventDefault();
  resultCard.classList.add('is-hidden');
  submitBtn.disabled = true;
  const hasKey = !!document.querySelector('#apiKey')?.value.trim();
  setStatus(hasKey ? 'Uploading — AI extraction in progress…' : 'Uploading and parsing the report…', 18);

  const data = new FormData(form);
  try {
    const pulse = setInterval(() => {
      const cur = parseFloat(progressBar.style.width) || 25;
      progressBar.style.width = Math.min(cur + 4, 82) + '%';
    }, 700);

    const res  = await fetch('/api/fill', { method: 'POST', body: data });
    clearInterval(pulse);
    const payload = await res.json();

    if (!res.ok || payload.error) throw new Error(payload.error || 'The workbook could not be filled.');

    const set = (id, val) => { const el = document.getElementById(id); if (el) el.textContent = val ?? '—'; };
    set('rProject',     payload.projectTitle);
    set('rMethod',      payload.parsingMethod === 'claude-ai' ? '✦ Claude AI' : 'Rule-based (regex)');
    set('rInstitution', payload.institution);
    set('rLakes',       payload.lakes);
    set('rRounds',      payload.rounds);
    set('rAssays',      Array.isArray(payload.assays) ? payload.assays.join(', ') : payload.assays);
    set('rSampleRows',  payload.sampleRows);
    set('rRunRows',     payload.experimentRows);

    latestUrl  = payload.downloadUrl;
    latestName = payload.outputName || 'FAIRe_metadata_filled.xlsx';
    downloadLink.href     = latestUrl;
    downloadLink.download = latestName;

    setStatus('Done — review the summary then download your workbook.', 100);
    resultCard.classList.remove('is-hidden');
  } catch (err) {
    setStatus('Error: ' + err.message, 100);
  } finally {
    submitBtn.disabled = false;
  }
});

downloadLink.addEventListener('click', async e => {
  if (!latestUrl) return;
  e.preventDefault();
  setStatus('Preparing download…', 100);
  try {
    const res  = await fetch(latestUrl);
    if (!res.ok) throw new Error('Could not retrieve the workbook.');
    const blob = await res.blob();
    const url  = URL.createObjectURL(blob);
    const a    = Object.assign(document.createElement('a'), { href: url, download: latestName });
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
    setStatus('Download started.', 100);
  } catch (err) {
    setStatus(err.message + ' — opening link directly.', 100);
    window.location.href = latestUrl;
  }
});
