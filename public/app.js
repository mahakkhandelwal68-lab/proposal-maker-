const form = document.getElementById('form');
const submitBtn = document.getElementById('submitBtn');
const resultEl = document.getElementById('result');

function esc(str) {
  const div = document.createElement('div');
  div.textContent = str;
  return div.innerHTML;
}

function base64ToBlobUrl(base64, mimeType) {
  const bytes = atob(base64);
  const buffer = new Uint8Array(bytes.length);
  for (let i = 0; i < bytes.length; i++) buffer[i] = bytes.charCodeAt(i);
  const blob = new Blob([buffer], { type: mimeType });
  return URL.createObjectURL(blob);
}

form.addEventListener('submit', async (e) => {
  e.preventDefault();
  resultEl.className = 'result hidden';
  resultEl.innerHTML = '';
  submitBtn.disabled = true;
  submitBtn.textContent = 'Generating…';

  const data = Object.fromEntries(new FormData(form).entries());

  try {
    const res = await fetch('/api/generate', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(data),
    });
    const body = await res.json();

    if (!res.ok) {
      throw new Error(body.error || 'Something went wrong.');
    }

    const pdfUrl = base64ToBlobUrl(body.pdfBase64, 'application/pdf');
    let html = `<a class="download" href="${pdfUrl}" download="${esc(body.filename)}">Download proposal PDF →</a>`;

    if (body.geminiNote) {
      html += `<div class="note-block"><div class="note-title">Note from the AI review:</div>${esc(body.geminiNote)}</div>`;
    }
    if (body.unappliedEdits && body.unappliedEdits.length) {
      html += `<div class="note-block"><div class="note-title">Please apply these manually in PowerPoint (couldn't be auto-applied safely):</div><ul>${body.unappliedEdits
        .map((s) => `<li>${esc(s)}</li>`)
        .join('')}</ul></div>`;
    }
    if (body.geminiError) {
      html += `<div class="note-block"><div class="note-title">Specific-change request couldn't be processed automatically:</div>${esc(body.geminiError)} — the proposal was still generated with your name/company/date filled in; please apply the change manually.</div>`;
    }

    resultEl.className = 'result';
    resultEl.innerHTML = html;
  } catch (err) {
    resultEl.className = 'result error';
    resultEl.textContent = err.message;
  } finally {
    submitBtn.disabled = false;
    submitBtn.textContent = 'Generate proposal';
  }
});
