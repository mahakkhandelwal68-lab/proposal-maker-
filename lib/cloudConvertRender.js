const CLOUDCONVERT_API = 'https://api.cloudconvert.com/v2';

// Multiple free CloudConvert accounts, each with its own daily credit
// allowance. Set CLOUDCONVERT_API_KEYS as a comma-separated list to enable
// automatic fallback — if one account is out of credits for the day, the
// next one is tried automatically. CLOUDCONVERT_API_KEY (singular) still
// works on its own for a single-account setup.
function getApiKeys() {
  const multi = process.env.CLOUDCONVERT_API_KEYS;
  if (multi && multi.trim()) {
    return multi.split(',').map((k) => k.trim()).filter(Boolean);
  }
  const single = process.env.CLOUDCONVERT_API_KEY;
  return single ? [single] : [];
}

function isCreditsExhausted(status, bodyText) {
  if (status !== 402) return false;
  return /CREDITS_EXCEEDED/i.test(bodyText) || /run out of conversion credits/i.test(bodyText);
}

async function runJob(apiKey, pptxBuffer, filename) {
  const headers = {
    Authorization: `Bearer ${apiKey}`,
    'Content-Type': 'application/json',
  };

  const createRes = await fetch(`${CLOUDCONVERT_API}/jobs`, {
    method: 'POST',
    headers,
    body: JSON.stringify({
      tag: 'proposal-maker',
      tasks: {
        'import-file': {
          operation: 'import/base64',
          file: pptxBuffer.toString('base64'),
          filename,
        },
        'convert-file': {
          operation: 'convert',
          input: 'import-file',
          output_format: 'pdf',
          engine: 'libreoffice',
        },
        'export-file': {
          operation: 'export/url',
          input: 'convert-file',
        },
      },
    }),
  });

  if (!createRes.ok) {
    const body = await createRes.text();
    const err = new Error(`CloudConvert job creation failed (${createRes.status}): ${body}`);
    err.creditsExhausted = isCreditsExhausted(createRes.status, body);
    throw err;
  }

  const { data: job } = await createRes.json();

  const waitRes = await fetch(`${CLOUDCONVERT_API}/jobs/${job.id}/wait`, {
    headers,
  });
  if (!waitRes.ok) {
    const body = await waitRes.text();
    throw new Error(`CloudConvert job wait failed (${waitRes.status}): ${body}`);
  }

  const { data: finishedJob } = await waitRes.json();

  if (finishedJob.status !== 'finished') {
    const failedTask = finishedJob.tasks.find((t) => t.status === 'error');
    const message = failedTask?.message || '';
    const err = new Error(
      `CloudConvert job did not finish (status: ${finishedJob.status}). ${
        failedTask ? `Task "${failedTask.name}" error: ${message}` : ''
      }`
    );
    err.creditsExhausted = isCreditsExhausted(402, message) || /credit/i.test(message);
    throw err;
  }

  const exportTask = finishedJob.tasks.find(
    (t) => t.name === 'export-file' && t.status === 'finished'
  );
  const fileUrl = exportTask?.result?.files?.[0]?.url;
  if (!fileUrl) throw new Error('CloudConvert job finished but produced no output file URL');

  const fileRes = await fetch(fileUrl);
  if (!fileRes.ok) throw new Error(`Failed to download converted PDF (${fileRes.status})`);

  const arrayBuffer = await fileRes.arrayBuffer();
  return Buffer.from(arrayBuffer);
}

/**
 * Converts a .pptx file (as a Buffer) to a .pdf Buffer using the CloudConvert
 * API (LibreOffice engine under the hood). Works anywhere Node/fetch runs —
 * no PowerPoint or LibreOffice needs to be installed locally, which is what
 * makes this usable from a serverless deployment (e.g. Vercel).
 *
 * Tries each configured API key in turn, moving to the next one only when
 * the current one is specifically out of credits for the day — any other
 * failure (bad file, network issue) is thrown immediately rather than
 * masked by retrying against a different account.
 *
 * @param {Buffer} pptxBuffer
 * @param {string} filename - used only as the job's source filename
 * @returns {Promise<Buffer>} the converted PDF
 */
async function renderPdfViaCloudConvert(pptxBuffer, filename = 'proposal.pptx') {
  const keys = getApiKeys();
  if (keys.length === 0) throw new Error('No CloudConvert API key configured (CLOUDCONVERT_API_KEY / CLOUDCONVERT_API_KEYS)');

  let lastError;
  for (const apiKey of keys) {
    try {
      return await runJob(apiKey, pptxBuffer, filename);
    } catch (err) {
      lastError = err;
      if (!err.creditsExhausted) throw err;
      // else: this account is out of credits for today — try the next one.
    }
  }

  throw new Error(`All CloudConvert accounts are out of credits for today. Last error: ${lastError.message}`);
}

module.exports = { renderPdfViaCloudConvert };
