const CLOUDCONVERT_API = 'https://api.cloudconvert.com/v2';

/**
 * Converts a .pptx file (as a Buffer) to a .pdf Buffer using the CloudConvert
 * API (LibreOffice engine under the hood). Works anywhere Node/fetch runs —
 * no PowerPoint or LibreOffice needs to be installed locally, which is what
 * makes this usable from a serverless deployment (e.g. Vercel).
 *
 * @param {Buffer} pptxBuffer
 * @param {string} filename - used only as the job's source filename
 * @returns {Promise<Buffer>} the converted PDF
 */
async function renderPdfViaCloudConvert(pptxBuffer, filename = 'proposal.pptx') {
  const apiKey = process.env.CLOUDCONVERT_API_KEY;
  if (!apiKey) throw new Error('CLOUDCONVERT_API_KEY is not set');

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
    throw new Error(`CloudConvert job creation failed (${createRes.status}): ${body}`);
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
    throw new Error(
      `CloudConvert job did not finish (status: ${finishedJob.status}). ${
        failedTask ? `Task "${failedTask.name}" error: ${failedTask.message}` : ''
      }`
    );
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

module.exports = { renderPdfViaCloudConvert };
