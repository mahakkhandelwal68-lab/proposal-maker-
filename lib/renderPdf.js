const { spawn } = require('child_process');
const fs = require('fs');
const path = require('path');

const SCRIPT_PATH = path.join(__dirname, 'renderPdf.ps1');

// PowerPoint COM automation can't reliably run two conversions at once,
// so we serialize all render calls through a simple queue.
let queue = Promise.resolve();

function runOnce(pptxPath, pdfPath) {
  return new Promise((resolve, reject) => {
    const ps = spawn('powershell.exe', [
      '-NoProfile',
      '-NonInteractive',
      '-ExecutionPolicy', 'Bypass',
      '-File', SCRIPT_PATH,
      '-PptxPath', pptxPath,
      '-PdfPath', pdfPath,
    ]);

    let stderr = '';
    ps.stderr.on('data', (d) => { stderr += d.toString(); });
    ps.on('error', reject);
    ps.on('close', (code) => {
      if (code === 0) resolve();
      else reject(new Error(`PowerPoint PDF render failed (exit ${code}): ${stderr}`));
    });
  });
}

/**
 * Renders a .pptx to .pdf via PowerPoint COM automation and returns the
 * resulting PDF as a Buffer. Calls are queued since PowerPoint COM can't
 * reliably run two conversions at once.
 */
function renderPdf(pptxPath, pdfPath) {
  const task = queue.then(async () => {
    await runOnce(pptxPath, pdfPath);
    return fs.readFileSync(pdfPath);
  });
  // Keep the queue alive even if this task fails, so later requests still run.
  queue = task.catch(() => {});
  return task;
}

module.exports = { renderPdf };
