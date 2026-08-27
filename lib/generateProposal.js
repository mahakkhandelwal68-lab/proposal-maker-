const fs = require('fs');
const os = require('os');
const path = require('path');
const crypto = require('crypto');

const { fillTemplate } = require('./fillTemplate');
const { extractEditableRuns } = require('./extractRuns');
const { getSpecificChangeEdits } = require('./geminiEdit');

const TEMPLATES = {
  foundation: {
    file: path.join(__dirname, '..', 'templates', 'foundation.pptx'),
    label: 'Digital Presence & Trust Foundation',
  },
  growth: {
    file: path.join(__dirname, '..', 'templates', 'growth.pptx'),
    label: 'Growth',
  },
  automation: {
    file: path.join(__dirname, '..', 'templates', 'automation.pptx'),
    label: 'Business Automation',
  },
};

function formatDate(d) {
  return d.toLocaleDateString('en-GB', { day: 'numeric', month: 'long', year: 'numeric' });
}

function sanitizeForFilename(str) {
  return String(str).replace(/[^a-z0-9]+/gi, '-').replace(/^-+|-+$/g, '').slice(0, 60);
}

/**
 * Runs the full proposal-generation pipeline: validates input, fills the
 * cover fields, optionally asks Gemini to turn a free-text "specific
 * change" request into precise edits, applies everything to the template,
 * and hands the finished .pptx to the given `renderPdf(pptxPath, pdfPath)`
 * function to produce the final PDF.
 *
 * `renderPdf` is injected so the same pipeline works with PowerPoint COM
 * automation (local Windows) or the CloudConvert API (serverless/cloud) —
 * this module has no idea which one it's talking to.
 *
 * @returns {Promise<{ pdfBuffer: Buffer, filename: string, appliedEdits: string[], unappliedEdits: string[], geminiNote: string, geminiError: string|null }>}
 */
async function generateProposal({ package: pkg, clientName, companyName, proposalDate, specificChanges }, renderPdf) {
  if (!pkg || !TEMPLATES[pkg]) {
    throw Object.assign(new Error('Unknown or missing package.'), { statusCode: 400 });
  }
  if (!clientName || !clientName.trim()) {
    throw Object.assign(new Error('Client name is required.'), { statusCode: 400 });
  }
  if (!companyName || !companyName.trim()) {
    throw Object.assign(new Error('Company name is required.'), { statusCode: 400 });
  }

  const templatePath = TEMPLATES[pkg].file;

  const startDate = proposalDate ? new Date(proposalDate) : new Date();
  if (isNaN(startDate.getTime())) {
    throw Object.assign(new Error('Invalid proposal date.'), { statusCode: 400 });
  }
  const validUntilDate = new Date(startDate);
  validUntilDate.setDate(validUntilDate.getDate() + 30);

  let extraEdits = [];
  let geminiNote = '';
  let geminiError = null;

  if (specificChanges && specificChanges.trim()) {
    try {
      const existingRuns = await extractEditableRuns(templatePath);
      const result = await getSpecificChangeEdits({
        packageName: TEMPLATES[pkg].label,
        existingRuns,
        instruction: specificChanges.trim(),
      });
      extraEdits = result.edits;
      geminiNote = result.note;
    } catch (err) {
      // Never fail the whole generation just because the optional AI step
      // had a problem — fall back to the plain filled proposal.
      geminiError = err.message;
    }
  }

  const tmpDir = os.tmpdir();
  const id = crypto.randomBytes(6).toString('hex');
  const filenameBase = `${sanitizeForFilename(companyName)}-${pkg}-${id}`;
  const pptxPath = path.join(tmpDir, `${filenameBase}.pptx`);
  const pdfPath = path.join(tmpDir, `${filenameBase}.pdf`);

  const { appliedEdits, unappliedEdits } = await fillTemplate({
    templatePath,
    outputPptxPath: pptxPath,
    clientName: clientName.trim(),
    companyName: companyName.trim(),
    proposalDate: formatDate(startDate),
    validUntil: formatDate(validUntilDate),
    extraEdits,
  });

  const pdfBuffer = await renderPdf(pptxPath, pdfPath);

  // Best-effort cleanup of the scratch files — fine if this fails.
  try { fs.unlinkSync(pptxPath); } catch {}
  try { fs.unlinkSync(pdfPath); } catch {}

  return {
    pdfBuffer,
    filename: `${filenameBase}.pdf`,
    appliedEdits,
    unappliedEdits,
    geminiNote,
    geminiError,
  };
}

module.exports = { generateProposal, TEMPLATES };
