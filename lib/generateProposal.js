const fs = require('fs');
const os = require('os');
const path = require('path');
const crypto = require('crypto');

const { fillTemplate } = require('./fillTemplate');
const { extractEditableRuns } = require('./extractRuns');
const { getSpecificChangeEdits } = require('./geminiEdit');
const { parsePageList, writeTrimmedTemplate } = require('./trimTemplate');

const TEMPLATES = {
  foundation: {
    file: path.join(__dirname, '..', 'templates', 'foundation.pptx'),
    label: 'Digital Presence & Trust Foundation',
  },
  growth: {
    file: path.join(__dirname, '..', 'templates', 'growth.pptx'),
    label: 'Growth',
    // This deck's cover only carries [Company Name] and a single [Date] —
    // there's no [Client Name] field and no second "valid until" date to
    // fill, unlike Foundation and Automation.
    hasClientName: false,
    hasValidUntil: false,
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
 * If `removePages` is given (e.g. "9, 10, 11"), those pages are cut out of
 * the template before anything else happens, so the AI step and the fill
 * step only ever see the pages that stay. `removePricing` additionally
 * deletes the price/payment boxes that repeat on the other pages (only for
 * packages that define `pricingShapes`).
 *
 * @returns {Promise<{ pdfBuffer: Buffer, filename: string, appliedEdits: string[], unappliedEdits: string[], geminiNote: string, geminiError: string|null, pageNote: string }>}
 */
async function generateProposal({ package: pkg, clientName, companyName, proposalDate, specificChanges, removePages, removePricing }, renderPdf) {
  if (!pkg || !TEMPLATES[pkg]) {
    throw Object.assign(new Error('Unknown or missing package.'), { statusCode: 400 });
  }
  const hasClientName = TEMPLATES[pkg].hasClientName !== false;
  const hasValidUntil = TEMPLATES[pkg].hasValidUntil !== false;
  if (hasClientName && (!clientName || !clientName.trim())) {
    throw Object.assign(new Error('Client name is required.'), { statusCode: 400 });
  }
  if (!companyName || !companyName.trim()) {
    throw Object.assign(new Error('Company name is required.'), { statusCode: 400 });
  }

  const pagesToRemove = parsePageList(removePages);

  const startDate = proposalDate ? new Date(proposalDate) : new Date();
  if (isNaN(startDate.getTime())) {
    throw Object.assign(new Error('Invalid proposal date.'), { statusCode: 400 });
  }
  const validUntilDate = new Date(startDate);
  validUntilDate.setDate(validUntilDate.getDate() + 30);

  const tmpDir = os.tmpdir();
  const id = crypto.randomBytes(6).toString('hex');
  const filenameBase = `${sanitizeForFilename(companyName)}-${pkg}-${id}`;
  const pptxPath = path.join(tmpDir, `${filenameBase}.pptx`);
  const pdfPath = path.join(tmpDir, `${filenameBase}.pdf`);
  const trimmedTemplatePath = path.join(tmpDir, `${filenameBase}-template.pptx`);

  let templatePath = TEMPLATES[pkg].file;
  let pageNote = '';

  const dropPricing = removePricing === true || removePricing === 'true' || removePricing === 'on';
  if (dropPricing && !TEMPLATES[pkg].pricingShapes) {
    throw Object.assign(new Error("Removing the price and payment boxes isn't set up for this package yet."), { statusCode: 400 });
  }

  if (pagesToRemove.length > 0 || dropPricing) {
    const { renumbered } = await writeTrimmedTemplate({
      templatePath,
      outputPath: trimmedTemplatePath,
      pages: pagesToRemove,
      shapesToRemove: dropPricing ? TEMPLATES[pkg].pricingShapes : {},
    });
    templatePath = trimmedTemplatePath;

    const removed = [];
    if (pagesToRemove.length > 0) {
      removed.push(`page${pagesToRemove.length > 1 ? 's' : ''} ${pagesToRemove.join(', ')}`);
    }
    if (dropPricing) removed.push('the price and payment boxes on the remaining pages');
    pageNote =
      `Removed ${removed.join(' and ')}. ` +
      (renumbered
        ? 'Page and section numbers were updated to match. '
        : "Page and section numbers weren't renumbered automatically for this package. ") +
      'Sentences that mention payments or pricing in passing are left as they were, so please read the remaining pages through. You can reword any of them in the specific-changes box.';
  }

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

  const { appliedEdits, unappliedEdits } = await fillTemplate({
    templatePath,
    outputPptxPath: pptxPath,
    clientName: hasClientName ? clientName.trim() : '',
    companyName: companyName.trim(),
    proposalDate: formatDate(startDate),
    validUntil: formatDate(validUntilDate),
    hasClientName,
    hasValidUntil,
    extraEdits,
  });

  const pdfBuffer = await renderPdf(pptxPath, pdfPath);

  // Best-effort cleanup of the scratch files — fine if this fails.
  try { fs.unlinkSync(pptxPath); } catch {}
  try { fs.unlinkSync(pdfPath); } catch {}
  try { fs.unlinkSync(trimmedTemplatePath); } catch {}

  return {
    pdfBuffer,
    filename: `${filenameBase}.pdf`,
    appliedEdits,
    unappliedEdits,
    geminiNote,
    geminiError,
    pageNote,
  };
}

module.exports = { generateProposal, TEMPLATES };
