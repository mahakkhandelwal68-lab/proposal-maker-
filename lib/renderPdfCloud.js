const fs = require('fs');
const { renderPdfViaCloudConvert } = require('./cloudConvertRender');

/**
 * Same (pptxPath, pdfPath) -> Buffer interface as lib/renderPdf.js (the
 * local PowerPoint COM renderer), but backed by the CloudConvert API —
 * this is the one used in serverless/cloud deployments where PowerPoint
 * and LibreOffice aren't available.
 */
async function renderPdf(pptxPath, _pdfPath) {
  const pptxBuffer = fs.readFileSync(pptxPath);
  return renderPdfViaCloudConvert(pptxBuffer, require('path').basename(pptxPath));
}

module.exports = { renderPdf };
