const fs = require('fs');
const JSZip = require('jszip');

function decodeXml(str) {
  return String(str)
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&amp;/g, '&');
}

// The cover slide (slide1) also carries the [Client Name] / [Company Name]
// / [Date] placeholders — these must never be offered to the AI edit step,
// even though the rest of that slide's content (e.g. the price summary
// boxes) is fair game just like any other slide.
const PROTECTED_PLACEHOLDERS = new Set(['[Client Name]', '[Company Name]', '[Date]']);

/**
 * Returns the de-duplicated list of visible text runs found across the
 * whole deck, excluding only the Client/Company/Date placeholder fields
 * (never the rest of the cover slide, which may contain real content like
 * price summaries), in reading order.
 */
async function extractEditableRuns(templatePath) {
  const buf = fs.readFileSync(templatePath);
  const zip = await JSZip.loadAsync(buf);

  const slideNames = Object.keys(zip.files)
    .filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))
    .sort((a, b) => {
      const na = parseInt(a.match(/slide(\d+)\.xml/)[1], 10);
      const nb = parseInt(b.match(/slide(\d+)\.xml/)[1], 10);
      return na - nb;
    });

  const seen = new Set();
  const runs = [];

  for (const slideName of slideNames) {
    const xml = await zip.file(slideName).async('string');
    const runRegex = /<a:t>([^<]*)<\/a:t>/g;
    let match;
    while ((match = runRegex.exec(xml)) !== null) {
      const text = decodeXml(match[1]).trim();
      if (text && !PROTECTED_PLACEHOLDERS.has(text) && !seen.has(text)) {
        seen.add(text);
        runs.push(text);
      }
    }
  }

  return runs;
}

module.exports = { extractEditableRuns, decodeXml };
