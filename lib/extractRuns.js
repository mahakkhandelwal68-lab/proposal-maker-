const fs = require('fs');
const JSZip = require('jszip');

function decodeXml(str) {
  return String(str)
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&amp;/g, '&');
}

/**
 * Returns the de-duplicated list of visible text runs found on every slide
 * except the cover (slide1, which holds the Client/Company/Date fields we
 * never want an AI edit to touch), in reading order.
 */
async function extractEditableRuns(templatePath) {
  const buf = fs.readFileSync(templatePath);
  const zip = await JSZip.loadAsync(buf);

  const slideNames = Object.keys(zip.files)
    .filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))
    .filter((name) => name !== 'ppt/slides/slide1.xml')
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
      if (text && !seen.has(text)) {
        seen.add(text);
        runs.push(text);
      }
    }
  }

  return runs;
}

module.exports = { extractEditableRuns, decodeXml };
