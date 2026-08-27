const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');
const { escapeXml, replaceNthOccurrence, replaceWithinSingleRun } = require('./xmlText');

const COVER_SLIDE = 'ppt/slides/slide1.xml';

/**
 * Fills the [Client Name] / [Company Name] / [Date] (x2) placeholders on the
 * cover slide, optionally applies a list of extra {find, replace} text edits
 * across all slides, and writes the result as a new .pptx.
 *
 * extraEdits entries are only applied when the exact `find` text lives
 * inside a single text run — if it spans a formatting boundary (e.g. part
 * bold, part not) it is left unapplied and reported back, rather than risk
 * corrupting the slide or silently guessing.
 *
 * @returns {Promise<{ appliedEdits: string[], unappliedEdits: string[] }>}
 */
async function fillTemplate({
  templatePath,
  outputPptxPath,
  clientName,
  companyName,
  proposalDate,
  validUntil,
  extraEdits = [],
}) {
  const buf = fs.readFileSync(templatePath);
  const zip = await JSZip.loadAsync(buf);

  const coverFile = zip.file(COVER_SLIDE);
  if (!coverFile) throw new Error(`Template is missing ${COVER_SLIDE}`);
  let cover = await coverFile.async('string');

  cover = cover.replace('[Client Name]', escapeXml(clientName));
  cover = cover.replace('[Company Name]', escapeXml(companyName));
  cover = replaceNthOccurrence(cover, '[Date]', escapeXml(proposalDate), 1);
  cover = replaceNthOccurrence(cover, '[Date]', escapeXml(validUntil), 1); // the remaining (2nd) [Date]

  zip.file(COVER_SLIDE, cover);

  const appliedEdits = [];
  const unappliedEdits = [];

  if (extraEdits.length > 0) {
    // Never let a "specific change" edit touch the cover slide — that's
    // where Client Name / Company Name / dates live and must stay exact.
    const slideNames = Object.keys(zip.files).filter(
      (name) => /^ppt\/slides\/slide\d+\.xml$/.test(name) && name !== COVER_SLIDE
    );

    for (const edit of extraEdits) {
      // Defensive backstop: even if the AI step misfires and proposes
      // blanking a line, refuse it here too — several checklist lines in
      // these decks sit next to a separate checkmark icon shape that a
      // text edit can't remove, so an empty replacement leaves an
      // orphaned icon rather than a clean deletion.
      if (!edit.replace || !edit.replace.trim()) {
        unappliedEdits.push(`"${edit.find}" -> "${edit.replace}"`);
        continue;
      }

      const findEscaped = escapeXml(edit.find);
      const replaceEscaped = escapeXml(edit.replace);
      let applied = false;

      for (const slideName of slideNames) {
        const slideXml = await zip.file(slideName).async('string');
        const result = replaceWithinSingleRun(slideXml, findEscaped, replaceEscaped);
        if (result.applied) {
          zip.file(slideName, result.xml);
          applied = true;
          break;
        }
      }

      if (applied) {
        appliedEdits.push(`"${edit.find}" -> "${edit.replace}"`);
      } else {
        unappliedEdits.push(`"${edit.find}" -> "${edit.replace}"`);
      }
    }
  }

  const outBuf = await zip.generateAsync({
    type: 'nodebuffer',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  });

  fs.mkdirSync(path.dirname(outputPptxPath), { recursive: true });
  fs.writeFileSync(outputPptxPath, outBuf);

  return { appliedEdits, unappliedEdits };
}

module.exports = { fillTemplate };
