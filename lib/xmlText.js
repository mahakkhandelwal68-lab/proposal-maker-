// Small helpers for working with raw OOXML text safely.

function escapeXml(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

// Replaces the Nth (1-indexed) occurrence of `search` in `str` with `replacement`.
// Returns the original string unchanged if there aren't at least N occurrences.
function replaceNthOccurrence(str, search, replacement, n) {
  let idx = -1;
  for (let i = 0; i < n; i++) {
    idx = str.indexOf(search, idx + 1);
    if (idx === -1) return str;
  }
  return str.slice(0, idx) + replacement + str.slice(idx + search.length);
}

// Some of these shapes (big price/stat callouts especially) use
// <a:spAutoFit/> — a box sized to fit specific original text — with a font
// size tuned right to the edge of that box. PowerPoint only recalculates
// that fit when a human edits the text inside its own UI; a direct XML
// text swap (and even swapping in <a:normAutofit/>) leaves the cached
// geometry untouched even through headless COM PDF export, so replacement
// text of a different rendered width can silently wrap and overlap the
// content below it. Since we can't measure real glyph widths here, when
// the replacement is at least as long as the original we defensively
// shrink that run's own font size a little — cheap insurance against a
// wrap, at the cost of a barely-noticeable size difference when it wasn't
// actually needed.
const SHRINK_FACTOR = 0.85;

function shrinkFontSizeForRun(xml, runStartIndex) {
  const searchWindowStart = Math.max(0, runStartIndex - 600);
  const window = xml.slice(searchWindowStart, runStartIndex);
  const szMatches = [...window.matchAll(/ sz="(\d+)"/g)];
  if (szMatches.length === 0) return xml;

  const last = szMatches[szMatches.length - 1];
  const originalSz = parseInt(last[1], 10);
  const newSz = Math.max(100, Math.round(originalSz * SHRINK_FACTOR));
  const matchStartInXml = searchWindowStart + last.index;
  const matchText = last[0];

  return (
    xml.slice(0, matchStartInXml) +
    ` sz="${newSz}"` +
    xml.slice(matchStartInXml + matchText.length)
  );
}

// When an edit empties out a run entirely (a "remove this bullet" request),
// leaving a lone bullet glyph with no text reads as broken, not deleted —
// the bullet marker belongs to the paragraph, not the run, so an empty run
// still shows a checkmark/bullet floating on its own line. If the run we
// just emptied was the paragraph's only text, drop the whole paragraph.
function removeParagraphIfNowEmpty(xml, runStartIndex) {
  const pOpenIdx = xml.lastIndexOf('<a:p>', runStartIndex);
  const altPOpenIdx = xml.lastIndexOf('<a:p ', runStartIndex);
  const openIdx = Math.max(pOpenIdx, altPOpenIdx);
  if (openIdx === -1) return xml;

  const pCloseTag = '</a:p>';
  const closeIdx = xml.indexOf(pCloseTag, runStartIndex);
  if (closeIdx === -1) return xml;

  const paragraph = xml.slice(openIdx, closeIdx + pCloseTag.length);
  const hasRemainingText = [...paragraph.matchAll(/<a:t>([^<]*)<\/a:t>/g)].some(
    (m) => m[1].trim().length > 0
  );
  if (hasRemainingText) return xml;

  return xml.slice(0, openIdx) + xml.slice(closeIdx + pCloseTag.length);
}

// Finds a single <a:t>...</a:t> run whose text contains `findEscaped` and
// replaces just that occurrence within the run, preserving the run's own
// formatting. Returns { xml, applied } — applied is false (xml unchanged)
// if no single run contained the full search text.
function replaceWithinSingleRun(xml, findEscaped, replaceEscaped) {
  const runRegex = /<a:t>([^<]*)<\/a:t>/g;
  let match;
  while ((match = runRegex.exec(xml)) !== null) {
    const runText = match[1];
    if (runText.includes(findEscaped)) {
      const newRunText = runText.replace(findEscaped, replaceEscaped);
      const newRun = `<a:t>${newRunText}</a:t>`;
      let newXml = xml.slice(0, match.index) + newRun + xml.slice(match.index + match[0].length);

      if (newRunText.trim() === '') {
        newXml = removeParagraphIfNowEmpty(newXml, match.index);
      } else if (replaceEscaped.length >= findEscaped.length) {
        newXml = shrinkFontSizeForRun(newXml, match.index);
      }

      return { xml: newXml, applied: true };
    }
  }
  return { xml, applied: false };
}

module.exports = { escapeXml, replaceNthOccurrence, replaceWithinSingleRun };
