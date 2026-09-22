const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');

const PRESENTATION = 'ppt/presentation.xml';
const PRESENTATION_RELS = 'ppt/_rels/presentation.xml.rels';
const CONTENT_TYPES = '[Content_Types].xml';
const NUMERIC_RUN = /<a:t>(\d{1,2})<\/a:t>/g;

function badRequest(message) {
  return Object.assign(new Error(message), { statusCode: 400 });
}

/**
 * Turns "9, 10, 11", "9-11", "9 10 11" (or an array of numbers) into a
 * sorted, de-duplicated list of page numbers.
 */
function parsePageList(input) {
  if (input === undefined || input === null) return [];
  const parts = Array.isArray(input) ? input.map(String) : String(input).split(/[\s,;]+/);
  const pages = new Set();
  for (const part of parts) {
    const trimmed = part.trim();
    if (!trimmed) continue;
    const range = trimmed.match(/^(\d+)\s*-\s*(\d+)$/);
    if (range) {
      const from = parseInt(range[1], 10);
      const to = parseInt(range[2], 10);
      if (from > to) throw badRequest(`"${trimmed}" is not a valid page range.`);
      for (let p = from; p <= to; p++) pages.add(p);
    } else if (/^\d+$/.test(trimmed)) {
      pages.add(parseInt(trimmed, 10));
    } else {
      throw badRequest(`"${trimmed}" is not a valid page number. Use something like 9, 10, 11 or 9-11.`);
    }
  }
  return [...pages].sort((a, b) => a - b);
}

function relationshipTag(relsXml, rId) {
  const match = relsXml.match(new RegExp(`<Relationship\\b[^>]*\\bId="${rId}"[^>]*/>`));
  return match ? match[0] : null;
}

function targetOf(tag) {
  const match = tag.match(/\bTarget="([^"]+)"/);
  return match ? match[1] : null;
}

function resolveFrom(dir, target) {
  return path.posix.normalize(target.startsWith('/') ? target.slice(1) : path.posix.join(dir, target));
}

// Removes one top-level shape (a text box, a group, a line ...) from a
// slide's XML by its shape id, together with everything nested inside it.
// The template is checked rather than trusted: a missing id means the deck
// was edited since this list was written, so fail loudly instead of quietly
// leaving the content in a client's proposal.
function removeShapeById(xml, id) {
  const idMatch = new RegExp(`<p:cNvPr\\b[^>]* id="${id}"`).exec(xml);
  if (!idMatch) throw new Error(`Template shape ${id} was not found, so it could not be removed.`);

  // A shape's own <p:cNvPr> is the first thing inside its opening tag, so
  // the last opening tag before it is the shape that owns this id.
  let start = -1;
  let tag = null;
  for (const open of xml.slice(0, idMatch.index).matchAll(/<p:(sp|grpSp|cxnSp|pic|graphicFrame)>/g)) {
    start = open.index;
    tag = open[1];
  }
  if (start === -1) throw new Error(`Template shape ${id} could not be located.`);

  // Groups can contain other groups, so walk the nesting to find the
  // matching close tag rather than stopping at the first one.
  const tagRegex = new RegExp(`<p:${tag}>|</p:${tag}>`, 'g');
  tagRegex.lastIndex = start;
  let depth = 0;
  let match;
  while ((match = tagRegex.exec(xml)) !== null) {
    depth += match[0].startsWith('</') ? -1 : 1;
    if (depth === 0) return xml.slice(0, start) + xml.slice(match.index + match[0].length);
  }
  throw new Error(`Template shape ${id} has no closing tag.`);
}

// Reads the deck's slide order: [{ rId, sldIdTag, slidePath }] in page order.
async function readSlideOrder(zip) {
  const pres = await zip.file(PRESENTATION).async('string');
  const rels = await zip.file(PRESENTATION_RELS).async('string');
  const order = [];
  for (const match of pres.matchAll(/<p:sldId\b[^>]*\br:id="([^"]+)"[^>]*\/>/g)) {
    const rId = match[1];
    const tag = relationshipTag(rels, rId);
    if (!tag) throw new Error(`Template is missing the relationship for slide ${rId}.`);
    order.push({ rId, sldIdTag: match[0], slidePath: resolveFrom('ppt', targetOf(tag)) });
  }
  return order;
}

// The decks print two kinds of running numbers on the slides:
//   - a section counter (01, 02, 03 ... on each block of content)
//   - a page number, which is the last number on a slide and equals its
//     position in the deck
// For each slide this splits its numeric runs into those two groups.
async function readNumbering(zip, order) {
  const slides = [];
  for (let i = 0; i < order.length; i++) {
    const xml = await zip.file(order[i].slidePath).async('string');
    const values = [...xml.matchAll(NUMERIC_RUN)].map((m) => parseInt(m[1], 10));
    const hasPageNumber = values.length > 0 && values[values.length - 1] === i + 1;
    slides.push({ hasPageNumber, sectionValues: hasPageNumber ? values.slice(0, -1) : values });
  }
  return slides;
}

// Renumbering rewrites bare digits inside slide text, which is only safe
// when the deck really uses the pattern above: every page between the
// cover and the closing page carries its page number, and the section
// numbers run 1, 2, 3 ... with no gaps or repeats across the whole deck.
// Decks that mix real digits into their content (say, a "1 2 3 4 5" step
// list that happens to look like a sequence) or that skip numbers fail this
// check and are left alone rather than risk corrupting actual text.
function hasCleanNumbering(numbering) {
  const middlePagesAllNumbered = numbering.slice(1, -1).every((s) => s.hasPageNumber);
  if (numbering.length < 3 || !middlePagesAllNumbered) return false;
  const all = numbering.flatMap((s) => s.sectionValues);
  return all.length > 0 && all.every((value, i) => value === i + 1);
}

async function renumber(zip, remaining) {
  let sectionCounter = 1;
  for (let i = 0; i < remaining.length; i++) {
    const { slidePath, hasPageNumber } = remaining[i];
    const xml = await zip.file(slidePath).async('string');
    const total = [...xml.matchAll(NUMERIC_RUN)].length;
    let index = 0;
    const updated = xml.replace(NUMERIC_RUN, () => {
      const isPage = hasPageNumber && index === total - 1;
      index++;
      return `<a:t>${isPage ? i + 1 : String(sectionCounter++).padStart(2, '0')}</a:t>`;
    });
    zip.file(slidePath, updated);
  }
}

/**
 * Removes the given pages (1-based, in deck order) from an already-loaded
 * pptx zip, and any listed shapes from the pages that stay, then re-flows
 * the deck's page and section numbers when that is safe. Page 1 (the cover,
 * which holds the client/company/date fields) can't be removed, and at
 * least one other page must remain.
 *
 * `shapesToRemove` maps a slide file path (e.g. "ppt/slides/slide12.xml") to
 * the ids of top-level shapes to delete from it. Entries for slides that are
 * removed anyway are ignored.
 *
 * @returns {Promise<{ renumbered: boolean }>}
 */
async function removePages(zip, pages, shapesToRemove = {}) {
  const order = await readSlideOrder(zip);

  for (const page of pages) {
    if (page < 1 || page > order.length) {
      throw badRequest(`This proposal has ${order.length} pages, so page ${page} doesn't exist.`);
    }
  }
  if (pages.includes(1)) {
    throw badRequest("The cover page (page 1) can't be removed.");
  }
  if (pages.length >= order.length) {
    throw badRequest('At least one page besides the cover has to stay in the proposal.');
  }

  const numbering = await readNumbering(zip, order);
  const clean = hasCleanNumbering(numbering);

  const removedSlidePaths = new Set(pages.map((page) => order[page - 1].slidePath));
  for (const [slidePath, ids] of Object.entries(shapesToRemove)) {
    if (removedSlidePaths.has(slidePath)) continue;
    const slideFile = zip.file(slidePath);
    if (!slideFile) throw new Error(`Template slide ${slidePath} was not found.`);
    let xml = await slideFile.async('string');
    for (const id of ids) xml = removeShapeById(xml, id);
    zip.file(slidePath, xml);
  }

  let pres = await zip.file(PRESENTATION).async('string');
  let presRels = await zip.file(PRESENTATION_RELS).async('string');
  let contentTypes = await zip.file(CONTENT_TYPES).async('string');
  const partsToDelete = [];

  for (const page of pages) {
    const { rId, sldIdTag, slidePath } = order[page - 1];
    const slideName = path.posix.basename(slidePath);
    const slideRelsPath = `${path.posix.dirname(slidePath)}/_rels/${slideName}.rels`;

    pres = pres.replace(sldIdTag, '');
    presRels = presRels.replace(relationshipTag(presRels, rId), '');
    partsToDelete.push(slidePath, slideRelsPath);

    // Each slide's speaker-notes part points back at the slide, so it has to
    // go too or the package is left with a dangling reference.
    const slideRels = zip.file(slideRelsPath);
    if (slideRels) {
      const relsXml = await slideRels.async('string');
      for (const rel of relsXml.matchAll(/<Relationship\b[^>]*\/>/g)) {
        if (!/relationships\/notesSlide"/.test(rel[0])) continue;
        const notesPath = resolveFrom(path.posix.dirname(slidePath), targetOf(rel[0]));
        const notesName = path.posix.basename(notesPath);
        partsToDelete.push(notesPath, `${path.posix.dirname(notesPath)}/_rels/${notesName}.rels`);
      }
    }
  }

  // The deck's presentation.xml.rels also points straight at each notes
  // slide, so those links have to be dropped along with the parts.
  const deletedParts = new Set(partsToDelete);
  for (const rel of [...presRels.matchAll(/<Relationship\b[^>]*\/>/g)]) {
    const target = targetOf(rel[0]);
    if (target && !/TargetMode="External"/.test(rel[0]) && deletedParts.has(resolveFrom('ppt', target))) {
      presRels = presRels.replace(rel[0], '');
    }
  }

  for (const part of partsToDelete) {
    contentTypes = contentTypes.replace(new RegExp(`<Override\\b[^>]*PartName="/${part.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}"[^>]*/>`), '');
    zip.remove(part);
  }

  zip.file(PRESENTATION, pres);
  zip.file(PRESENTATION_RELS, presRels);
  zip.file(CONTENT_TYPES, contentTypes);

  if (clean) {
    const removed = new Set(pages);
    const remaining = order
      .map((slide, i) => ({ slidePath: slide.slidePath, hasPageNumber: numbering[i].hasPageNumber, page: i + 1 }))
      .filter((slide) => !removed.has(slide.page));
    await renumber(zip, remaining);
  }

  return { renumbered: clean };
}

/**
 * Reads a template .pptx, removes the given pages and shapes (see
 * removePages), and writes the result to `outputPath`.
 *
 * @returns {Promise<{ renumbered: boolean }>}
 */
async function writeTrimmedTemplate({ templatePath, outputPath, pages = [], shapesToRemove = {} }) {
  const zip = await JSZip.loadAsync(fs.readFileSync(templatePath));
  const { renumbered } = await removePages(zip, pages, shapesToRemove);
  const buf = await zip.generateAsync({
    type: 'nodebuffer',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  });
  fs.mkdirSync(path.dirname(outputPath), { recursive: true });
  fs.writeFileSync(outputPath, buf);
  return { renumbered };
}

module.exports = { parsePageList, removePages, removeShapeById, writeTrimmedTemplate };
