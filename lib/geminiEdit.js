const GEMINI_MODEL = 'gemini-3.5-flash-lite';
const GEMINI_URL = `https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_MODEL}:generateContent`;

const RESPONSE_SCHEMA = {
  type: 'object',
  properties: {
    edits: {
      type: 'array',
      items: {
        type: 'object',
        properties: {
          find: { type: 'string' },
          replace: { type: 'string' },
        },
        required: ['find', 'replace'],
      },
    },
    note: { type: 'string' },
  },
  required: ['edits', 'note'],
};

/**
 * Asks Gemini to turn a free-text "specific change" request into precise
 * find/replace edits against the deck's existing text runs. `find` values
 * are constrained to copy verbatim from `existingRuns` so every edit can be
 * applied deterministically (no free-form rewriting of the document).
 *
 * @returns {Promise<{ edits: {find:string, replace:string}[], note: string }>}
 */
async function getSpecificChangeEdits({ packageName, existingRuns, instruction }) {
  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey) throw new Error('GEMINI_API_KEY is not set');

  const prompt = `You are editing a commercial proposal document (PowerPoint deck, package: "${packageName}") for Lumelush Studio, a digital agency.

An employee typed this request for a specific change to make to the proposal before it's sent to a client:
"""
${instruction}
"""

Below is the exact list of every editable text line currently in the deck (each is one line of visible text). You may ONLY propose edits where "find" is copied EXACTLY, character-for-character, from one of these lines. Do not invent, paraphrase, or partially match — copy the line verbatim.

"replace" must always be non-empty, real replacement text. Never propose an empty string as "replace" — many checklist lines in this deck sit next to a separate checkmark icon shape that a text edit cannot remove, so blanking a line's text leaves an orphaned icon floating with nothing beside it, which looks broken. If the request is to remove/delete a line entirely rather than reword it, do NOT attempt it as an edit — leave edits empty and tell the employee in "note" which line to delete manually in PowerPoint.

If the request maps cleanly onto rewording or changing the value of one or more existing lines (e.g. changing a price, changing a package name, rewording a bullet to say something else), return those as edits.

If the request cannot be safely mapped onto an exact existing line this way (e.g. it asks to add something entirely new, delete a line, or is ambiguous), do NOT guess — leave edits empty and explain what needs manual attention in "note".

Existing lines in the deck:
${existingRuns.map((r) => `- ${r}`).join('\n')}

Respond with the edits (each with the exact original line as "find" and the new full replacement line as "replace") and a short "note" for the employee — empty string if nothing needs manual follow-up.`;

  const requestBody = JSON.stringify({
    contents: [{ role: 'user', parts: [{ text: prompt }] }],
    generationConfig: {
      responseMimeType: 'application/json',
      responseSchema: RESPONSE_SCHEMA,
      temperature: 0.2,
    },
  });

  const delays = [0, 800, 2000]; // ms — a couple of retries for transient 503/429s
  let lastError;

  for (const delay of delays) {
    if (delay) await new Promise((r) => setTimeout(r, delay));

    const res = await fetch(`${GEMINI_URL}?key=${apiKey}`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: requestBody,
    });

    if (res.ok) {
      const data = await res.json();
      return parseGeminiResponse(data);
    }

    const bodyText = await res.text();
    lastError = new Error(`Gemini API error ${res.status}: ${bodyText}`);
    if (res.status !== 503 && res.status !== 429) throw lastError;
  }

  throw lastError;
}

function parseGeminiResponse(data) {
  const text = data?.candidates?.[0]?.content?.parts?.[0]?.text;
  if (!text) throw new Error('Gemini returned no content');

  const parsed = JSON.parse(text);
  return {
    edits: Array.isArray(parsed.edits) ? parsed.edits : [],
    note: typeof parsed.note === 'string' ? parsed.note : '',
  };
}

module.exports = { getSpecificChangeEdits };
