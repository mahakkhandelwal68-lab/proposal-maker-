const { generateProposal } = require('../lib/generateProposal');
const { renderPdf } = require('../lib/renderPdfCloud');

module.exports = async (req, res) => {
  if (req.method !== 'POST') {
    res.status(405).json({ error: 'Method not allowed' });
    return;
  }

  try {
    const result = await generateProposal(req.body || {}, renderPdf);
    res.status(200).json({
      pdfBase64: result.pdfBuffer.toString('base64'),
      filename: result.filename,
      appliedEdits: result.appliedEdits,
      unappliedEdits: result.unappliedEdits,
      geminiNote: result.geminiNote,
      geminiError: result.geminiError,
    });
  } catch (err) {
    console.error(err);
    res.status(err.statusCode || 500).json({ error: err.message || 'Something went wrong generating the proposal.' });
  }
};
