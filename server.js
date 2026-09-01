require('dotenv').config();
const express = require('express');
const path = require('path');

const { generateProposal } = require('./lib/generateProposal');
const { renderPdf } = require('./lib/renderPdf');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(express.json());
app.use(express.static(path.join(__dirname, 'public'), { extensions: ['html'] }));

app.post('/api/generate', async (req, res) => {
  try {
    const result = await generateProposal(req.body || {}, renderPdf);
    res.json({
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
});

app.listen(PORT, () => {
  console.log(`Lumelush proposal maker running at http://localhost:${PORT}`);
});
