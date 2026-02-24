const express = require('express');
const path = require('path');
const buildDocx = require('./lib/buildDocx');

const app = express();
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

app.post('/generate', async (req, res) => {
  try {
    const config = req.body;
    const buffer = await buildDocx(config);
    const fname = `MIT_Universal_AI_Module_${config.modNum || 'N'}_Template.docx`;
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="${fname}"`);
    res.send(buffer);
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: e.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on http://localhost:${PORT}`));
