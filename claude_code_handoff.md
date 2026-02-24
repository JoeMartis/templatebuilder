# MIT Universal AI — Template Generator
## Claude Code Handoff Document

---

## Project Overview

A local web tool that lets MIT faculty/TAs configure a module structure via a browser form, then downloads a fully formatted `.docx` authoring template. The frontend is plain HTML/CSS/JS; the backend is a minimal Node.js/Express server that generates the `.docx` using the `docx` npm package.

---

## Recommended Project Structure

```
mit-template-generator/
├── server.js               # Express server — handles POST /generate
├── package.json
├── public/
│   └── index.html          # The entire frontend (HTML/CSS/JS in one file)
└── lib/
    └── buildDocx.js        # All docx generation logic (extracted from server.js)
```

---

## Step 1 — Scaffold the project

```bash
mkdir mit-template-generator
cd mit-template-generator
npm init -y
npm install express docx
```

---

## Step 2 — package.json start script

Add this to `package.json`:
```json
"scripts": {
  "start": "node server.js",
  "dev": "node --watch server.js"
}
```

---

## Step 3 — server.js

```js
const express = require('express');
const path = require('path');
const { buildDocx } = require('./lib/buildDocx');

const app = express();
app.use(express.json());
app.use(express.static('public'));

app.post('/generate', async (req, res) => {
  try {
    const cfg = req.body;
    const buffer = await buildDocx(cfg);
    const filename = `MIT_Universal_AI_Module_${cfg.modNum || 'N'}_Template.docx`;
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.send(buffer);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Running at http://localhost:${PORT}`));
```

---

## Step 4 — lib/buildDocx.js

Move ALL docx generation logic here. The function signature should be:

```js
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  HeadingLevel, AlignmentType, LevelFormat, BorderStyle, WidthType,
  ShadingType, PageBreak, TableOfContents, Footer, PageNumber } = require('docx');

async function buildDocx(cfg) {
  // ... all generation logic ...
  // cfg shape documented below
  return await Packer.toBuffer(doc);
}

module.exports = { buildDocx };
```

---

## Step 5 — public/index.html

Same UI as the current HTML tool, but replace the in-browser docx generation with a fetch call:

```js
async function generate() {
  const cfg = collectConfig();   // gather all form values into an object

  const res = await fetch('/generate', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(cfg)
  });

  if (!res.ok) throw new Error(await res.text());

  const blob = await res.blob();
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `MIT_Universal_AI_Module_${cfg.modNum || 'N'}_Template.docx`;
  a.click();
  URL.revokeObjectURL(url);
}
```

---

## Config object shape (sent from frontend → server)

```js
{
  // Module metadata
  modNum:        "3",
  modTitle:      "Supervised Learning",
  authors:       "Prof. Jane Smith",
  tas:           "Alex Johnson",
  time:          "2.5 hours",
  prereqs:       "Basic Python",

  // Structure
  lectures:      3,       // integer, 1–10
  segments:      3,       // video segments per lecture, 1–6
  kcs:           1,       // knowledge checks per segment, 0–3

  hasRecit:      false,
  recitations:   1,       // integer, 1–5 (ignored if hasRecit false)
  recitSegs:     1,       // segments per recitation, 1–4

  hasAssign:     true,
  assignments:   1,       // integer, 1–5 (ignored if hasAssign false)

  // Options
  inclTOC:       true,
  inclInstr:     true,    // include yellow instruction boxes
  inclProdNotes: true,    // include production notes section at end
}
```

---

## Existing assets to port

### 1. buildDocx logic
The full working docx generation code is in the attached `MIT_Template_Generator.html` file (inside the `buildDocx()` async function in the `<script>` tag). 

Key changes when moving to Node.js:
- Remove the `const { Document, ... } = docx;` destructuring from the CDN global — replace with `require('docx')` at the top of `buildDocx.js`
- Remove `Packer.toBuffer` → `saveAs(blob)` — instead just `return await Packer.toBuffer(doc)`
- The rest of the generation logic is identical

### 2. Frontend UI
The full UI is in the same `MIT_Template_Generator.html` file. Copy everything between `<style>` and `</style>` and the `<body>` content verbatim into `public/index.html`. The only JS change is replacing the in-browser `buildDocx()` call with the `fetch('/generate', ...)` approach shown above.

---

## Prompt to give Claude Code

Paste this at the start of your Claude Code session:

---

**CLAUDE CODE PROMPT:**

```
I have a working single-file HTML tool that generates MIT module authoring templates as .docx files. I want to convert it into a proper local Node.js/Express project.

Here is the existing HTML file: [paste or attach MIT_Template_Generator.html]

Please:

1. Create the project structure:
   mit-template-generator/
   ├── server.js
   ├── package.json
   ├── public/index.html
   └── lib/buildDocx.js

2. Move all docx generation logic into lib/buildDocx.js as a module:
   - Use require('docx') instead of the CDN global
   - Export: module.exports = { buildDocx }
   - Return: await Packer.toBuffer(doc)

3. Create server.js with Express:
   - Serve public/ as static files
   - POST /generate — accepts JSON config, calls buildDocx(cfg), streams the .docx back as a download
   - Run on port 3000

4. Update public/index.html:
   - Keep all HTML, CSS, and UI logic exactly as-is
   - Replace the in-browser docx generation with a fetch POST to /generate
   - On response, trigger a browser download of the returned blob
   - Show the existing spinner/status UI during the request

5. Run npm install and verify the server starts with node server.js

6. Test by filling out the form and confirming a valid .docx downloads.

Config object shape the frontend sends:
{
  modNum, modTitle, authors, tas, time, prereqs,
  lectures, segments, kcs,
  hasRecit, recitations, recitSegs,
  hasAssign, assignments,
  inclTOC, inclInstr, inclProdNotes
}
```

---

## Running it

```bash
cd mit-template-generator
npm start
# Open http://localhost:3000
```

---

## Future enhancements to consider

- **Name individual lectures** — add lecture title fields to the form, pass them through cfg, use them as H2 headings instead of "Lecture 1 of N"
- **Save/load configs** — localStorage or a simple JSON file so faculty can resume a session
- **Preview pane** — show an outline of what will be generated before downloading
- **Multiple templates** — support different module types (standard, recitation-heavy, project-based)
- **Accessibility pass** — run the output through an a11y checker before Open edX entry
