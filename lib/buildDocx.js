const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, LevelFormat, BorderStyle, WidthType,
  ShadingType, PageBreak, TableOfContents, Footer, PageNumber
} = require('docx');

async function buildDocx(cfg) {
  // Colors
  const MIT_RED   = "A31F34";
  const DARK      = "2D2D2D";
  const MED       = "666666";
  const LIGHT     = "F5F5F5";
  const BORD      = "CCCCCC";
  const BLUE_BG   = "E8F0F8";
  const GREEN_BG  = "E8F5E9";
  const YELLOW_BG = "FFF8E1";
  const PURPLE_BG = "F3E5F5";
  const ORANGE_BG = "FFF3E0";
  const GREEN_H   = "1A6B3C";
  const PURPLE_H  = "6B21A8";
  const ORANGE_H  = "92400E";
  const BLUE_H    = "1D4ED8";

  // ── Primitives ──
  const sp = (pts=120) => new Paragraph({ spacing:{before:pts,after:0}, children:[] });
  const pageBreak = () => new Paragraph({ spacing:{before:360,after:0}, children:[new PageBreak()] });
  const rule = (color="DDDDDD", sz=6) => new Paragraph({
    spacing:{before:60,after:60},
    border:{bottom:{style:BorderStyle.SINGLE,size:sz,color,space:1}},
    children:[]
  });
  const txt = (text, opts={}) => new TextRun({ text, font:"Arial", size:20, color:DARK, ...opts });
  const body = (text, opts={}) => new Paragraph({ spacing:{before:60,after:100},
    children:[txt(text,opts)] });

  // ── Numbering setup ──
  let bIdx = 0;
  const bRefs = [];
  const makeBRef = () => { const r = `b${bIdx++}`; bRefs.push(r); return r; };

  // Pre-allocate refs
  const bGoals   = makeBRef();
  const bLecObj  = makeBRef();
  const bLecKT   = makeBRef();
  const bRecit   = makeBRef();
  const bModKT   = makeBRef();
  const bProd    = makeBRef();

  // ── Field box ──
  function fieldBox(label, placeholder, shade=LIGHT, hint=null) {
    const b = {style:BorderStyle.SINGLE,size:1,color:BORD};
    const borders = {top:b,bottom:b,left:b,right:b};
    const kids = [
      new Paragraph({spacing:{before:0,after:50}, children:[txt(label,{bold:true})]}),
      new Paragraph({spacing:{before:0,after:hint?60:0}, children:[txt(placeholder,{color:"999999",italics:true})]}),
    ];
    if (hint) kids.push(new Paragraph({spacing:{before:0,after:0},
      children:[txt(`💡 ${hint}`,{size:18,color:"888888"})]}));
    return boxWrap(kids, shade, borders);
  }

  function bigFieldBox(label, lines, shade=LIGHT, hint=null) {
    const b = {style:BorderStyle.SINGLE,size:1,color:BORD};
    const borders = {top:b,bottom:b,left:b,right:b};
    const kids = [
      new Paragraph({spacing:{before:0,after:60}, children:[txt(label,{bold:true})]}),
      ...lines.map(line => new Paragraph({spacing:{before:0,after:40},
        children:[txt(line,{color:"999999",italics:true})]})),
    ];
    if (hint) { kids.push(sp(60)); kids.push(new Paragraph({spacing:{before:0,after:0},
      children:[txt(`💡 ${hint}`,{size:18,color:"888888"})]})); }
    return boxWrap(kids, shade, borders);
  }

  function boxWrap(kids, shade, borders) {
    return new Table({
      width:{size:9360,type:WidthType.DXA}, columnWidths:[9360],
      rows:[new TableRow({ children:[new TableCell({
        borders, width:{size:9360,type:WidthType.DXA},
        shading:{fill:shade,type:ShadingType.CLEAR},
        margins:{top:120,bottom:120,left:180,right:180},
        children:kids
      })]})]
    });
  }

  function noteBox(text) {
    const b = {style:BorderStyle.SINGLE,size:4,color:"FFCC02"};
    const borders = {top:b,bottom:b,left:b,right:b};
    return new Table({
      width:{size:9360,type:WidthType.DXA}, columnWidths:[9360],
      rows:[new TableRow({ children:[new TableCell({
        borders, width:{size:9360,type:WidthType.DXA},
        shading:{fill:YELLOW_BG,type:ShadingType.CLEAR},
        margins:{top:100,bottom:100,left:180,right:180},
        children:[new Paragraph({ children:[
          txt("📋 INSTRUCTIONS:  ",{bold:true,size:18,color:"7B5200"}),
          txt(text,{size:18,color:"7B5200"})
        ]})]
      })]})]
    });
  }

  function bullet(text, ref) {
    return new Paragraph({ numbering:{reference:ref,level:0}, spacing:{before:60,after:60},
      children:[txt(text)] });
  }

  function h1(text) {
    return new Paragraph({ style:"MITHeading1", children:[new TextRun({text,font:"Arial"})] });
  }
  function h2(text) {
    return new Paragraph({ style:"MITHeading2", children:[new TextRun({text,font:"Arial"})] });
  }
  function h3(text) {
    return new Paragraph({ style:"MITHeading3", children:[new TextRun({text,font:"Arial"})] });
  }

  // ── Knowledge Check ──
  function knowledgeCheck(num) {
    const b = {style:BorderStyle.SINGLE,size:2,color:"3B82F6"};
    const borders = {top:b,bottom:b,left:b,right:b};
    const mkLine = (t, bold=false) => new Paragraph({spacing:{before:0,after:50},
      children:[txt(t,{bold,color:bold?DARK:"999999",italics:!bold})]});
    return [
      sp(120),
      new Table({
        width:{size:9360,type:WidthType.DXA}, columnWidths:[9360],
        rows:[new TableRow({ children:[new TableCell({
          borders, width:{size:9360,type:WidthType.DXA},
          shading:{fill:"EFF6FF",type:ShadingType.CLEAR},
          margins:{top:140,bottom:140,left:200,right:200},
          children:[
            new Paragraph({spacing:{before:0,after:80}, children:[
              txt(`✓  KNOWLEDGE CHECK ${num}`,{bold:true,color:BLUE_H})
            ]}),
            mkLine("Question:",true),
            new Paragraph({spacing:{before:0,after:120}, children:[
              txt("Type your question here.",{color:"999999",italics:true})
            ]}),
            new Paragraph({spacing:{before:0,after:40}, children:[
              txt("Type:  □ Multiple Choice (single)   □ Multiple Choice (multi)   □ True/False",{size:18,color:MED})
            ]}),
            new Paragraph({spacing:{before:80,after:40}, children:[
              txt("Answer choices (mark correct with ★):",{bold:true})
            ]}),
            ...["A)","B)","C)","D)"].map(l => new Paragraph({spacing:{before:0,after:40},
              children:[txt(`□  ${l} `,{color:"999999",italics:true})]})),
            new Paragraph({spacing:{before:80,after:40}, children:[
              txt("Feedback — correct answer:",{bold:true})
            ]}),
            new Paragraph({spacing:{before:0,after:40}, children:[
              txt("Explain why the correct answer is right.",{color:"999999",italics:true})
            ]}),
            new Paragraph({spacing:{before:60,after:40}, children:[
              txt("Feedback — incorrect answer:",{bold:true})
            ]}),
            new Paragraph({spacing:{before:0,after:0}, children:[
              txt("Guide the learner to reconsider; point to the relevant concept.",{color:"999999",italics:true})
            ]}),
          ]
        })]})]
      }),
      sp(80),
    ];
  }

  // ════════════════════════════════════════════════════════════
  // COVER PAGE
  // ════════════════════════════════════════════════════════════
  const cover = [
    sp(1440),
    new Paragraph({ spacing:{before:0,after:120}, children:[
      txt("MIT Universal AI",{bold:true,size:60,color:MIT_RED})
    ]}),
    new Paragraph({ spacing:{before:0,after:60}, children:[
      txt("Module Authoring Template",{size:36,color:DARK})
    ]}),
    new Paragraph({ spacing:{before:0,after:200}, children:[
      txt(`Module ${cfg.modNum}: ${cfg.modTitle}`,{size:26,color:MED})
    ]}),
    rule(MIT_RED, 8),
    sp(200),
    fieldBox("Module", `Module ${cfg.modNum}: ${cfg.modTitle}`),
    sp(80),
    fieldBox("Author(s)", cfg.authors),
    sp(80),
    fieldBox("Teaching Assistant(s)", cfg.tas || "[TA name]"),
    sp(400),
    ...(cfg.inclInstr ? [
      new Paragraph({spacing:{before:0,after:80},
        children:[txt("How to use this template",{bold:true})]}),
      body("1. Fill in every gray field — replace italic placeholder text with your content."),
      body("2. Yellow instruction boxes are guides — remove them before submitting."),
      body(`3. This template includes ${cfg.lectures} lecture${cfg.lectures!==1?'s':''}, ${cfg.hasRecit?cfg.recitations+' recitation'+(cfg.recitations!==1?'s':''):'no recitations'}, and ${cfg.assignments} assignment${cfg.assignments!==1?'s':''}.`),
      body("4. Submit completed template to your Production Manager for Open edX entry."),
    ] : []),
    pageBreak(),
  ];

  // ════════════════════════════════════════════════════════════
  // TOC
  // ════════════════════════════════════════════════════════════
  const toc = cfg.inclTOC ? [
    h1("Contents"),
    new TableOfContents("Contents", {hyperlink:true, stylesWithLevels:[{styleName:"MITHeading1",level:1},{styleName:"MITHeading2",level:2},{styleName:"MITHeading3",level:3}]}),
    pageBreak(),
  ] : [];

  // ════════════════════════════════════════════════════════════
  // SECTION 1: INTRODUCTION
  // ════════════════════════════════════════════════════════════
  const intro = [
    h1("Section 1: Introduction"),
    rule(MIT_RED),

    h2("1.1  Module Overview"),
    sp(60),
    ...(cfg.inclInstr ? [noteBox("Write 2–4 sentences describing what this module is about, why it matters, and how it fits in the course arc."), sp(80)] : []),
    bigFieldBox("Module Overview", [
      "Write 2–4 sentences here.",
      "",
      `Example: In this module, you will explore the foundations of ${cfg.modTitle}. By the end, you will understand the key concepts and be prepared to apply them in practice.`,
    ], BLUE_BG),
    sp(160),

    h2("1.2  Module Learning Goals"),
    sp(60),
    ...(cfg.inclInstr ? [noteBox("List 3–5 learning goals using action verbs (Bloom's Taxonomy): define, explain, apply, analyze, evaluate, design."), sp(80)] : []),
    new Paragraph({spacing:{before:0,after:80},
      children:[txt("By the end of this module, learners will be able to:",{bold:true})]}),
    bullet("Learning goal 1 — [Action verb] + [concept] + [context or condition]", bGoals),
    bullet("Learning goal 2", bGoals),
    bullet("Learning goal 3", bGoals),
    bullet("Learning goal 4 (optional)", bGoals),
    bullet("Learning goal 5 (optional)", bGoals),
    sp(200),
    pageBreak(),
  ];

  // ════════════════════════════════════════════════════════════
  // SECTION 2: LECTURES
  // ════════════════════════════════════════════════════════════
  const sectionNum = 2;
  const lecturesSection = [
    h1(`Section ${sectionNum}: Lectures`),
    rule(GREEN_H),
    sp(60),
    ...(cfg.inclInstr ? [noteBox(`This module has ${cfg.lectures} lecture${cfg.lectures!==1?'s':''}. Each lecture below is structured identically — fill in each one.`), sp(80)] : []),
  ];

  for (let l = 1; l <= cfg.lectures; l++) {
    lecturesSection.push(
      h2(`Lecture ${l} of ${cfg.lectures}`),
      sp(60),
      fieldBox("Lecture Title", `e.g., Lecture ${l}: [Title]`, GREEN_BG),
      sp(80),
      fieldBox("Estimated Total Video Time", "e.g., 12 minutes across segments", GREEN_BG),
      sp(160),

      h3(`Lecture ${l} Overview`),
      sp(60),
      ...(cfg.inclInstr ? [noteBox("1–2 sentences previewing what this lecture covers. Shown to learners before they watch."), sp(80)] : []),
      bigFieldBox("Lecture Overview", ["Type 1–2 sentences here."], GREEN_BG),
      sp(120),

      h3(`Lecture ${l} Learning Objectives`),
      sp(60),
      ...(cfg.inclInstr ? [noteBox("2–4 specific, measurable objectives scoped to this lecture. These map to the Knowledge Checks below."), sp(80)] : []),
      new Paragraph({spacing:{before:0,after:80},
        children:[txt("After this lecture, learners will be able to:",{bold:true})]}),
      bullet("Objective 1", bLecObj),
      bullet("Objective 2", bLecObj),
      bullet("Objective 3", bLecObj),
      sp(180),

      h3(`Lecture ${l} Videos`),
      sp(60),
      ...(cfg.inclInstr ? [noteBox(`This lecture has ${(cfg.segmentsPerLecture||[])[l-1]||3} segment${((cfg.segmentsPerLecture||[])[l-1]||3)!==1?'s':''}. Fill in the title for each, then complete the Knowledge Check(s) that follow.`), sp(80)] : []),
    );

    // Segments
    const numSegs_l = (cfg.segmentsPerLecture||[])[l-1] || 3;
    for (let s = 1; s <= numSegs_l; s++) {
      const kcTotal = ((cfg.kcsPerSegment||[])[l-1]||[])[s-1] ?? 1;
      lecturesSection.push(
        sp(80),
        new Paragraph({spacing:{before:0,after:80},
          children:[txt(`Lecture ${l}.${s}`,{bold:true,size:22,color:MED})]}),
        fieldBox("Segment Title", `e.g., ${l}.${s} [Segment title]`, GREEN_BG),
      );
      for (let k = 1; k <= kcTotal; k++) {
        lecturesSection.push(...knowledgeCheck(k));
      }
      lecturesSection.push(sp(120));
    }

    // Lecture Summary
    lecturesSection.push(
      h3(`Lecture ${l} Summary`),
      sp(60),
      ...(cfg.inclInstr ? [noteBox("1–2 sentence recap followed by 3–5 key takeaways. Appears as a text block in Open edX after videos."), sp(80)] : []),
      bigFieldBox("Lecture Summary", ["1–2 sentence wrap-up of what was covered and why it matters."], GREEN_BG),
      sp(80),
      new Paragraph({spacing:{before:0,after:80},
        children:[txt("Key Takeaways",{bold:true})]}),
      bullet("Key takeaway 1", bLecKT),
      bullet("Key takeaway 2", bLecKT),
      bullet("Key takeaway 3", bLecKT),
      bullet("Key takeaway 4 (optional)", bLecKT),
      sp(200),
    );

    if (l < cfg.lectures) lecturesSection.push(pageBreak());
  }
  lecturesSection.push(pageBreak());

  // ════════════════════════════════════════════════════════════
  // SECTION 3: RECITATIONS (optional)
  // ════════════════════════════════════════════════════════════
  const recitSection = [];
  if (cfg.hasRecit) {
    const rSec = sectionNum + 1;
    recitSection.push(
      h1(`Section ${rSec}: Recitations`),
      rule(PURPLE_H),
      sp(60),
      ...(cfg.inclInstr ? [noteBox(`This module has ${cfg.recitations} recitation${cfg.recitations!==1?'s':''}. Fill in each section below.`), sp(80)] : []),
    );

    for (let r = 1; r <= cfg.recitations; r++) {
      recitSection.push(
        h2(`Recitation ${r} of ${cfg.recitations}`),
        sp(120),

        h3(`Recitation ${r} Overview`),
        sp(60),
        ...(cfg.inclInstr ? [noteBox("1–2 sentences describing the purpose of this recitation and how it connects to the lectures."), sp(80)] : []),
        bigFieldBox("Recitation Overview", [
          "1–2 sentences here.",
          "Example: This recitation reinforces lecture concepts by walking through a worked example.",
        ], PURPLE_BG),
        sp(120),

        h3(`Recitation ${r} Video(s)`),
        sp(60),
        ...(cfg.inclInstr ? [noteBox("Fill in a block for each video segment."), sp(80)] : []),
      );

      for (let s = 1; s <= cfg.recitSegs; s++) {
        recitSection.push(
          sp(80),
          new Paragraph({spacing:{before:0,after:80}, children:[
            txt(`Recitation ${r}.${s}`,{bold:true,size:22,color:MED})
          ]}),
          fieldBox("Segment Title", `e.g., Recitation ${r}.${s} [Title]`, PURPLE_BG),
          sp(120),
        );
      }

      recitSection.push(
        h3(`Recitation ${r} Summary`),
        sp(60),
        bigFieldBox("Recitation Summary", [
          "1–2 sentence wrap-up of what the recitation demonstrated.",
        ], PURPLE_BG),
        sp(200),
      );
      if (r < cfg.recitations) recitSection.push(pageBreak());
    }
    recitSection.push(pageBreak());
  }

  // ════════════════════════════════════════════════════════════
  // SECTION 4: ASSIGNMENTS (optional)
  // ════════════════════════════════════════════════════════════
  const assignSection = [];
  if (cfg.hasAssign) {
    const aSec = sectionNum + (cfg.hasRecit ? 2 : 1);
    assignSection.push(
      h1(`Section ${aSec}: Assignments`),
      rule(ORANGE_H),
      sp(60),
      ...(cfg.inclInstr ? [noteBox(`This module has ${cfg.assignments} assignment${cfg.assignments!==1?'s':''}. Fill in each section below.`), sp(80)] : []),
    );

    for (let a = 1; a <= cfg.assignments; a++) {
      assignSection.push(
        h2(`Assignment ${a} of ${cfg.assignments}`),
        sp(120),

        h3(`Assignment ${a} Overview`),
        sp(60),
        ...(cfg.inclInstr ? [noteBox("1–2 sentences explaining what the assignment asks learners to do and how it connects to the module learning goals."), sp(80)] : []),
        bigFieldBox("Assignment Overview", [
          "Write 1–2 sentences describing the assignment.",
          "Example: In this assignment, you will apply the concepts from this module to a real dataset.",
        ], ORANGE_BG),
        sp(120),

        h3(`Graded Assignment ${a}`),
        sp(60),
        ...(cfg.inclInstr ? [noteBox("Provide full assignment content. Include all questions, prompts, or problem statements. Attach code notebooks, data files, or rubrics separately."), sp(80)] : []),
        fieldBox("Assignment Title", `e.g., Assignment ${a}: [Title]`, ORANGE_BG),
        sp(80),
        fieldBox("Point Value / Weight", "e.g., 20 points (10% of final grade)", ORANGE_BG),
        sp(80),
        fieldBox("Estimated Completion Time", "e.g., 2–3 hours", ORANGE_BG),
        sp(80),
        bigFieldBox("Instructions to Learners", [
          "Write the instructions exactly as learners will see them.",
          "",
          "Example:",
          "  1. Load and explore the provided dataset.",
          "  2. Build and evaluate a model.",
          "  3. Submit your code and a short written analysis.",
        ], ORANGE_BG, "Attach rubric, starter code, or data files separately and note filenames below."),
        sp(80),
        fieldBox("Attached Files (list filenames)", "e.g., data.csv, starter_code.ipynb, rubric.pdf", ORANGE_BG),
        sp(80),
        bigFieldBox("Grading Rubric / Answer Key (Production Use Only — Not Shown to Learners)", [
          "Provide the rubric or answer key here.",
          "  • Criterion 1: [X pts] Description",
          "  • Criterion 2: [X pts] Description",
          "  • Criterion 3: [X pts] Description",
        ], LIGHT),
        sp(120),

        h3(`Assignment ${a} Summary`),
        sp(60),
        bigFieldBox("Assignment Summary", [
          "1–2 sentences shown to learners after submission, recapping what the assignment assessed.",
        ], ORANGE_BG),
        sp(200),
      );
      if (a < cfg.assignments) assignSection.push(pageBreak());
    }
    assignSection.push(pageBreak());
  }

  // ════════════════════════════════════════════════════════════
  // SECTION 5: CONCLUSION
  // ════════════════════════════════════════════════════════════
  const concludeNum = sectionNum + (cfg.hasRecit?1:0) + (cfg.hasAssign?2:1);
  const conclusion = [
    h1(`Section ${concludeNum}: Conclusion`),
    rule(GREEN_H),
    sp(120),

    h2(`Module Summary`),
    sp(60),
    ...(cfg.inclInstr ? [noteBox("Write 2–3 sentences recapping the entire module — what was covered, what was practiced, and what learners are now equipped to do."), sp(80)] : []),
    bigFieldBox("Module Summary", [
      "Write your 2–3 sentence module summary here.",
      "",
      `Example: In this module, you explored the fundamentals of ${cfg.modTitle}. Through lectures, practice, and hands-on work, you built both conceptual understanding and practical skills.`,
    ], GREEN_BG),
    sp(120),

    h3("Module Key Takeaways"),
    sp(60),
    ...(cfg.inclInstr ? [noteBox("4–6 key things a learner should walk away knowing from the entire module."), sp(80)] : []),
    bullet("Key takeaway 1", bModKT),
    bullet("Key takeaway 2", bModKT),
    bullet("Key takeaway 3", bModKT),
    bullet("Key takeaway 4", bModKT),
    bullet("Key takeaway 5 (optional)", bModKT),
    sp(200),
  ];

  if (cfg.inclProdNotes) {
    conclusion.push(
      rule(),
      sp(80),
      h2("Production Notes (for Production Manager)"),
      sp(60),
      bigFieldBox("Notes to Production Manager", [
        "Any special instructions, accessibility notes, or flagged items for Open edX entry.",
        "",
        "Examples:",
        "  • Video segment references a downloadable file — attached.",
        "  • Knowledge Check requires a diagram — see attached image.",
        "  • Assignment rubric is confidential — do not publish to learners.",
      ], LIGHT),
    );
  }

  // ════════════════════════════════════════════════════════════
  // ASSEMBLE
  // ════════════════════════════════════════════════════════════
  const numberingConfig = bRefs.map(ref => ({
    reference: ref,
    levels: [{ level:0, format:LevelFormat.BULLET, text:"\u2022",
      alignment:AlignmentType.LEFT,
      style:{ paragraph:{ indent:{ left:720, hanging:360 } } } }]
  }));

  const doc = new Document({
    features: { updateFields: true },
    numbering: { config: numberingConfig },
    styles: {
      default: { document: { run: { font:"Arial", size:20, color:DARK } } },
      paragraphStyles: [
        { id:"MITHeading1", name:"MIT Heading 1", basedOn:"Normal", next:"Normal", quickFormat:true,
          run:{size:38,bold:true,font:"Arial",color:MIT_RED},
          paragraph:{spacing:{before:480,after:120},outlineLevel:0} },
        { id:"MITHeading2", name:"MIT Heading 2", basedOn:"Normal", next:"Normal", quickFormat:true,
          run:{size:26,bold:true,font:"Arial",color:DARK},
          paragraph:{spacing:{before:320,after:80},outlineLevel:1} },
        { id:"MITHeading3", name:"MIT Heading 3", basedOn:"Normal", next:"Normal", quickFormat:true,
          run:{size:20,bold:true,font:"Arial",color:MED},
          paragraph:{spacing:{before:200,after:60},outlineLevel:2} },
      ]
    },
    sections:[{
      properties:{
        page:{
          size:{width:12240,height:15840},
          margin:{top:1080,right:1080,bottom:1080,left:1080}
        }
      },
      footers:{
        default: new Footer({ children:[
          new Paragraph({
            alignment:AlignmentType.CENTER,
            children:[
              new TextRun({text:`MIT Universal AI  |  Module ${cfg.modNum}: ${cfg.modTitle}  |  Page `,size:16,font:"Arial",color:"999999"}),
              new TextRun({children:[PageNumber.CURRENT],size:16,font:"Arial",color:"999999"}),
            ]
          })
        ]})
      },
      children:[
        ...cover,
        ...toc,
        ...intro,
        ...lecturesSection,
        ...recitSection,
        ...assignSection,
        ...conclusion,
      ]
    }]
  });

  return await Packer.toBuffer(doc);
}

module.exports = buildDocx;
