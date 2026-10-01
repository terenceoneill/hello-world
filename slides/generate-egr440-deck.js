const Presentation = require("pptxgenjs");

const pptx = new Presentation();

// Global Deck Settings
pptx.layout = "LAYOUT_WIDE"; // 16:9 Widescreen (13.33 x 7.5 inches)
pptx.author = "Terence O'Neill";
pptx.title = "EGR 440: Researching Your Venture - Scalable Innovation";

// Aesthetic Palette (High-Contrast Dark Slate)
const COLORS = {
  bg: "1E293B",        // Dark Slate
  textMain: "F8FAFC",  // Pure Off-White
  textMuted: "94A3B8", // Cool Gray
  accent: "38BDF8",    // Electric Cyan
  accentAlt: "F59E0B", // Warm Amber
  cardBg: "334155",    // Medium Slate
  border: "475569"     // Soft Divider
};

const FONTS = {
  header: "Calibri",
  body: "Calibri"
};

// Base Slide Style Helper
function createBaseSlide() {
  const slide = pptx.addSlide();
  slide.background = { color: COLORS.bg };
  return slide;
}

// Helper: Standard Slide Header
function addHeader(slide, title, category) {
  if (category) {
    slide.addText(category.toUpperCase(), {
      x: 0.8, y: 0.5, w: 11.5, h: 0.3,
      fontFace: FONTS.body, fontSize: 11, bold: true, color: COLORS.accent, charSpacing: 1.5
    });
  }
  slide.addText(title, {
    x: 0.8, y: 0.8, w: 11.5, h: 0.6,
    fontFace: FONTS.header, fontSize: 26, bold: true, color: COLORS.textMain
  });
}

// Helper: Bottom Contact Bar
function addFooter(slide) {
  slide.addText("EGR 440 | MSU Libraries Entrepreneurship Hub: libguides.lib.msu.edu/entrepreneur", {
    x: 0.8, y: 7.0, w: 11.7, h: 0.3,
    fontFace: FONTS.body, fontSize: 10, color: COLORS.textMuted
  });
}

// -----------------------------------------------------------------------------
// SLIDE 0: Plain Title Slide
// -----------------------------------------------------------------------------
const slide0 = createBaseSlide();

slide0.addText("Researching Your Venture", {
  x: 0.8, y: 2.6, w: 11.7, h: 1.0,
  fontFace: FONTS.header, fontSize: 40, bold: true, color: COLORS.textMain, align: "center"
});
slide0.addText("EGR 440", {
  x: 0.8, y: 3.6, w: 11.7, h: 0.5,
  fontFace: FONTS.body, fontSize: 20, color: COLORS.textMuted, align: "center"
});
slide0.addText("Terence O'Neill  |  Entrepreneurship Librarian, MSU Libraries", {
  x: 0.8, y: 4.6, w: 11.7, h: 0.4,
  fontFace: FONTS.body, fontSize: 16, color: COLORS.textMain, align: "center"
});

// -----------------------------------------------------------------------------
// SLIDE 1: Title & Engineering Assumption Trap
// -----------------------------------------------------------------------------
const slide1 = createBaseSlide();

slide1.addText("EGR 440 | RESEARCHING YOUR VENTURE", {
  x: 0.8, y: 1.2, w: 11.5, h: 0.3,
  fontFace: FONTS.body, fontSize: 12, bold: true, color: COLORS.accent, charSpacing: 2
});

slide1.addText("Engineering Feasibility vs. Scalable Viability", {
  x: 0.8, y: 1.6, w: 11.5, h: 1.2,
  fontFace: FONTS.header, fontSize: 36, bold: true, color: COLORS.textMain
});

// Left Card: Feasibility
slide1.addShape(pptx.shapes.RECTANGLE, {
  x: 0.8, y: 3.2, w: 5.6, h: 2.8,
  fill: { color: COLORS.cardBg }, line: { color: COLORS.border, width: 1 }
});
slide1.addText("SYSTEM FEASIBILITY", {
  x: 1.1, y: 3.5, w: 5.0, h: 0.3,
  fontFace: FONTS.body, fontSize: 13, bold: true, color: COLORS.accentAlt
});
slide1.addText("• \"Can we build it in the lab?\"\n• Sensor precision & code execution\n• Subsystem hardware integration", {
  x: 1.1, y: 4.0, w: 5.0, h: 1.8,
  fontFace: FONTS.body, fontSize: 15, color: COLORS.textMain, lineSpacing: 24
});

// Right Card: Viability
slide1.addShape(pptx.shapes.RECTANGLE, {
  x: 6.9, y: 3.2, w: 5.6, h: 2.8,
  fill: { color: COLORS.cardBg }, line: { color: COLORS.accent, width: 1.5 }
});
slide1.addText("SUPER-SYSTEM VIABILITY", {
  x: 7.2, y: 3.5, w: 5.0, h: 0.3,
  fontFace: FONTS.body, fontSize: 13, bold: true, color: COLORS.accent
});
slide1.addText("• \"Where does it sit in the market?\"\n• Supply chain & CapEx limits\n• Regulatory barriers & prior art", {
  x: 7.2, y: 4.0, w: 5.0, h: 1.8,
  fontFace: FONTS.body, fontSize: 15, color: COLORS.textMain, lineSpacing: 24
});

slide1.addText("\"Intuition sparks the engineering design; library evidence proves its scalability.\"", {
  x: 0.8, y: 6.3, w: 11.7, h: 0.4,
  fontFace: FONTS.body, fontSize: 14, italic: true, color: COLORS.accent, align: "center"
});

// -----------------------------------------------------------------------------
// SLIDE 2: Eugene Shteyn's Scalable Innovation Engine
// -----------------------------------------------------------------------------
const slide2 = createBaseSlide();
addHeader(slide2, "Eugene Shteyn’s Scalable Innovation Engine", "Theoretical Framework");

const levels = [
  { title: "1. SUBSYSTEM", desc: "Components, micro-controllers, sensors, raw materials." },
  { title: "2. SYSTEM", desc: "The assembled physical prototype, software stack, bench process." },
  { title: "3. SUPER-SYSTEM", desc: "FDA/HACCP compliance, co-packers, global supply chains, end-user margins." }
];

levels.forEach((lvl, idx) => {
  const yPos = 1.8 + (idx * 1.6);
  slide2.addShape(pptx.shapes.RECTANGLE, {
    x: 0.8, y: yPos, w: 11.7, h: 1.3,
    fill: { color: COLORS.cardBg }, line: { color: idx === 2 ? COLORS.accent : COLORS.border, width: idx === 2 ? 1.5 : 1 }
  });
  slide2.addText(lvl.title, {
    x: 1.1, y: yPos + 0.2, w: 3.0, h: 0.9,
    fontFace: FONTS.header, fontSize: 16, bold: true, color: idx === 2 ? COLORS.accent : COLORS.accentAlt
  });
  slide2.addText(lvl.desc, {
    x: 4.2, y: yPos + 0.2, w: 8.0, h: 0.9,
    fontFace: FONTS.body, fontSize: 15, color: COLORS.textMain, align: "left"
  });
});

addFooter(slide2);

// -----------------------------------------------------------------------------
// SLIDE 3: Hero Scenario & Assumption Audit
// -----------------------------------------------------------------------------
const slide3 = createBaseSlide();
addHeader(slide3, "Case Study: Micro-Industrial Food-Lab-as-a-Service", "Hero Scenario");

slide3.addText("THE VENTURE IDEA: Modular, robotic cleanroom pods rented as 'Production-as-a-Service' for artisan ferments, koji serums, and functional food creators.", {
  x: 0.8, y: 1.6, w: 11.7, h: 0.6,
  fontFace: FONTS.body, fontSize: 15, bold: true, color: COLORS.accent
});

const assumptions = [
  { title: "Regulatory Audit", desc: "Will local health departments & FDA/HACCP pass modular robotic cleanrooms?" },
  { title: "Supply Chain Audit", desc: "Is the local co-packer radius big enough to justify hardware CapEx?" },
  { title: "Economic Audit", desc: "Do artisan food creators want automated precision or hand-crafted appeal?" }
];

assumptions.forEach((item, idx) => {
  const xPos = 0.8 + (idx * 4.0);
  slide3.addShape(pptx.shapes.RECTANGLE, {
    x: xPos, y: 2.4, w: 3.7, h: 3.6,
    fill: { color: COLORS.cardBg }, line: { color: COLORS.border, width: 1 }
  });
  slide3.addText(`0${idx + 1}`, {
    x: xPos + 0.3, y: 2.6, w: 3.1, h: 0.4,
    fontFace: FONTS.header, fontSize: 22, bold: true, color: COLORS.accentAlt
  });
  slide3.addText(item.title, {
    x: xPos + 0.3, y: 3.1, w: 3.1, h: 0.4,
    fontFace: FONTS.header, fontSize: 16, bold: true, color: COLORS.textMain
  });
  slide3.addText(item.desc, {
    x: xPos + 0.3, y: 3.6, w: 3.1, h: 2.2,
    fontFace: FONTS.body, fontSize: 14, color: COLORS.textMuted, lineSpacing: 20, valign: "top"
  });
});

slide3.addText("ACTION: Open your shared doc now. Log your 3 biggest technical assumptions before searching.", {
  x: 0.8, y: 6.3, w: 11.7, h: 0.4,
  fontFace: FONTS.body, fontSize: 13, bold: true, color: COLORS.textMain, align: "center"
});

addFooter(slide3);

// -----------------------------------------------------------------------------
// SLIDE 4: The Master Data Triad (Single Live Demo Slide)
// -----------------------------------------------------------------------------
const slide4 = createBaseSlide();
addHeader(slide4, "The Database Triad: 15-Minute Live Research Blitz", "Triangulation Framework");

const triad = [
  {
    layer: "LAYER 1",
    name: "IBISWorld & MarketLine",
    focus: "Industry CapEx & Supply Chain",
    terms: "Commercial Catering\nSpecialty Food Manufacturing\nIndustrial Robotics",
    color: COLORS.accent
  },
  {
    layer: "LAYER 2",
    name: "Mintel & Statista",
    focus: "Consumer Pull & Adoption",
    terms: "Kitchen Automation\nGhost Kitchens\nFermented Foods & CPG",
    color: COLORS.accentAlt
  },
  {
    layer: "LAYER 3",
    name: "IEEE Xplore & Data Axle",
    focus: "Prior Art & Local Partners",
    terms: "Automated Food Handling\nData Axle: 50-Mile Radius\nCo-Packers / Commercial Kitchens",
    color: COLORS.accent
  }
];

triad.forEach((col, idx) => {
  const xPos = 0.8 + (idx * 4.0);
  slide4.addShape(pptx.shapes.RECTANGLE, {
    x: xPos, y: 1.6, w: 3.7, h: 4.8,
    fill: { color: COLORS.cardBg }, line: { color: col.color, width: 1.5 }
  });
  slide4.addText(col.layer, {
    x: xPos + 0.3, y: 1.9, w: 3.1, h: 0.3,
    fontFace: FONTS.body, fontSize: 11, bold: true, color: col.color, charSpacing: 1.5
  });
  slide4.addText(col.name, {
    x: xPos + 0.3, y: 2.2, w: 3.1, h: 0.6,
    fontFace: FONTS.header, fontSize: 18, bold: true, color: COLORS.textMain
  });
  slide4.addText(col.focus, {
    x: xPos + 0.3, y: 2.8, w: 3.1, h: 0.4,
    fontFace: FONTS.body, fontSize: 12, italic: true, color: COLORS.textMuted
  });
  slide4.addShape(pptx.shapes.LINE, {
    x: xPos + 0.3, y: 3.3, w: 3.1, h: 0,
    line: { color: COLORS.border, width: 1 }
  });
  slide4.addText("LIVE SEARCH TARGETS:", {
    x: xPos + 0.3, y: 3.5, w: 3.1, h: 0.3,
    fontFace: FONTS.body, fontSize: 10, bold: true, color: COLORS.accentAlt
  });
  slide4.addText(col.terms, {
    x: xPos + 0.3, y: 3.9, w: 3.1, h: 2.2,
    fontFace: FONTS.body, fontSize: 13, color: COLORS.textMain, lineSpacing: 22, valign: "top"
  });
});

addFooter(slide4);

// -----------------------------------------------------------------------------
// SLIDE 5: Database AI & CRAFT Engine
// -----------------------------------------------------------------------------
const slide5 = createBaseSlide();
addHeader(slide5, "Database AI: Structured Prompting with CRAFT", "AI Engine");

slide5.addShape(pptx.shapes.RECTANGLE, {
  x: 0.8, y: 1.6, w: 11.7, h: 4.8,
  fill: { color: COLORS.cardBg }, line: { color: COLORS.accent, width: 1.5 }
});

slide5.addText("THE CRAFT FRAMEWORK (Ethan Mollick Mental Model)", {
  x: 1.2, y: 1.9, w: 10.9, h: 0.4,
  fontFace: FONTS.body, fontSize: 12, bold: true, color: COLORS.accentAlt, charSpacing: 1.5
});

const promptText = [
  { text: "\"Act as a commercial food facility compliance officer ", options: { color: COLORS.textMain } },
  { text: "[ROLE]", options: { color: COLORS.accent, bold: true } },
  { text: " reviewing modular robotic cleanroom kitchens ", options: { color: COLORS.textMain } },
  { text: "[CONTEXT]", options: { color: COLORS.accent, bold: true } },
  { text: ". Identify 5 regulatory and HACCP sanitation bottlenecks ", options: { color: COLORS.textMain } },
  { text: "[ACTION]", options: { color: COLORS.accent, bold: true } },
  { text: " formatted as bullet points with industry standard citations ", options: { color: COLORS.textMain } },
  { text: "[FORMAT]", options: { color: COLORS.accent, bold: true } },
  { text: ". Tone: Analytical ", options: { color: COLORS.textMain } },
  { text: "[TONE]", options: { color: COLORS.accent, bold: true } },
  { text: ".\"", options: { color: COLORS.textMain } }
];

slide5.addText(promptText, {
  x: 1.2, y: 2.4, w: 10.9, h: 2.2,
  fontFace: FONTS.body, fontSize: 18, lineSpacing: 28
});

slide5.addText("CRITICAL RULE: Always prime LLMs with library database context. Never ask AI for facts without data grounding.", {
  x: 1.2, y: 5.4, w: 10.9, h: 0.6,
  fontFace: FONTS.body, fontSize: 13, italic: true, color: COLORS.textMuted
});

addFooter(slide5);

// -----------------------------------------------------------------------------
// SLIDE 5B: CRAFT Starter Prompts (Basic Progression)
// -----------------------------------------------------------------------------
const slide5b = createBaseSlide();
addHeader(slide5b, "CRAFT Starter Prompts: Brainstorm → Search → Make Sense", "AI Engine");

// Turns [["plain text", "TAG"], ...] into styled text runs
function craftRuns(parts) {
  const runs = [];
  parts.forEach(([text, tag]) => {
    runs.push({ text, options: { color: COLORS.textMain } });
    if (tag) runs.push({ text: ` [${tag}]`, options: { color: COLORS.accent, bold: true } });
  });
  return runs;
}

const starterPrompts = [
  {
    step: "01",
    label: "BRAINSTORM",
    parts: [
      ["\"Act as a startup mentor", "ROLE"],
      [". I'm an engineering student with an idea for rentable food-production pods for small food makers", "CONTEXT"],
      [". List 5 questions I should answer before building a prototype", "ACTION"],
      [" as a numbered list", "FORMAT"],
      [". Keep it simple and encouraging", "TONE"],
      [".\"", null]
    ]
  },
  {
    step: "02",
    label: "FIND KEYWORDS",
    parts: [
      ["\"Act as a research librarian", "ROLE"],
      [". I'm researching the market for small-batch food manufacturing", "CONTEXT"],
      [". Suggest 10 keywords and synonyms I can search in IBISWorld or Statista", "ACTION"],
      [" as a table: keyword | why it helps", "FORMAT"],
      [". Plain and practical", "TONE"],
      [".\"", null]
    ]
  },
  {
    step: "03",
    label: "MAKE SENSE OF A SOURCE",
    parts: [
      ["\"Act as a teaching assistant", "ROLE"],
      [". Below is a section of an IBISWorld report I pasted in", "CONTEXT"],
      [". Explain the 3 most important takeaways for a new business", "ACTION"],
      [" in 3 short bullet points", "FORMAT"],
      [". Plain language, no jargon", "TONE"],
      [".\"", null]
    ]
  }
];

starterPrompts.forEach((p, idx) => {
  const yPos = 1.6 + (idx * 1.6);
  slide5b.addShape(pptx.shapes.RECTANGLE, {
    x: 0.8, y: yPos, w: 11.7, h: 1.4,
    fill: { color: COLORS.cardBg }, line: { color: idx === 2 ? COLORS.accent : COLORS.border, width: idx === 2 ? 1.5 : 1 }
  });
  slide5b.addText(p.step, {
    x: 1.1, y: yPos + 0.2, w: 2.2, h: 0.4,
    fontFace: FONTS.header, fontSize: 20, bold: true, color: COLORS.accentAlt
  });
  slide5b.addText(p.label, {
    x: 1.1, y: yPos + 0.65, w: 2.2, h: 0.6,
    fontFace: FONTS.body, fontSize: 12, bold: true, color: COLORS.textMuted, charSpacing: 1, valign: "top"
  });
  slide5b.addText(craftRuns(p.parts), {
    x: 3.4, y: yPos + 0.15, w: 8.8, h: 1.1,
    fontFace: FONTS.body, fontSize: 14, lineSpacing: 20, valign: "middle"
  });
});

slide5b.addText("Prompt 03 is the habit to build: paste the library source in first, then ask.", {
  x: 0.8, y: 6.4, w: 11.7, h: 0.4,
  fontFace: FONTS.body, fontSize: 13, italic: true, color: COLORS.textMuted, align: "center"
});

addFooter(slide5b);

// -----------------------------------------------------------------------------
// SLIDE 6: High-Signal Deep Tech Signal (Acquired & Core Memory)
// -----------------------------------------------------------------------------
const slide6 = createBaseSlide();
addHeader(slide6, "Beyond the Library: High-Signal Strategy", "Recommended Media");

const mediaCards = [
  {
    title: "ACQUIRED PODCAST",
    sub: "Ben Gilbert & David Rosenthal",
    desc: "3-hour corporate biographies & strategy teardowns (NVIDIA, TSMC, Skunk Works). Learn how technical subsystems scale into super-systems.",
    color: COLORS.accent
  },
  {
    title: "CORE MEMORY",
    sub: "Ashlee Vance",
    desc: "Front-line reporting on hard tech, physical manufacturing, and frontier hardware. Essential perspective on atom-level engineering hurdles.",
    color: COLORS.accentAlt
  }
];

mediaCards.forEach((card, idx) => {
  const xPos = 0.8 + (idx * 6.0);
  slide6.addShape(pptx.shapes.RECTANGLE, {
    x: xPos, y: 1.8, w: 5.7, h: 4.6,
    fill: { color: COLORS.cardBg }, line: { color: card.color, width: 1.5 }
  });
  slide6.addText(card.title, {
    x: xPos + 0.4, y: 2.2, w: 4.9, h: 0.4,
    fontFace: FONTS.header, fontSize: 20, bold: true, color: card.color
  });
  slide6.addText(card.sub, {
    x: xPos + 0.4, y: 2.6, w: 4.9, h: 0.3,
    fontFace: FONTS.body, fontSize: 12, italic: true, color: COLORS.textMuted
  });
  slide6.addText(card.desc, {
    x: xPos + 0.4, y: 3.2, w: 4.9, h: 2.8,
    fontFace: FONTS.body, fontSize: 15, color: COLORS.textMain, lineSpacing: 24, valign: "top"
  });
});

addFooter(slide6);

// -----------------------------------------------------------------------------
// SLIDE 7: The Frozen Close
// -----------------------------------------------------------------------------
const slide7 = createBaseSlide();

slide7.addText("10 MINUTES OF DATA VS. 6 MONTHS IN THE LAB", {
  x: 0.8, y: 1.8, w: 11.5, h: 0.4,
  fontFace: FONTS.body, fontSize: 12, bold: true, color: COLORS.accentAlt, charSpacing: 2, align: "center"
});

slide7.addText("\"10 minutes of database triangulation prevents 6 months of building the wrong physical prototype.\"", {
  x: 1.0, y: 2.4, w: 11.3, h: 2.0,
  fontFace: FONTS.header, fontSize: 32, bold: true, color: COLORS.textMain, align: "center", lineSpacing: 42
});

slide7.addText("Engineer the business model before you solder the board.", {
  x: 0.8, y: 4.6, w: 11.5, h: 0.5,
  fontFace: FONTS.body, fontSize: 20, bold: true, color: COLORS.accent, align: "center"
});

// Closing Contact Card
slide7.addShape(pptx.shapes.RECTANGLE, {
  x: 3.3, y: 5.5, w: 6.7, h: 1.2,
  fill: { color: COLORS.cardBg }, line: { color: COLORS.border, width: 1 }
});
slide7.addText("Terence O'Neill | Entrepreneurship Librarian", {
  x: 3.3, y: 5.7, w: 6.7, h: 0.3,
  fontFace: FONTS.body, fontSize: 13, bold: true, color: COLORS.textMain, align: "center"
});
slide7.addText("MSU Libraries Hub: libguides.lib.msu.edu/entrepreneur", {
  x: 3.3, y: 6.1, w: 6.7, h: 0.3,
  fontFace: FONTS.body, fontSize: 12, color: COLORS.accent, align: "center"
});

// Write Out Presentation File
pptx.writeFile({ fileName: "EGR_440_Scalable_Innovation.pptx" })
  .then((fileName) => {
    console.log(`Successfully generated deck: ${fileName}`);
  })
  .catch((err) => {
    console.error("Error generating deck:", err);
  });
