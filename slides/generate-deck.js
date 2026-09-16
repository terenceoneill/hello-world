const pptxgen = require("pptxgenjs");

const pptx = new pptxgen();
pptx.layout = "LAYOUT_WIDE"; // 16:9 Widescreen (13.33 x 7.5 inches)

// Color Palette (Strict 6-digit hex strings without #)
const COLOR_BG = "F8F9FA";         // Off-white background
const COLOR_CARD = "FFFFFF";       // White card container
const COLOR_TEXT = "2D3748";       // Slate body text
const COLOR_TITLE = "1A202C";      // Dark header text
const COLOR_MSU_GREEN = "18453B"; // Official MSU Green accent
const COLOR_MUTED = "718096";      // Muted gray text

// Standard Footer Helper
function addMSUFooter(slide) {
  slide.addText("MSU Libraries  |  Terence O'Neill, Entrepreneurship Librarian (oneillt@msu.edu)", {
    x: 0.8, y: 7.0, w: 11.7, h: 0.4,
    fontFace: "Calibri", fontSize: 12, color: COLOR_MUTED, isTextBox: true
  });
}

// ------------------------------------------------------------------------
// SLIDE 1: Cover Page & ESHP 491.003 Course Announcement
// ------------------------------------------------------------------------
let slide1 = pptx.addSlide();
slide1.background = { fill: COLOR_MSU_GREEN };

slide1.addText("ESHP 190: Introduction to Entrepreneurship", {
  x: 0.8, y: 0.5, w: 11.7, h: 0.7,
  fontFace: "Cambria", fontSize: 32, bold: true, color: "FFFFFF", isTextBox: true
});

slide1.addText("Terence O'Neill | Entrepreneurship Librarian | oneillt@msu.edu", {
  x: 0.8, y: 1.2, w: 11.7, h: 0.4,
  fontFace: "Calibri", fontSize: 18, color: "E2E8F0", isTextBox: true
});

// Course Announcement Card
slide1.addShape("rect", {
  x: 0.8, y: 1.8, w: 11.7, h: 5.0,
  fill: { color: COLOR_CARD }, line: { color: "CBD5E0", width: 1 }
});

slide1.addText("FEATURED COURSE: SPRING SEMESTER 2027", {
  x: 1.2, y: 2.1, w: 10.9, h: 0.4,
  fontFace: "Cambria", fontSize: 16, bold: true, color: COLOR_MSU_GREEN, isTextBox: true
});

slide1.addText("ESHP 491.003: Researching Your Venture: Customers, Markets, and Context", {
  x: 1.2, y: 2.5, w: 10.9, h: 0.5,
  fontFace: "Cambria", fontSize: 22, bold: true, color: COLOR_TITLE, isTextBox: true
});

slide1.addText("Wednesdays, 4:00 PM – 5:20 PM | Main Library", {
  x: 1.2, y: 3.1, w: 10.9, h: 0.4,
  fontFace: "Calibri", fontSize: 16, bold: true, color: COLOR_MUTED, isTextBox: true
});

slide1.addText(
  "“This course teaches you to research the three domains that shape every venture: customers, markets, and business environment. " +
  "You will conduct primary customer research, use databases to analyze industries and competitors, and scan environmental forces.\n\n" +
  "No prior business coursework required. Come with a venture idea, at any stage, that you are willing to research honestly, " +
  "including the possibility that your assumptions need to change.”",
  {
    x: 1.2, y: 3.65, w: 10.9, h: 2.9,
    fontFace: "Calibri", fontSize: 15, color: COLOR_TEXT, lineSpacingMultiple: 1.3, isTextBox: true
  }
);


// ------------------------------------------------------------------------
// SLIDE 2: The Hook - Japan Hotel Smoothie Scenario
// ------------------------------------------------------------------------
let slide2 = pptx.addSlide();
slide2.background = { fill: COLOR_BG };
addMSUFooter(slide2);

slide2.addText("The Japan Hotel Smoothie: From Intuition to Evidence", {
  x: 0.8, y: 0.5, w: 11.7, h: 0.7,
  fontFace: "Cambria", fontSize: 30, bold: true, color: COLOR_TITLE, isTextBox: true
});

slide2.addShape("rect", {
  x: 0.8, y: 1.5, w: 6.8, h: 5.2,
  fill: { color: COLOR_CARD }, line: { color: "CBD5E0", width: 1 }
});

slide2.addText(
  [
    { text: "Scenario: Fictional Event Venue Concept", options: { fontFace: "Cambria", fontSize: 18, bold: true, color: COLOR_MSU_GREEN, breakLine: true } },
    { text: "“While traveling in Japan, I had an incredible smoothie every morning at my hotel. What if we made a signature smoothie bar a feature of our event venue—a twist on wellness weddings?”", options: { breakLine: true } },
    { text: "Questions that immediately arise:", options: { bold: true, breakLine: true } },
    { text: "Are ‘wellness weddings’ actually trending, or a passing niche?", options: { bullet: true, breakLine: true } },
    { text: "What functional ingredients/flavors are consumers demanding?", options: { bullet: true, breakLine: true } },
    { text: "Can local growers lock down supply for these agricultural inputs?", options: { bullet: true } }
  ],
  {
    x: 1.1, y: 1.8, w: 6.2, h: 4.7,
    fontFace: "Calibri", fontSize: 15, color: COLOR_TEXT, lineSpacingMultiple: 1.2,
    paraSpaceAfter: 12, isTextBox: true
  }
);

slide2.addShape("rect", {
  x: 7.9, y: 1.5, w: 4.6, h: 5.2,
  fill: { color: "EDF2F7" }, line: { color: "CBD5E0", width: 1 }
});

slide2.addText("[ PHOTO PLACEHOLDER ]\n\nJapanese Hotel Smoothie / Wellness Concept", {
  x: 8.1, y: 3.7, w: 4.2, h: 1.2,
  fontFace: "Calibri", fontSize: 16, color: COLOR_MUTED, align: "center", isTextBox: true
});


// ------------------------------------------------------------------------
// SLIDE 3: Setup & Research Hygiene (Visual Anchor: Flossing)
// ------------------------------------------------------------------------
let slide3 = pptx.addSlide();
slide3.background = { fill: COLOR_BG };
addMSUFooter(slide3);

slide3.addText("Workflow Setup & Research Hygiene", {
  x: 0.8, y: 0.5, w: 11.7, h: 0.7,
  fontFace: "Cambria", fontSize: 30, bold: true, color: COLOR_TITLE, isTextBox: true
});

slide3.addShape("rect", {
  x: 0.8, y: 1.5, w: 6.8, h: 5.2,
  fill: { color: COLOR_CARD }, line: { color: COLOR_MSU_GREEN, width: 2 }
});

slide3.addText(
  [
    { text: "Step 1: Open a digital document right now", options: { fontFace: "Cambria", fontSize: 18, bold: true, color: COLOR_TITLE, breakLine: true } },
    { text: "Write in your document right now:", options: { breakLine: true } },
    { text: "“What specific questions do I have about my venture or client’s market right now?”", options: { italic: true, breakLine: true } },
    { text: "Core Practices Embedded Here:", options: { bold: true, breakLine: true } },
    { text: "Research Hygiene: Routine habits (like daily flossing) keep your project organized.", options: { bullet: true, breakLine: true } },
    { text: "Traceability: Give yourself a single place to log links, notes, and citations.", options: { bullet: true, breakLine: true } },
    { text: "Collaboration: How will your team share these notes?", options: { bullet: true, breakLine: true } },
    { text: "Mindset Shift: Remember—you are NOT your customer.", options: { bullet: true } }
  ],
  {
    x: 1.1, y: 1.8, w: 6.2, h: 4.7,
    fontFace: "Calibri", fontSize: 15, color: COLOR_TEXT, lineSpacingMultiple: 1.2,
    paraSpaceAfter: 12, isTextBox: true
  }
);

// Hygiene Visual Anchor Box
slide3.addShape("rect", {
  x: 7.9, y: 1.5, w: 4.6, h: 5.2,
  fill: { color: "EDF2F7" }, line: { color: "CBD5E0", width: 1 }
});

slide3.addText("[ PHOTO PLACEHOLDER ]\n\nDaily Habits & Clean Research", {
  x: 8.1, y: 3.7, w: 4.2, h: 1.2,
  fontFace: "Calibri", fontSize: 16, color: COLOR_MUTED, align: "center", isTextBox: true
});


// ------------------------------------------------------------------------
// SLIDE 4: Database Deep-Dive & Hub
// ------------------------------------------------------------------------
let slide4 = pptx.addSlide();
slide4.background = { fill: COLOR_BG };
addMSUFooter(slide4);

slide4.addText("Database Deep-Dive & Exploration", {
  x: 0.8, y: 0.5, w: 11.7, h: 0.7,
  fontFace: "Cambria", fontSize: 30, bold: true, color: COLOR_TITLE, isTextBox: true
});

slide4.addShape("rect", {
  x: 0.8, y: 1.5, w: 11.7, h: 5.2,
  fill: { color: COLOR_CARD }, line: { color: "CBD5E0", width: 1 }
});

slide4.addText("Entrepreneurship Hub: libguides.lib.msu.edu/entrepreneur", {
  x: 1.2, y: 1.8, w: 10.9, h: 0.4,
  fontFace: "Cambria", fontSize: 19, bold: true, color: COLOR_MSU_GREEN, isTextBox: true
});

slide4.addText(
  [
    { text: "Mintel: ", options: { bold: true, bullet: true, breakLine: false } },
    { text: "Consumer trends, demographic behavior, and market forecasts (check login on entry!).", options: { breakLine: true } },
    { text: "Statista: ", options: { bold: true, bullet: true, breakLine: false } },
    { text: "Visualized statistics, consumer surveys, and market size reports.", options: { breakLine: true } },
    { text: "Data Axle Reference Solutions: ", options: { bold: true, bullet: true, breakLine: false } },
    { text: "Map local competitors by zip code, pull private company profiles, and build target customer lists.", options: { breakLine: true } },
    { text: "Technical Access Note: ", options: { bold: true, bullet: true, breakLine: false } },
    { text: "If hit by a login or proxy loop, switch to another browser (e.g., Chrome to Firefox/Incognito) to bypass cookie issues.", options: {} }
  ],
  {
    x: 1.2, y: 2.4, w: 10.9, h: 4.0,
    fontFace: "Calibri", fontSize: 16, color: COLOR_TEXT, lineSpacingMultiple: 1.3,
    paraSpaceAfter: 14, isTextBox: true
  }
);


// ------------------------------------------------------------------------
// SLIDE 5: Database AI & CRAFT Framework
// ------------------------------------------------------------------------
let slide5 = pptx.addSlide();
slide5.background = { fill: COLOR_BG };
addMSUFooter(slide5);

slide5.addText("Database AI & The CRAFT Prompting Framework", {
  x: 0.8, y: 0.5, w: 11.7, h: 0.7,
  fontFace: "Cambria", fontSize: 28, bold: true, color: COLOR_TITLE, isTextBox: true
});

// Left Column: Principles & CRAFT Definition
slide5.addShape("rect", {
  x: 0.8, y: 1.4, w: 5.7, h: 5.3,
  fill: { color: COLOR_CARD }, line: { color: "CBD5E0", width: 1 }
});

slide5.addText(
  [
    { text: "“Think of a generative AI tool as a machine you are programming with words.”", options: { fontFace: "Cambria", italic: true, color: COLOR_MSU_GREEN, breakLine: true } },
    { text: "— Ethan Mollick", options: { fontFace: "Cambria", italic: true, color: COLOR_MSU_GREEN, breakLine: true } },
    { text: "CRAFT Mnemonic:", options: { bold: true, breakLine: true } },
    { text: "Context | Role | Action | Format | Tone", options: { breakLine: true } },
    { text: "Two Advantages of Library Database AIs:", options: { bold: true, breakLine: true } },
    { text: "Primed specifically for business/entrepreneurial research.", options: { bullet: true, breakLine: true } },
    { text: "Grounded in proprietary data inaccessible to open web LLMs.", options: { bullet: true, breakLine: true } },
    { text: "Rule of Thumb: Ask AI to grade its confidence level and reveal its sources.", options: { italic: true } }
  ],
  {
    x: 1.0, y: 1.6, w: 5.3, h: 4.9,
    fontFace: "Calibri", fontSize: 14, color: COLOR_TEXT, lineSpacingMultiple: 1.2,
    paraSpaceAfter: 10, isTextBox: true
  }
);

// Right Column: CRAFT Example Prompt
slide5.addShape("rect", {
  x: 6.8, y: 1.4, w: 5.7, h: 5.3,
  fill: { color: "EDF2F7" }, line: { color: COLOR_MSU_GREEN, width: 2 }
});

slide5.addText("Example CRAFT Prompt (Try in Mintel/Statista):", {
  x: 7.0, y: 1.6, w: 5.3, h: 0.4,
  fontFace: "Cambria", fontSize: 15, bold: true, color: COLOR_TITLE, isTextBox: true
});

slide5.addText(
  "“Please act as my experienced food business mentor [ROLE] coaching me as I search for a business model in my chosen area of developing a local food smoothie system around the broader concept of a ‘wellness wedding’ and similar retreats [CONTEXT].\n\n" +
  "Please present five trends [ACTION] in bullets with citations [FORMAT]. Keep the tone conversational and explain any acronyms [TONE].”",
  {
    x: 7.0, y: 2.1, w: 5.3, h: 4.4,
    fontFace: "Calibri", fontSize: 14, color: COLOR_TEXT, lineSpacingMultiple: 1.3, isTextBox: true
  }
);


// Save Deck
pptx.writeFile({ fileName: "ESHP190_Library_Session_Master.pptx" })
  .then((file) => console.log(`Deck successfully generated: ${file}`))
  .catch((err) => console.error("Generation error:", err));
