import pptx
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor

def generate_deck():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # Theme Palette
    DARK_NAVY = RGBColor(15, 23, 42)      # #0F172A
    FOREST_GREEN = RGBColor(20, 83, 45)   # #14532D
    SLATE_GRAY = RGBColor(71, 85, 105)   # #475569
    LIGHT_BG = RGBColor(248, 250, 252)   # #F8FAFC
    CARD_BG = RGBColor(255, 255, 255)    # #FFFFFF
    WHITE = RGBColor(255, 255, 255)
    ACCENT_GREEN = RGBColor(34, 197, 94)  # #22C55E
    BORDER_GRAY = RGBColor(226, 232, 240) # #E2E8F0

    def add_slide_background(slide):
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
        bg.fill.solid()
        bg.fill.fore_color.rgb = LIGHT_BG
        bg.line.fill.background()

    def add_header(slide, title_text, category_text="PRODUCT CENTER PARTNERSHIP"):
        header_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.35), Inches(11.733), Inches(1.1))
        tf = header_box.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

        p0 = tf.paragraphs[0]
        p0.text = category_text.upper()
        p0.font.name = "Calibri"
        p0.font.size = Pt(11)
        p0.font.bold = True
        p0.font.color.rgb = FOREST_GREEN

        p1 = tf.add_paragraph()
        p1.text = title_text
        p1.font.name = "Cambria"
        p1.font.size = Pt(20)
        p1.font.bold = True
        p1.font.color.rgb = DARK_NAVY

    def add_card(slide, left, top, width, height, title, body_bullets, accent_title=False):
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
        card.fill.solid()
        card.fill.fore_color.rgb = CARD_BG
        card.line.color.rgb = BORDER_GRAY
        card.line.width = Pt(1)
        card.adjustments[0] = 0.04
        card.shadow.inherit = False

        tf = card.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.TOP
        tf.margin_left = tf.margin_right = Inches(0.22)
        tf.margin_top = tf.margin_bottom = Inches(0.2)

        p0 = tf.paragraphs[0]
        p0.text = title
        p0.alignment = PP_ALIGN.LEFT
        p0.font.name = "Cambria"
        p0.font.size = Pt(15)
        p0.font.bold = True
        p0.font.color.rgb = FOREST_GREEN if accent_title else DARK_NAVY
        p0.space_after = Pt(10)

        for bullet in body_bullets:
            p = tf.add_paragraph()
            p.alignment = PP_ALIGN.LEFT
            run = p.add_run()
            run.text = bullet
            run.font.name = "Calibri"
            run.font.size = Pt(11.5)
            run.font.color.rgb = SLATE_GRAY
            p.space_after = Pt(8)
            p.line_spacing = 1.15
            _set_bullet_char(p, "•", FOREST_GREEN)
            p.level = 0

    # SLIDE 1: Title
    slide1 = prs.slides.add_slide(blank_layout)
    bg1 = slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg1.fill.solid()
    bg1.fill.fore_color.rgb = DARK_NAVY
    bg1.line.fill.background()

    tf1 = slide1.shapes.add_textbox(Inches(1.0), Inches(2.35), Inches(11.333), Inches(4.5)).text_frame
    tf1.word_wrap = True

    p = tf1.paragraphs[0]
    p.text = "PRODUCT CENTER CONSULTATIVE BRIEFING"
    p.font.name = "Calibri"
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = ACCENT_GREEN
    p.space_after = Pt(12)

    p = tf1.add_paragraph()
    p.text = "Triangulation from Imperfect Sources"
    p.font.name = "Cambria"
    p.font.size = Pt(38)
    p.font.bold = True
    p.font.color.rgb = WHITE
    p.space_after = Pt(8)

    p = tf1.add_paragraph()
    p.text = "An Evidence-Based Information Strategy for CPG Venture Validation"
    p.font.name = "Calibri"
    p.font.size = Pt(20)
    p.font.color.rgb = RGBColor(203, 213, 225)
    p.space_after = Pt(28)

    p = tf1.add_paragraph()
    p.text = '"Intuition sparks the venture; library evidence proves it."'
    p.font.name = "Cambria"
    p.font.size = Pt(16)
    p.font.italic = True
    p.font.color.rgb = ACCENT_GREEN

    # SLIDE 2: Compliant Commercial Stack
    slide2 = prs.slides.add_slide(blank_layout)
    add_slide_background(slide2)
    add_header(slide2, "The Compliant Commercial Stack")

    add_card(slide2, Inches(0.8), Inches(1.75), Inches(3.64), Inches(5.0),
             "1. Michigan eLibrary (MeL)",
             ["AtoZdatabases: B2B sales lists, NAICS/SIC industry code filtering, and competitor radius mapping.",
              "DemographicsNow: Local demographic profiles, household income metrics, and geographic heatmaps.",
              "100% state-funded and commercially compliant for Michigan businesses."])

    add_card(slide2, Inches(4.84), Inches(1.75), Inches(3.64), Inches(5.0),
             "2. Public & Gov Repositories",
             ["US Census Bureau & BEA: Baseline population ratios and regional macroeconomic metrics.",
              "USDA & FDA: Official ingredient standards, regulatory guidelines, and nutrition labeling rules.",
              "Public domain datasets: Legally unassailable and fully accessible to clients permanently."])

    add_card(slide2, Inches(8.88), Inches(1.75), Inches(3.64), Inches(5.0),
             "3. Trade & Open Web",
             ["Industry Associations: High-level market figures (e.g., Outdoor Industry Assoc., National Confectioners Assoc.).",
              "Specialty Grocers & Market Stubs: Benchmark pricing from outlets like Mouth.com or FairTrade America.",
              "Academic tools note: IBISWorld/Mintel sit behind strict educational walls and are not used for third-party client hand-offs."],
             accent_title=True)

    caption2 = slide2.shapes.add_textbox(Inches(0.8), Inches(6.85), Inches(11.733), Inches(0.35)).text_frame
    _set_caption(caption2, "Navigating licensing boundaries across three compliant source tiers", SLATE_GRAY)

    # SLIDE 3: Tactical Moves & Proxy Modeling
    slide3 = prs.slides.add_slide(blank_layout)
    add_slide_background(slide3)
    add_header(slide3, "Core Tactical Moves: Triangulating Imperfect Data")

    add_card(slide3, Inches(0.8), Inches(1.75), Inches(3.64), Inches(5.0),
             "Move 1: Pivot to Umbrella Drivers",
             ["Challenge: Niche concepts (e.g., local vegan cookies) lack direct market reports.",
              "Tactical move: Zoom out to macro consumer shifts.",
              "Execution: Benchmark against flexitarianism (36% of US consumers) and Gen Z/Millennial ESG spending trends."])

    add_card(slide3, Inches(4.84), Inches(1.75), Inches(3.64), Inches(5.0),
             "Move 2: Apply Proxy Modeling",
             ["Challenge: Early-stage food ventures have no budget for $5,000 custom market data.",
              "Tactical move: Calculate local baselines using national trade totals.",
              "Execution: Multiply national camping/specialty food totals (OIA, KOA) by Michigan's ~3% population ratio."])

    add_card(slide3, Inches(8.88), Inches(1.75), Inches(3.64), Inches(5.0),
             "Move 3: Cross-Category Matching",
             ["Challenge: Standard commodity pricing creates margin compression.",
              "Tactical move: Borrow positioning tactics from adjacent high-margin industries.",
              "Execution: Elevate premium maple syrup with craft-spirit-style packaging, barrel-aging narratives, and tasting notes."])

    # SLIDE 4: Targeted AI (CRAFT) & Deliverables
    slide4 = prs.slides.add_slide(blank_layout)
    add_slide_background(slide4)
    add_header(slide4, "Targeted AI (CRAFT Framework) & Permissible Deliverables")

    add_card(slide4, Inches(0.8), Inches(1.75), Inches(5.66), Inches(5.0),
             "The CRAFT Prompting Engine for GenAI",
             ["Context: Ground the AI in venture specifics (e.g., Native-owned beef jerky / Tanka Bar).",
              "Role: Assign a specific mentor persona (e.g., CPG market expansion advisor).",
              "Action: Define an explicit synthesis task (e.g., benchmark values-based premium pricing).",
              "Format: Instruct output in structured bullet points with explicit source citations.",
              "Tone: Require direct, plain English without marketing fluff.",
              "Essential rule: Instruct the AI to flag missing data and grade its own confidence level."])

    add_card(slide4, Inches(6.86), Inches(1.75), Inches(5.66), Inches(5.0),
             "Permissible Client Deliverables",
             ["Curated open-web research folders: Organized links to government repositories and trade press.",
              "Synthesized market summaries: Custom executive briefs highlighting macro trends and proxy math.",
              "Targeted search operators: Ready-to-use search strings (e.g., site:.gov, filetype:pdf).",
              "AtoZ lead spreadsheets: Full B2B prospect lists, competitor locations, and contact info from MeL.",
              "10% excerpt limit: Direct quotes or data from proprietary sources are strictly capped at 10%."],
             accent_title=True)

    # SLIDE 5: Counselor Hand-off & Workflow
    slide5 = prs.slides.add_slide(blank_layout)
    add_slide_background(slide5)
    add_header(slide5, "The Counselor Hand-Off & Triad Workflow")

    add_card(slide5, Inches(0.8), Inches(1.75), Inches(3.64), Inches(5.0),
             "1. Counselor Intake",
             ["Frames the venture challenge with the client during initial meetings.",
              "Applies the three-part compliant stack (MeL, gov data, trade press) directly.",
              "Identifies complex data gaps requiring strategic search design."])

    add_card(slide5, Inches(4.84), Inches(1.75), Inches(3.64), Inches(5.0),
             "2. Librarian Partner",
             ["Co-designs search strategies for non-obvious or hard-to-search niches.",
              "Constructs targeted CRAFT AI prompts and identifies valid proxy datasets.",
              "Guides counselors and entrepreneurs through pattern matching."],
             accent_title=True)

    add_card(slide5, Inches(8.88), Inches(1.75), Inches(3.64), Inches(5.0),
             "3. Entrepreneur Ownership",
             ["Evaluates source credibility and cross-checks assumptions.",
              "Calculates unit economics, margin structures, and proxy market sizes.",
              "Owns final financial projections and pitch deck validation."])

    prs.save("Product_Center_Presentation.pptx")


def _set_bullet_char(paragraph, char, color):
    """Attach a real PowerPoint bullet (buChar) to a paragraph instead of a literal glyph in the text."""
    from pptx.oxml.ns import qn
    import copy
    pPr = paragraph._pPr
    if pPr is None:
        pPr = paragraph._p.get_or_add_pPr()
    for tag in ("a:buNone", "a:buChar", "a:buAutoNum"):
        el = pPr.find(qn(tag))
        if el is not None:
            pPr.remove(el)
    buFont = pPr.makeelement(qn("a:buFont"), {"typeface": "Arial"})
    buChar = pPr.makeelement(qn("a:buChar"), {"char": char})
    pPr.append(buFont)
    pPr.append(buChar)
    pPr.set("indent", "-137160")
    pPr.set("marL", "137160")


def _set_caption(text_frame, text, color):
    text_frame.word_wrap = True
    p = text_frame.paragraphs[0]
    p.text = text
    p.font.name = "Calibri"
    p.font.size = Pt(11)
    p.font.italic = True
    p.font.color.rgb = color


if __name__ == "__main__":
    generate_deck()
