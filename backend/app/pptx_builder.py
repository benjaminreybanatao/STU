"""
Builds the two-page (or three-page, if a downside case is present) offering
summary deck.

No reference firm template was supplied for this build, so the layout below
is a clean, generic institutional two-pager: navy header band, gold accent
rule, a hero photo pulled from the OM, a woven narrative paragraph, and a
native (editable) return-summary table. If a real firm deck is provided
later, this module is the one place to reverse-engineer its exact fonts,
colors, logo placement, and table style into -- the rest of the pipeline
(parsing, gap analysis, field resolution) does not need to change.

Every value placed on the slide comes from Deal.display_value(), which
walks the OM -> model -> analyst provenance chain and falls back to the
PLACEHOLDER string. Nothing is fabricated here.
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

from .schema import Deal, FIELD_BY_KEY, PLACEHOLDER

NAVY = RGBColor(0x14, 0x24, 0x40)
GOLD = RGBColor(0xB8, 0x8A, 0x2E)
SLATE = RGBColor(0x3B, 0x44, 0x54)
LIGHT_GRAY = RGBColor(0xF1, 0xF2, 0xF4)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
PLACEHOLDER_RED = RGBColor(0xB0, 0x2A, 0x2A)

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)

FONT = "Calibri"

RETURN_TABLE_KEYS = [
    "purchase_price",
    "price_per_unit",
    "going_in_cap",
    "exit_cap",
    "leverage",
    "debt_rate",
    "initial_equity",
    "peak_equity",
    "hold_period",
    "unlevered_irr",
    "levered_irr",
    "unlevered_em",
    "levered_em",
    "cash_on_cash",
]


def _set_cell_text(cell, text, *, bold=False, size=12, color=SLATE, align=PP_ALIGN.LEFT, font=FONT):
    cell.text = str(text)
    p = cell.text_frame.paragraphs[0]
    p.alignment = align
    run = p.runs[0]
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.name = font
    run.font.color.rgb = color


def _add_header_band(slide, title_text, subtitle_text=""):
    band = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, Inches(1.05))
    band.fill.solid()
    band.fill.fore_color.rgb = NAVY
    band.line.fill.background()
    band.shadow.inherit = False

    rule = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Inches(1.05), SLIDE_W, Inches(0.06))
    rule.fill.solid()
    rule.fill.fore_color.rgb = GOLD
    rule.line.fill.background()
    rule.shadow.inherit = False

    title_box = slide.shapes.add_textbox(Inches(0.45), Inches(0.14), Inches(9.5), Inches(0.5))
    tf = title_box.text_frame
    tf.text = title_text
    p = tf.paragraphs[0]
    p.runs[0].font.size = Pt(24)
    p.runs[0].font.bold = True
    p.runs[0].font.name = FONT
    p.runs[0].font.color.rgb = WHITE

    if subtitle_text:
        sub_box = slide.shapes.add_textbox(Inches(0.45), Inches(0.62), Inches(9.5), Inches(0.35))
        tf2 = sub_box.text_frame
        tf2.text = subtitle_text
        p2 = tf2.paragraphs[0]
        p2.runs[0].font.size = Pt(13)
        p2.runs[0].font.name = FONT
        p2.runs[0].font.color.rgb = RGBColor(0xC9, 0xCE, 0xDA)

    logo_box = slide.shapes.add_textbox(Inches(10.3), Inches(0.3), Inches(2.6), Inches(0.5))
    tf3 = logo_box.text_frame
    tf3.text = "[FIRM LOGO]"
    p3 = tf3.paragraphs[0]
    p3.alignment = PP_ALIGN.RIGHT
    p3.runs[0].font.size = Pt(12)
    p3.runs[0].font.italic = True
    p3.runs[0].font.name = FONT
    p3.runs[0].font.color.rgb = RGBColor(0x9A, 0xA3, 0xB5)


def _build_narrative(deal: Deal) -> str:
    dv = deal.display_value

    address = dv("address")
    ptype = dv("property_type")
    year_built = dv("year_built")
    size = dv("sf_or_units")
    submarket = dv("submarket_desc")
    tenant_summary = dv("tenant_summary")
    occupancy = dv("occupancy")
    price = dv("purchase_price")
    ppu = dv("price_per_unit")
    cap = dv("going_in_cap")
    exit_cap = dv("exit_cap")
    leverage = dv("leverage")
    debt_rate = dv("debt_rate")
    initial_equity = dv("initial_equity")
    peak_equity = dv("peak_equity")
    hold = dv("hold_period")
    unl_irr = dv("unlevered_irr")
    lev_irr = dv("levered_irr")
    unl_em = dv("unlevered_em")
    lev_em = dv("levered_em")

    sentences = []

    sentences.append(
        f"{address} is a {size} {ptype.lower() if ptype != PLACEHOLDER else ptype} asset "
        f"built in {year_built}."
    )
    if submarket != PLACEHOLDER:
        sentences.append(submarket)

    sentences.append(
        f"The property is {occupancy} occupied; {tenant_summary}"
        if tenant_summary != PLACEHOLDER
        else f"The property is {occupancy} occupied."
    )

    sentences.append(
        f"The Sponsor is underwriting an acquisition at {price} ({ppu}), reflecting a "
        f"{cap} going-in cap rate against a projected {exit_cap} exit cap rate at disposition."
    )

    sentences.append(
        f"The capital structure assumes {leverage} leverage at a {debt_rate} debt rate, "
        f"with {initial_equity} of initial equity"
        + (f" and {peak_equity} at peak funding" if peak_equity != PLACEHOLDER else "")
        + f" over a {hold} hold period."
    )

    sentences.append(
        f"On this basis, the model projects a {unl_irr} unlevered IRR ({unl_em} unlevered "
        f"equity multiple) and a {lev_irr} levered IRR ({lev_em} levered equity multiple)."
    )

    return " ".join(sentences)


def _render_narrative_textbox(slide, narrative: str, left, top, width, height):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    tf.text = ""
    p = tf.paragraphs[0]

    # Split on the placeholder string so it can be styled distinctly (bold red)
    # while everything else renders as normal body copy -- makes missing data
    # visually obvious on the deck itself, per the "no fabrication" requirement.
    parts = narrative.split(PLACEHOLDER)
    for i, part in enumerate(parts):
        if part:
            run = p.add_run()
            run.text = part
            run.font.size = Pt(13)
            run.font.name = FONT
            run.font.color.rgb = SLATE
        if i < len(parts) - 1:
            run = p.add_run()
            run.text = PLACEHOLDER
            run.font.size = Pt(13)
            run.font.bold = True
            run.font.name = FONT
            run.font.color.rgb = PLACEHOLDER_RED
    return box


def _add_kpi_strip(slide, deal: Deal, left, top, width):
    kpis = [
        ("Purchase Price", deal.display_value("purchase_price")),
        ("Going-In Cap", deal.display_value("going_in_cap")),
        ("Leverage", deal.display_value("leverage")),
        ("Levered IRR", deal.display_value("levered_irr")),
        ("Levered EM", deal.display_value("levered_em")),
    ]
    n = len(kpis)
    cell_w = Emu(int(width / n))
    for i, (label, value) in enumerate(kpis):
        x = Emu(int(left) + i * int(cell_w))
        box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, top, cell_w - Emu(Inches(0.05)), Inches(0.95))
        box.fill.solid()
        box.fill.fore_color.rgb = LIGHT_GRAY
        box.line.color.rgb = RGBColor(0xDD, 0xDE, 0xE2)
        box.line.width = Pt(0.75)
        box.shadow.inherit = False
        tf = box.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.margin_left = Pt(6)
        tf.margin_right = Pt(6)
        p0 = tf.paragraphs[0]
        p0.alignment = PP_ALIGN.CENTER
        run0 = p0.add_run()
        run0.text = value
        run0.font.size = Pt(16)
        run0.font.bold = True
        run0.font.name = FONT
        run0.font.color.rgb = PLACEHOLDER_RED if value == PLACEHOLDER else NAVY

        p1 = tf.add_paragraph()
        p1.alignment = PP_ALIGN.CENTER
        run1 = p1.add_run()
        run1.text = label
        run1.font.size = Pt(9.5)
        run1.font.name = FONT
        run1.font.color.rgb = SLATE


def _add_hero_photo(slide, image_path, left, top, width, height):
    if not image_path:
        placeholder = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
        placeholder.fill.solid()
        placeholder.fill.fore_color.rgb = RGBColor(0xE4, 0xE5, 0xE8)
        placeholder.line.color.rgb = RGBColor(0xC7, 0xC9, 0xCE)
        tf = placeholder.text_frame
        tf.text = "No property photo extracted from OM"
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        p.runs[0].font.italic = True
        p.runs[0].font.color.rgb = RGBColor(0x8A, 0x8D, 0x94)
        return

    pic = slide.shapes.add_picture(image_path, left, top, width=width, height=height)
    # Crop to fill the frame without distortion (center-crop on the long axis).
    native_w, native_h = pic.image.size
    target_ratio = width / height
    native_ratio = native_w / native_h
    if native_ratio > target_ratio:
        crop = (1 - target_ratio / native_ratio) / 2
        pic.crop_left = crop
        pic.crop_right = crop
    else:
        crop = (1 - native_ratio / target_ratio) / 2
        pic.crop_top = crop
        pic.crop_bottom = crop


def _build_slide1(prs, deal: Deal, hero_image_path):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _add_header_band(slide, "Transaction Summary", deal.display_value("address"))

    photo_left, photo_top = Inches(0.45), Inches(1.35)
    photo_w, photo_h = Inches(5.4), Inches(3.55)
    _add_hero_photo(slide, hero_image_path, photo_left, photo_top, photo_w, photo_h)

    narrative_left = Inches(6.1)
    narrative_top = Inches(1.35)
    narrative_w = Inches(6.8)
    narrative_h = Inches(3.55)
    narrative = _build_narrative(deal)
    _render_narrative_textbox(slide, narrative, narrative_left, narrative_top, narrative_w, narrative_h)

    _add_kpi_strip(slide, deal, Inches(0.45), Inches(5.15), Inches(12.44))

    footer = slide.shapes.add_textbox(Inches(0.45), Inches(7.05), Inches(12.4), Inches(0.35))
    tf = footer.text_frame
    tf.text = "Confidential -- for internal investment committee discussion purposes only."
    p = tf.paragraphs[0]
    p.runs[0].font.size = Pt(9)
    p.runs[0].font.italic = True
    p.runs[0].font.name = FONT
    p.runs[0].font.color.rgb = RGBColor(0x9A, 0xA3, 0xB5)
    return slide


def _add_return_table(slide, deal: Deal, left, top, width, height):
    rows = len(RETURN_TABLE_KEYS) + 1
    cols = 2
    table_shape = slide.shapes.add_table(rows, cols, left, top, width, height)
    table = table_shape.table
    table.columns[0].width = Emu(int(width * 0.62))
    table.columns[1].width = Emu(int(width * 0.38))

    _set_cell_text(table.cell(0, 0), "Metric", bold=True, size=12, color=WHITE, align=PP_ALIGN.LEFT)
    _set_cell_text(table.cell(0, 1), "Value", bold=True, size=12, color=WHITE, align=PP_ALIGN.RIGHT)
    for col in range(2):
        cell = table.cell(0, col)
        cell.fill.solid()
        cell.fill.fore_color.rgb = NAVY

    for i, key in enumerate(RETURN_TABLE_KEYS, start=1):
        label = FIELD_BY_KEY[key].label
        value = deal.display_value(key)
        row_fill = WHITE if i % 2 else LIGHT_GRAY
        for col, (text, align) in enumerate([(label, PP_ALIGN.LEFT), (value, PP_ALIGN.RIGHT)]):
            cell = table.cell(i, col)
            cell.fill.solid()
            cell.fill.fore_color.rgb = row_fill
            color = PLACEHOLDER_RED if text == PLACEHOLDER else SLATE
            _set_cell_text(cell, text, bold=(col == 1), size=11, color=color, align=align)

    return table_shape


def _add_assumptions_block(slide, deal: Deal, left, top, width, height):
    box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    box.fill.solid()
    box.fill.fore_color.rgb = LIGHT_GRAY
    box.line.color.rgb = RGBColor(0xDD, 0xDE, 0xE2)
    box.shadow.inherit = False

    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = Pt(12)
    tf.margin_right = Pt(12)
    tf.margin_top = Pt(10)

    heading = tf.paragraphs[0]
    run = heading.add_run()
    run.text = "Key Assumptions"
    run.font.bold = True
    run.font.size = Pt(13)
    run.font.name = FONT
    run.font.color.rgb = NAVY

    rows = [
        ("Lease Term", deal.display_value("lease_term_assumption")),
        ("Downtime", deal.display_value("downtime_assumption")),
        ("Exit Assumption", deal.display_value("exit_assumption")),
    ]
    for label, value in rows:
        p = tf.add_paragraph()
        p.space_before = Pt(6)
        r1 = p.add_run()
        r1.text = f"{label}: "
        r1.font.bold = True
        r1.font.size = Pt(11.5)
        r1.font.name = FONT
        r1.font.color.rgb = SLATE
        r2 = p.add_run()
        r2.text = value
        r2.font.size = Pt(11.5)
        r2.font.name = FONT
        r2.font.color.rgb = PLACEHOLDER_RED if value == PLACEHOLDER else SLATE


def _build_return_slide(prs, deal: Deal, slide_title):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _add_header_band(slide, slide_title, deal.display_value("address"))

    _add_return_table(slide, deal, Inches(0.45), Inches(1.35), Inches(7.0), Inches(5.2))
    _add_assumptions_block(slide, deal, Inches(7.75), Inches(1.35), Inches(5.15), Inches(5.2))

    footer = slide.shapes.add_textbox(Inches(0.45), Inches(7.05), Inches(12.4), Inches(0.35))
    tf = footer.text_frame
    tf.text = "Every figure above traces to the OM, the underwriting model, or an explicit analyst input -- nothing is estimated."
    p = tf.paragraphs[0]
    p.runs[0].font.size = Pt(9)
    p.runs[0].font.italic = True
    p.runs[0].font.name = FONT
    p.runs[0].font.color.rgb = RGBColor(0x9A, 0xA3, 0xB5)
    return slide


def build_deck(deal: Deal, hero_image_path: str | None, output_path: str, downside_deal: Deal | None = None):
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H

    _build_slide1(prs, deal, hero_image_path)
    _build_return_slide(prs, deal, "Base Case Underwriting Summary")

    if downside_deal is not None:
        _build_return_slide(prs, downside_deal, "Downside Case Underwriting Summary")

    prs.save(output_path)
    return output_path
