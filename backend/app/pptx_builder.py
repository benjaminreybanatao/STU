"""
Builds the two-page (or three-page, if a downside case is present) offering
summary deck.

Reverse-engineered from two real DivcoWest ("DW") two-pager decks (400 Castro,
2390 Mission College), inspected directly via python-pptx + raw XML/theme
parsing:

  - Theme: "DW Colors 2021" (ppt/theme/theme1.xml) -- accent1 green #6AA23A,
    accent3 navy #002554, accent6 gold #E9A800, dk2 gray #626369, white bg.
  - Fonts: "Gandhi Sans" for titles/labels/headers, "Gandhi Serif" for body
    narrative copy.
  - Slide 1 ("Transaction Summary"): plain white background, no header band.
    Title top-left reads "TRANSACTION SUMMARY / {SHORT ADDRESS}", 28pt bold
    Gandhi Sans, black with the property name in green (#6AA442, matches
    exactly across both source decks). Two stacked property photos on the
    left; a narrative textbox on the right written as one short paragraph
    per topic (not one merged block) in 12-12.5pt Gandhi Serif, with defined
    terms ("Investment"/"Property") highlighted in dark green (#3A592A).
  - Slide 2 ("Base Case"): same title style. In both source decks this is a
    single pasted image of a financial summary exhibit (a paste from Excel --
    exactly the "PDF/PPT export with embedded images" case flagged in the
    spec). Decoding the embedded EMF's text records recovered its structure:
    Property Overview / Pricing / Debt / Equity / Gross Returns panels, each
    a small label/value table. That structure is rebuilt here as native,
    editable pptx tables -- not a pasted image -- using only fields this
    tool can actually trace to the OM/model/analyst input. (Line items from
    the source decks that require a granular Sources & Uses breakdown this
    tool doesn't parse -- Acquisition Cost, DD/Closing Costs, Working
    Capital, Financing Cost -- are intentionally omitted rather than
    fabricated.)

Every value placed on the slide comes from Deal.display_value(), which
walks the OM -> model -> analyst provenance chain and falls back to the
PLACEHOLDER string. Nothing is fabricated here.
"""
import re

from lxml import etree

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

from .schema import Deal, FIELD_BY_KEY, PLACEHOLDER

# DW Colors 2021 theme, read directly out of ppt/theme/theme1.xml.
BLACK = RGBColor(0x00, 0x00, 0x00)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
GRAY_DARK = RGBColor(0x62, 0x63, 0x69)   # dk2
GRAY_LIGHT = RGBColor(0x95, 0x96, 0x9A)  # lt2
NAVY = RGBColor(0x00, 0x25, 0x54)        # accent3
GOLD = RGBColor(0xE9, 0xA8, 0x00)        # accent6
GREEN = RGBColor(0x6A, 0xA4, 0x42)       # title property-name highlight (exact value in both source decks)
DARK_GREEN = RGBColor(0x3A, 0x59, 0x2A)  # defined-term highlight in body copy
ROW_ALT = RGBColor(0xF3, 0xF4, 0xF5)
BORDER_GRAY = RGBColor(0xD9, 0xDA, 0xDC)
PLACEHOLDER_RED = RGBColor(0xB0, 0x2A, 0x2A)

SANS = "Gandhi Sans"
SERIF = "Gandhi Serif"

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)

TITLE_LEFT = Inches(0.545)
TITLE_TOP = Inches(0.813)
TITLE_HEIGHT = Inches(0.542)

PANELS = [
    ("Property Overview", ["address", "property_type", "sf_or_units", "occupancy", "hold_period"]),
    ("Pricing", ["purchase_price", "price_per_unit", "going_in_cap", "exit_cap"]),
    ("Debt", ["leverage", "debt_rate"]),
    ("Equity", ["initial_equity", "peak_equity"]),
    ("Gross Returns", ["unlevered_irr", "unlevered_em", "levered_irr", "levered_em", "cash_on_cash"]),
]


def _strip_street_suffix(street: str) -> str:
    """'400 Castro Street' -> '400 Castro' -- both source decks drop the
    street-type suffix when using the address as the property's short name."""
    stripped = re.sub(r"\b(Street|St|Avenue|Ave|Boulevard|Blvd|Road|Rd|Drive|Dr)\b\.?", "", street, flags=re.IGNORECASE)
    return re.sub(r"\s+", " ", stripped).strip()


def _short_title_label(address: str) -> str:
    """'1200 Market Street, Austin, TX 78701' -> '1200 MARKET' -- title uses
    street number + name only, dropping the suffix and city/state/zip
    (matches source decks' '400 CASTRO', '2390 MISSION COLLEGE')."""
    if address == PLACEHOLDER:
        return "PROPERTY TBD"
    return _strip_street_suffix(address.split(",")[0].strip()).upper()


def _add_title(slide, address: str):
    box = slide.shapes.add_textbox(TITLE_LEFT, TITLE_TOP, Inches(10.5), TITLE_HEIGHT)
    tf = box.text_frame
    tf.margin_left = 0
    tf.margin_top = 0
    tf.margin_right = 0
    tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]

    r1 = p.add_run()
    r1.text = "TRANSACTION SUMMARY / "
    r1.font.size = Pt(28)
    r1.font.bold = True
    r1.font.name = SANS
    r1.font.color.rgb = BLACK

    r2 = p.add_run()
    r2.text = _short_title_label(address)
    r2.font.size = Pt(28)
    r2.font.bold = True
    r2.font.name = SANS
    r2.font.color.rgb = GREEN
    return box


def _split_address(address: str) -> tuple[str, str]:
    """'1200 Market Street, Austin, TX 78701' -> ('1200 Market Street', 'Austin, TX')."""
    if address == PLACEHOLDER:
        return address, address
    parts = [p.strip() for p in address.split(",")]
    street = _strip_street_suffix(parts[0]) if parts else address
    city_state = ", ".join(parts[1:-1]) if len(parts) > 2 else (parts[1] if len(parts) > 1 else "")
    return street, city_state


def _build_narrative_paragraphs(deal: Deal) -> list[str]:
    """One short paragraph per topic (property, ownership/description,
    tenancy, pricing, business plan, capital structure, exit & returns),
    matching the source decks' pattern of separate sentence-groups rather
    than a single merged block."""
    dv = deal.display_value
    paras = []

    ptype = dv("property_type")
    ptype_phrase = ptype.lower() if ptype != PLACEHOLDER else ptype
    article = "an" if ptype_phrase[:1].lower() in "aeiou" else "a"
    street, city_state = _split_address(dv("address"))
    location = f" located in {city_state}" if city_state else ""
    paras.append(
        f'Opportunity to acquire {street}, {article} {ptype_phrase} building totaling '
        f'{dv("sf_or_units")}{location} (the “Investment” or the “Property”).'
    )

    year_built = dv("year_built")
    if year_built != PLACEHOLDER:
        paras.append(f"The Property was built in {year_built}.")
    submarket = dv("submarket_desc")
    if submarket != PLACEHOLDER:
        paras.append(submarket)

    tenant_summary = dv("tenant_summary")
    occupancy = dv("occupancy")
    walt = dv("walt")
    tenancy_sentence = f"The Property is {occupancy} occupied"
    if tenant_summary != PLACEHOLDER:
        tenancy_sentence += f"; {tenant_summary}"
    if walt != PLACEHOLDER:
        tenancy_sentence += f", resulting in a {walt} WALT"
    paras.append(tenancy_sentence + ".")

    paras.append(
        f'The projected purchase price for the building is {dv("purchase_price")} '
        f'({dv("price_per_unit")}), resulting in a {dv("going_in_cap")} going-in cap rate.'
    )

    lease_term = dv("lease_term_assumption")
    downtime = dv("downtime_assumption")
    exit_assumption = dv("exit_assumption")
    if any(v != PLACEHOLDER for v in (lease_term, downtime, exit_assumption)):
        bits = []
        if lease_term != PLACEHOLDER:
            bits.append(f"a {lease_term} lease term")
        if downtime != PLACEHOLDER:
            bits.append(f"{downtime} downtime")
        if exit_assumption != PLACEHOLDER:
            bits.append(exit_assumption.rstrip("."))
        paras.append("Base case business plan assumes " + "; ".join(bits) + ".")

    peak_equity = dv("peak_equity")
    debt_sentence = f'With {dv("leverage")} leverage at {dv("debt_rate")}, the initial equity is {dv("initial_equity")}'
    if peak_equity != PLACEHOLDER:
        debt_sentence += f" and the projected peak equity for the transaction is {peak_equity}"
    paras.append(debt_sentence + ".")

    paras.append(
        f'Assuming the Property is sold in {dv("hold_period")} at a {dv("exit_cap")} exit cap rate, '
        f'projected unlevered returns are {dv("unlevered_irr")} / {dv("unlevered_em")} and levered '
        f'returns are {dv("levered_irr")} / {dv("levered_em")}.'
    )

    return paras


BULLET_INDENT = Pt(16)
BULLET_HANG = Pt(14)


def _set_bullet(paragraph, char="•", color=GREEN):
    """python-pptx has no high-level bullet API -- set marL/indent plus a
    buFont/buChar pair directly on the paragraph's pPr, in the schema order
    CT_TextParagraphProperties requires (after spcAft, before defRPr)."""
    pPr = paragraph._p.get_or_add_pPr()
    pPr.set("marL", str(int(BULLET_INDENT)))
    pPr.set("indent", str(-int(BULLET_HANG)))

    buClr = etree.SubElement(pPr, qn("a:buClr"))
    solidFill = etree.SubElement(buClr, qn("a:srgbClr"))
    solidFill.set("val", str(color))

    buFont = etree.SubElement(pPr, qn("a:buFont"))
    buFont.set("typeface", "Arial")

    buChar = etree.SubElement(pPr, qn("a:buChar"))
    buChar.set("char", char)


def _render_narrative(slide, deal: Deal, left, top, width, height):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = 0
    tf.margin_right = 0

    paragraphs = _build_narrative_paragraphs(deal)
    first = True
    for para_text in paragraphs:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.space_after = Pt(10)
        _set_bullet(p)

        # Highlight the defined terms in the opening sentence, and any
        # placeholder text everywhere else, distinctly -- everything else is
        # plain body copy.
        segments = re.split(r'(“Investment”|“Property”|' + re.escape(PLACEHOLDER) + r')', para_text)
        for seg in segments:
            if not seg:
                continue
            run = p.add_run()
            run.text = seg
            run.font.size = Pt(12.5)
            run.font.name = SERIF
            if seg == PLACEHOLDER:
                run.font.bold = True
                run.font.color.rgb = PLACEHOLDER_RED
            elif seg in ('“Investment”', '“Property”'):
                run.font.color.rgb = DARK_GREEN
            else:
                run.font.color.rgb = BLACK
    return box


def _add_photos(slide, image_paths, left, top, width, height, gap=Inches(0.08)):
    slots = 2
    slot_h = Emu(int((int(height) - int(gap)) / slots))
    paths = (image_paths or [None, None])[:slots]
    while len(paths) < slots:
        paths.append(paths[-1] if paths else None)

    for i, path in enumerate(paths):
        slot_top = Emu(int(top) + i * (int(slot_h) + int(gap)))
        if not path:
            placeholder = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, slot_top, width, slot_h)
            placeholder.fill.solid()
            placeholder.fill.fore_color.rgb = ROW_ALT
            placeholder.line.color.rgb = BORDER_GRAY
            tf = placeholder.text_frame
            tf.text = "No property photo extracted from OM"
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            p.runs[0].font.italic = True
            p.runs[0].font.size = Pt(10)
            p.runs[0].font.color.rgb = GRAY_DARK
            continue

        pic = slide.shapes.add_picture(path, left, slot_top, width=width, height=slot_h)
        native_w, native_h = pic.image.size
        target_ratio = width / slot_h
        native_ratio = native_w / native_h
        if native_ratio > target_ratio:
            crop = (1 - target_ratio / native_ratio) / 2
            pic.crop_left = crop
            pic.crop_right = crop
        else:
            crop = (1 - native_ratio / target_ratio) / 2
            pic.crop_top = crop
            pic.crop_bottom = crop


def _build_slide1(prs, deal: Deal, hero_image_paths):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _add_title(slide, deal.display_value("address"))

    _add_photos(slide, hero_image_paths, Inches(0.562), Inches(1.541), Inches(4.42), Inches(5.03))
    _render_narrative(slide, deal, Inches(5.216), Inches(1.541), Inches(7.37), Inches(5.39))
    return slide


def _panel_table(slide, deal: Deal, title, keys, left, top, width):
    row_h = Inches(0.28)
    header_h = Inches(0.32)
    rows = len(keys) + 1

    table_shape = slide.shapes.add_table(rows, 2, left, top, width, header_h + row_h * len(keys))
    table = table_shape.table
    table.columns[0].width = Emu(int(width * 0.6))
    table.columns[1].width = Emu(int(width * 0.4))
    table.rows[0].height = header_h
    for r in range(1, rows):
        table.rows[r].height = row_h

    header_cell = table.cell(0, 0)
    header_cell.merge(table.cell(0, 1))
    header_cell.fill.solid()
    header_cell.fill.fore_color.rgb = NAVY
    header_cell.text = title.upper()
    hp = header_cell.text_frame.paragraphs[0]
    hp.alignment = PP_ALIGN.LEFT
    hr = hp.runs[0]
    hr.font.size = Pt(10.5)
    hr.font.bold = True
    hr.font.name = SANS
    hr.font.color.rgb = WHITE

    for i, key in enumerate(keys, start=1):
        label = FIELD_BY_KEY[key].label
        value = deal.display_value(key)
        row_fill = WHITE if i % 2 else ROW_ALT
        for col, (text, align, bold) in enumerate([(label, PP_ALIGN.LEFT, False), (value, PP_ALIGN.RIGHT, True)]):
            cell = table.cell(i, col)
            cell.fill.solid()
            cell.fill.fore_color.rgb = row_fill
            cell.margin_left = Pt(6)
            cell.margin_right = Pt(6)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.text = str(text)
            p = cell.text_frame.paragraphs[0]
            p.alignment = align
            run = p.runs[0]
            run.font.size = Pt(10)
            run.font.bold = bold
            run.font.name = SANS
            run.font.color.rgb = PLACEHOLDER_RED if text == PLACEHOLDER else GRAY_DARK

    return table_shape


def _build_return_slide(prs, deal: Deal, address: str, case_label: str | None = None):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _add_title(slide, address)

    if case_label:
        tag = slide.shapes.add_textbox(TITLE_LEFT, Inches(1.28), Inches(3.0), Inches(0.32))
        tf = tag.text_frame
        tf.margin_left = 0
        p = tf.paragraphs[0]
        run = p.add_run()
        run.text = case_label.upper()
        run.font.size = Pt(12)
        run.font.bold = True
        run.font.name = SANS
        run.font.color.rgb = GOLD

    panel_top = Inches(1.75)
    panel_width = Inches(2.35)
    gap = Inches(0.12)
    left = TITLE_LEFT
    for title, keys in PANELS:
        _panel_table(slide, deal, title, keys, left, panel_top, panel_width)
        left = Emu(int(left) + int(panel_width) + int(gap))

    footer = slide.shapes.add_textbox(TITLE_LEFT, Inches(7.05), Inches(12.4), Inches(0.35))
    tf = footer.text_frame
    tf.text = "Every figure above traces to the OM, the underwriting model, or an explicit analyst input -- nothing is estimated."
    p = tf.paragraphs[0]
    p.runs[0].font.size = Pt(9)
    p.runs[0].font.italic = True
    p.runs[0].font.name = SERIF
    p.runs[0].font.color.rgb = GRAY_LIGHT
    return slide


def build_deck(deal: Deal, hero_image_path, output_path: str, downside_deal: Deal | None = None):
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H

    hero_image_paths = hero_image_path if isinstance(hero_image_path, list) else [hero_image_path]
    address = deal.display_value("address")

    _build_slide1(prs, deal, hero_image_paths)
    _build_return_slide(prs, deal, address, case_label="Base Case" if downside_deal is not None else None)

    if downside_deal is not None:
        _build_return_slide(prs, downside_deal, address, case_label="Downside Case")

    prs.save(output_path)
    return output_path
