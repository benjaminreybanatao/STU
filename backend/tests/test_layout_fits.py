"""Layout guard for the generated deck.

The exhibit tables are laid out at measured absolute coordinates with cell
wrapping switched off, which keeps them faithful to the reference decks but
means over-long text silently spills outside its column instead of reflowing.
That is exactly how the Building Name cell once pushed a full postal address
across the gap into the Sources & Uses table.

These tests assert every shape stays on the slide and every table cell's text
fits the width it was given, for both a full deal and an OM-only deal (where
every financial cell is a placeholder).

Run with: python -m pytest tests/ -q     (from the backend/ directory)
"""
from pathlib import Path

import pytest
from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation

from app.pipeline import build_deal, generate_deck

EMU_PER_PT = 12700
SLIDE_W_PT = 960.0
SLIDE_H_PT = 540.0
CELL_INSET_PT = 10.0  # 5pt left + 5pt right, per _style_cell
FOOTER_TOP_PT = 514.0  # measured footer baseline; tables must stay above it

SAMPLE_DIR = Path(__file__).resolve().parent.parent / "sample_data"

# DejaVu is metrically wider than Gandhi Sans / Calibri, so measuring with it
# makes this check conservative: anything that passes here has headroom in the
# real fonts.
_FONT_PATH = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
_SCALE = 10  # measure at 10x for sub-point resolution
_measure = ImageDraw.Draw(Image.new("RGB", (8, 8)))


def _text_width_pt(text: str, size_pt: float) -> float:
    font = ImageFont.truetype(_FONT_PATH, int(round(size_pt * _SCALE)))
    return _measure.textlength(text, font=font) / _SCALE


def _build(tmp_path, with_model: bool):
    out = tmp_path / "deck.pptx"
    deal = build_deal(
        str(SAMPLE_DIR / "sample_om.pdf"),
        str(SAMPLE_DIR / "sample_model.xlsx") if with_model else None,
        {},
        str(tmp_path),
    )
    generate_deck(deal, str(out))
    return Presentation(str(out))


@pytest.fixture(scope="module")
def decks(tmp_path_factory):
    return {
        "full": _build(tmp_path_factory.mktemp("full"), True),
        "om_only": _build(tmp_path_factory.mktemp("om"), False),
    }


@pytest.mark.parametrize("variant", ["full", "om_only"])
def test_all_shapes_stay_on_slide(decks, variant):
    prs = decks[variant]
    for i, slide in enumerate(prs.slides, start=1):
        for shape in slide.shapes:
            right = (shape.left + shape.width) / EMU_PER_PT
            bottom = (shape.top + shape.height) / EMU_PER_PT
            assert right <= SLIDE_W_PT + 0.5, f"slide {i}: {shape.shape_type} runs off the right edge"
            assert bottom <= SLIDE_H_PT + 0.5, f"slide {i}: {shape.shape_type} runs off the bottom"


@pytest.mark.parametrize("variant", ["full", "om_only"])
def test_table_cell_text_fits_its_column(decks, variant):
    prs = decks[variant]
    overflows = []
    for i, slide in enumerate(prs.slides, start=1):
        for shape in slide.shapes:
            if not shape.has_table:
                continue
            col_widths = [c.width / EMU_PER_PT for c in shape.table.columns]
            for r, row in enumerate(shape.table.rows):
                for c, cell in enumerate(row.cells):
                    if getattr(cell, "is_spanned", False):
                        continue
                    text = cell.text.strip()
                    if not text:
                        continue
                    span = getattr(cell, "span_width", 1)
                    available = sum(col_widths[c:c + span]) - CELL_INSET_PT
                    runs = [run for p in cell.text_frame.paragraphs for run in p.runs if run.text]
                    size_pt = runs[0].font.size.pt if (runs and runs[0].font.size) else 11.2
                    needed = _text_width_pt(text, size_pt)
                    if needed > available:
                        overflows.append(
                            f"slide {i} r{r}c{c}: needs {needed:.0f}pt, has {available:.0f}pt -- {text!r}"
                        )
    assert not overflows, "table text overflows its column:\n  " + "\n  ".join(overflows)


@pytest.mark.parametrize("variant", ["full", "om_only"])
def test_tables_clear_the_footer_even_if_powerpoint_grows_rows(decks, variant):
    """A declared row height is only a minimum.

    PowerPoint grows each row to fit its line and ignores wrap="none" on table
    cells, so a table sized purely on declared heights can still run off the
    slide once the brand font is substituted. Re-check the bottom edge against
    the height PowerPoint would actually use.
    """
    # Deliberately NOT imported from the builder: this is an external fact
    # about PowerPoint. 1.58 was measured from real output (11.2pt text on a
    # ~17.7pt row with the brand font substituted) and still proved too low, so
    # this asserts the stronger margin the builder now designs for. Importing
    # the builder's own constant would make the test merely self-consistent and
    # would stop catching a regression that changed both together.
    POWERPOINT_MIN_ROW_PER_PT = 1.78

    footer_top_pt = FOOTER_TOP_PT
    for i, slide in enumerate(decks[variant].slides, start=1):
        for shape in slide.shapes:
            if not shape.has_table:
                continue
            table = shape.table
            grown = 0.0
            for r, row in enumerate(table.rows):
                declared = row.height / EMU_PER_PT
                sizes = [
                    run.font.size.pt
                    for c in range(len(table.columns))
                    for p in table.cell(r, c).text_frame.paragraphs
                    for run in p.runs
                    if run.font.size
                ]
                needed = max(sizes) * POWERPOINT_MIN_ROW_PER_PT if sizes else 0.0
                grown += max(declared, needed)
            bottom = shape.top / EMU_PER_PT + grown
            assert bottom <= footer_top_pt, (
                f"slide {i}: table would render {bottom:.1f}pt deep, "
                f"past the footer at {footer_top_pt:.1f}pt"
            )


@pytest.mark.parametrize("variant", ["full", "om_only"])
def test_exhibit_tables_do_not_collide(decks, variant):
    """The two exhibit tables sit side by side with only ~9pt between them."""
    prs = decks[variant]
    for i, slide in enumerate(prs.slides, start=1):
        tables = sorted(
            (sh for sh in slide.shapes if sh.has_table), key=lambda sh: sh.left
        )
        for left_tbl, right_tbl in zip(tables, tables[1:]):
            left_right_edge = (left_tbl.left + left_tbl.width) / EMU_PER_PT
            right_left_edge = right_tbl.left / EMU_PER_PT
            assert left_right_edge <= right_left_edge, (
                f"slide {i}: exhibit tables overlap "
                f"({left_right_edge:.1f}pt > {right_left_edge:.1f}pt)"
            )
