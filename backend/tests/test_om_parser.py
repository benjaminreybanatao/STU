"""Image-extraction regression coverage for the OM parser.

Run with: python -m pytest tests/ -q     (from the backend/ directory)
"""
import io

import pytest
from PIL import Image
from pptx import Presentation
from pptx.util import Inches

from app.parsing.om_parser import _extract_images_from_pptx, _normalize_to_rgb


def test_cmyk_jpeg_is_not_double_inverted():
    """Pillow's JPEG decoder already un-inverts CMYK sample data for every
    CMYK-mode JPEG as it decodes (it assumes Adobe conventions unconditionally,
    regardless of the APP14 transform byte). A previous version of
    _normalize_to_rgb re-inverted the image again whenever it saw a "YCCK"
    (transform == 2) marker, undoing Pillow's already-correct decode and
    rendering the photo as a color negative -- exactly what shipped to a real
    deck. This locks in the fix by patching a real Adobe marker's transform
    byte to 2 and confirming the round-tripped color still matches."""
    original = Image.new("RGB", (32, 32), (40, 120, 200))
    buf = io.BytesIO()
    original.convert("CMYK").save(buf, format="JPEG", quality=95)
    raw = bytearray(buf.getvalue())

    idx = raw.find(b"Adobe")
    assert idx != -1, "test JPEG should carry an Adobe APP14 marker"
    raw[idx + 10] = 2  # force the "YCCK" transform byte the old code targeted

    normalized = _normalize_to_rgb(bytes(raw))
    assert normalized.mode == "RGB"
    r, g, b = normalized.getpixel((16, 16))
    # JPEG is lossy, so allow some slack -- an inverted result would be off
    # by ~255 per channel, nowhere close to this tolerance.
    assert abs(r - 40) < 20 and abs(g - 120) < 20 and abs(b - 200) < 20


def test_unloadable_picture_is_skipped_not_fatal(tmp_path):
    """A picture that Image.open() accepts (valid header) but can't actually
    be decoded (truncated data) used to crash extraction for the whole deck,
    since the failure only surfaces on save()/load(), after the original
    try/except around open() had already passed. One bad picture should be
    skipped, not take down every other picture in the deck."""
    truncated_path = tmp_path / "truncated.png"
    good_path = tmp_path / "good.png"

    full = io.BytesIO()
    Image.new("RGB", (400, 300), (10, 20, 30)).save(full, format="PNG")
    truncated_path.write_bytes(full.getvalue()[: len(full.getvalue()) // 2])
    Image.new("RGB", (400, 300), (50, 60, 70)).save(good_path)

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.shapes.add_picture(str(truncated_path), Inches(0), Inches(0), width=Inches(2))
    slide.shapes.add_picture(str(good_path), Inches(3), Inches(0), width=Inches(2))
    pptx_path = tmp_path / "deck.pptx"
    prs.save(str(pptx_path))

    out_dir = tmp_path / "extracted"
    out_dir.mkdir()
    results = _extract_images_from_pptx(str(pptx_path), str(out_dir))

    assert len(results) == 1
    assert Image.open(results[0]["path"]).getpixel((0, 0)) == (50, 60, 70)
