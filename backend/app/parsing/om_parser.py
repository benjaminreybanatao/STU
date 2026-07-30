"""
OM (offering memorandum) PDF parser.

Two jobs, kept deliberately separate:
  1. extract_facts() - regex/heuristic pull of narrative deal facts out of
     the raw text layer (address, year built, SF/units, occupancy, etc).
  2. extract_images() - pull embedded raster images out of the PDF and rank
     them so the "hero photo" picked for slide 1 is an actual building
     photograph, not a logo, chart, or floor-plan line-art graphic.

Both are heuristic by nature -- an OM is free-form marketing prose, not a
structured document -- so every extracted fact is a best-effort regex match.
Nothing here is invented; if a pattern doesn't match, the field is simply
left out of the returned dict and the gap analysis will flag it as missing.
"""
import io
import re
import statistics
from pathlib import Path

import fitz  # PyMuPDF
from PIL import Image, ImageChops

ADDRESS_RE = re.compile(
    r"\b(\d{2,6}\s+[A-Z][A-Za-z0-9.'\-]*(?:\s+[A-Z][A-Za-z0-9.'\-]*){0,4}"
    r"\s+(?:Street|St|Avenue|Ave|Boulevard|Blvd|Road|Rd|Drive|Dr|Way|Lane|Ln|"
    r"Parkway|Pkwy|Court|Ct|Place|Pl|Highway|Hwy)\.?,?\s*"
    r"(?:[A-Za-z .]+,\s*)?[A-Z]{2}\s*\d{5}(?:-\d{4})?)\b"
)

YEAR_BUILT_RE = re.compile(
    r"(?:built|constructed|delivered|year built)\D{0,15}((?:19|20)\d{2})",
    re.IGNORECASE,
)

SF_RE = re.compile(
    r"([\d,]{4,10})\s*(?:rentable\s+)?(?:square feet|sf|rsf)\b", re.IGNORECASE
)
UNITS_RE = re.compile(r"([\d,]{1,5})\s*(?:units|unit count|residential units)\b", re.IGNORECASE)

OCCUPANCY_RE = re.compile(r"(\d{1,3}(?:\.\d+)?)\s*%\s*(?:leased|occupied|occupancy)", re.IGNORECASE)

PROPERTY_TYPE_KEYWORDS = {
    "multifamily": ["multifamily", "apartment", "residential units"],
    "office": ["office building", "office property", "office asset"],
    "industrial": ["industrial", "warehouse", "distribution center", "logistics"],
    "retail": ["retail center", "shopping center", "retail property"],
    "mixed-use": ["mixed-use", "mixed use"],
}

YEAR_BUILT_LABEL_RE = re.compile(r"year\s*built\D{0,10}((?:19|20)\d{2})", re.IGNORECASE)


def _find_first(pattern: re.Pattern, text: str) -> str | None:
    m = pattern.search(text)
    return m.group(1).strip() if m else None


def extract_facts(pdf_path: str) -> dict:
    doc = fitz.open(pdf_path)
    full_text = "\n".join(page.get_text() for page in doc)
    doc.close()

    facts: dict = {}

    address = _find_first(ADDRESS_RE, full_text)
    if address:
        facts["address"] = re.sub(r"\s+", " ", address)

    year_built = _find_first(YEAR_BUILT_LABEL_RE, full_text) or _find_first(YEAR_BUILT_RE, full_text)
    if year_built:
        facts["year_built"] = year_built

    sf = _find_first(SF_RE, full_text)
    units = _find_first(UNITS_RE, full_text)
    if sf:
        facts["sf_or_units"] = f"{sf} SF"
    elif units:
        facts["sf_or_units"] = f"{units} Units"

    occupancy = _find_first(OCCUPANCY_RE, full_text)
    if occupancy:
        facts["occupancy"] = f"{occupancy}%"

    lower_text = full_text.lower()
    for ptype, keywords in PROPERTY_TYPE_KEYWORDS.items():
        if any(kw in lower_text for kw in keywords):
            facts["property_type"] = ptype.title()
            break

    # Submarket description: grab a sentence containing "submarket".
    submarket_match = re.search(r"([^.\n]{0,220}\bsubmarket\b[^.\n]{0,220}\.)", full_text, re.IGNORECASE)
    if submarket_match:
        facts["submarket_desc"] = re.sub(r"\s+", " ", submarket_match.group(1)).strip()

    # Tenant summary: look for a sentence mentioning "tenant" with a percentage or name list.
    tenant_match = re.search(r"([^.\n]{0,220}\btenant[s]?\b[^.\n]{0,220}\.)", full_text, re.IGNORECASE)
    if tenant_match:
        facts["tenant_summary"] = re.sub(r"\s+", " ", tenant_match.group(1)).strip()

    return facts


def _is_adobe_inverted_cmyk(raw_bytes: bytes) -> bool:
    """Some Adobe-pipeline JPEGs (Photoshop/InDesign "YCCK" exports) store
    CMYK channel-inverted, which makes naive CMYK->RGB conversion produce
    wrong colors. The JPEG APP14 marker's transform byte -- not just the
    presence of the "Adobe" string, which Pillow's own CMYK encoder also
    writes without inverting -- distinguishes the two: transform == 2
    (YCCK) is the documented inverted case; transform == 0 (plain CMYK,
    e.g. Pillow's own output) is not."""
    idx = raw_bytes.find(b"Adobe")
    if idx == -1 or idx + 11 > len(raw_bytes):
        return False
    transform_byte = raw_bytes[idx + 10]
    return transform_byte == 2


def _normalize_to_rgb(raw_bytes: bytes) -> Image.Image:
    """Normalize every extracted image to plain RGB regardless of source
    colorspace (CMYK, grayscale, palette, etc.), so color rendering is
    consistent across PowerPoint, browsers, and the Pillow preview
    fallback -- rather than passing through whatever colorspace the PDF
    happened to embed."""
    img = Image.open(io.BytesIO(raw_bytes))
    if img.mode == "CMYK":
        if _is_adobe_inverted_cmyk(raw_bytes):
            img = ImageChops.invert(img)
        img = img.convert("RGB")
    elif img.mode != "RGB":
        img = img.convert("RGB")
    return img


def _is_logo_or_chart(width: int, height: int) -> bool:
    """Heuristic filter: reject tiny icons, thin banners, and near-square
    small graphics that are almost always logos/charts/floor plans rather
    than photographs."""
    if width < 300 or height < 200:
        return True
    aspect = width / height
    if aspect > 4.0 or aspect < 0.25:
        return True  # thin banner strip
    return False


def extract_images(pdf_path: str, out_dir: str, max_images: int = 5) -> list[dict]:
    """Extract embedded raster images, filter out logos/charts/floor plans
    by size + aspect ratio, and rank the survivors by pixel area (bigger ==
    more likely to be a full-bleed hero photograph in an OM layout).

    Deterministic: images are always visited in the same (page, xref) order
    a given PDF produces, and ties are broken by that same order, so the
    same OM always yields the same ranked list.
    """
    Path(out_dir).mkdir(parents=True, exist_ok=True)
    doc = fitz.open(pdf_path)
    candidates = []

    for page_index in range(len(doc)):
        page = doc[page_index]
        for img in page.get_images(full=True):
            xref = img[0]
            try:
                base_image = doc.extract_image(xref)
            except Exception:
                continue
            width, height = base_image.get("width", 0), base_image.get("height", 0)
            if _is_logo_or_chart(width, height):
                continue
            try:
                normalized = _normalize_to_rgb(base_image["image"])
            except Exception:
                continue
            filename = f"page{page_index + 1}_xref{xref}.png"
            out_path = str(Path(out_dir) / filename)
            normalized.save(out_path)
            candidates.append(
                {
                    "path": out_path,
                    "width": width,
                    "height": height,
                    "page": page_index + 1,
                    "xref": xref,
                    "area": width * height,
                }
            )

    doc.close()
    candidates.sort(key=lambda c: (-c["area"], c["page"], c["xref"]))
    return candidates[:max_images]
