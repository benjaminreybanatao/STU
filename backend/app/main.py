import json
import logging
import os
import shutil
import subprocess
import uuid
from dataclasses import asdict
from pathlib import Path

import fitz  # PyMuPDF -- used to rasterize the LibreOffice-rendered PDF preview
from fastapi import Depends, FastAPI, Header, UploadFile, File, Form, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

from .pipeline import build_deal, generate_deck
from .pptx_builder import deck_filename
from .preview_renderer import render_pptx_to_png
from .schema import Deal

logger = logging.getLogger("uvicorn.error")

BASE_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = BASE_DIR / "run_data"
DATA_DIR.mkdir(exist_ok=True)
FRONTEND_DIR = BASE_DIR.parent / "frontend"

# Every deal handled here can contain confidential underwriting data, so all
# /api/* routes require a shared access code (ACCESS_CODE env var). Falls
# back to a well-known default rather than silently running wide open if
# someone forgets to set it -- but that fallback is loudly logged, since it
# provides no real protection.
ACCESS_CODE = os.environ.get("ACCESS_CODE")
if not ACCESS_CODE:
    ACCESS_CODE = "changeme"
    logger.warning(
        "ACCESS_CODE is not set -- falling back to the default 'changeme', "
        "which provides no real protection. Set ACCESS_CODE in your deployment "
        "environment before pointing real deal data at this service."
    )


def require_access_code(x_access_code: str | None = Header(default=None), code: str | None = None):
    # Accept the code via header (used by fetch() JSON calls) or a query
    # param (used by plain <img>/<a> GETs for images/previews/downloads,
    # which can't attach custom headers).
    if ACCESS_CODE not in (x_access_code, code):
        raise HTTPException(401, "Missing or invalid access code.")


# Restrict cross-origin calls to a known frontend origin when deployed
# separately (e.g. GitHub Pages calling a Render-hosted backend); defaults to
# "*" for local development where frontend and backend share an origin.
ALLOWED_ORIGIN = os.environ.get("ALLOWED_ORIGIN", "*")

app = FastAPI(title="CRE Deal Screen", dependencies=[Depends(require_access_code)])

app.add_middleware(
    CORSMiddleware,
    allow_origins=[ALLOWED_ORIGIN],
    allow_methods=["*"],
    allow_headers=["*"],
)

# In-memory registry of deals for this demo process (fine for a single-session tool).
DEALS: dict[str, Deal] = {}


def _deal_work_dir(deal_id: str) -> Path:
    d = DATA_DIR / deal_id
    d.mkdir(parents=True, exist_ok=True)
    return d


def _gap_analysis_payload(deal: Deal) -> list[dict]:
    return [asdict(rf) for rf in deal.gap_analysis()]


def _images_payload(deal: Deal, deal_id: str) -> list[dict]:
    out = []
    for img in deal.om_images:
        out.append(
            {
                "url": f"/api/deals/{deal_id}/image/{Path(img['path']).name}",
                "width": img["width"],
                "height": img["height"],
                "page": img["page"],
            }
        )
    return out


@app.get("/api/auth/check")
async def auth_check():
    """No-op endpoint the frontend calls to validate a stored access code
    without side effects -- the require_access_code dependency above already
    rejects the request with 401 before this body ever runs if it's wrong."""
    return {"ok": True}


@app.post("/api/deals")
async def create_deal(
    om_file: UploadFile | None = File(default=None),
    xlsx_file: UploadFile | None = File(default=None),
    analyst_notes_json: str = Form(default="{}"),
):
    if om_file is None and xlsx_file is None:
        raise HTTPException(400, "Upload at least an OM (PDF or PPTX) or an underwriting Excel model.")

    deal_id = uuid.uuid4().hex[:12]
    work_dir = _deal_work_dir(deal_id)

    om_path = None
    if om_file is not None:
        om_name = om_file.filename.lower()
        if not om_name.endswith((".pdf", ".pptx", ".pptm")):
            raise HTTPException(400, "OM file must be a .pdf or .pptx")
        om_ext = ".pdf" if om_name.endswith(".pdf") else Path(om_name).suffix
        om_path = str(work_dir / f"om{om_ext}")
        with open(om_path, "wb") as f:
            shutil.copyfileobj(om_file.file, f)

    xlsx_path = None
    if xlsx_file is not None:
        if not xlsx_file.filename.lower().endswith((".xlsx", ".xlsm")):
            raise HTTPException(400, "Underwriting model must be a .xlsx or .xlsm")
        xlsx_path = str(work_dir / "model.xlsx")
        with open(xlsx_path, "wb") as f:
            shutil.copyfileobj(xlsx_file.file, f)

    try:
        analyst_inputs = json.loads(analyst_notes_json) if analyst_notes_json else {}
    except json.JSONDecodeError:
        analyst_inputs = {}

    deal = build_deal(om_path, xlsx_path, analyst_inputs, str(work_dir))
    DEALS[deal_id] = deal

    return {
        "deal_id": deal_id,
        "gap_analysis": _gap_analysis_payload(deal),
        "images": _images_payload(deal, deal_id),
    }


@app.patch("/api/deals/{deal_id}/analyst-inputs")
async def update_analyst_inputs(deal_id: str, payload: dict):
    deal = DEALS.get(deal_id)
    if deal is None:
        raise HTTPException(404, "Unknown deal_id")
    deal.analyst_inputs.update(payload)
    return {"gap_analysis": _gap_analysis_payload(deal)}


@app.get("/api/deals/{deal_id}/image/{filename}")
async def get_image(deal_id: str, filename: str):
    path = _deal_work_dir(deal_id) / "om_images" / filename
    if not path.exists():
        raise HTTPException(404, "Image not found")
    return FileResponse(str(path))


@app.post("/api/deals/{deal_id}/select-hero-image")
async def select_hero_image(deal_id: str, payload: dict):
    deal = DEALS.get(deal_id)
    if deal is None:
        raise HTTPException(404, "Unknown deal_id")
    filename = payload.get("filename")
    for img in deal.om_images:
        if Path(img["path"]).name == filename:
            deal.om_images.remove(img)
            deal.om_images.insert(0, img)
            return {"ok": True}
    raise HTTPException(404, "Image not found among extracted images")


@app.post("/api/deals/{deal_id}/generate")
async def generate(deal_id: str):
    deal = DEALS.get(deal_id)
    if deal is None:
        raise HTTPException(404, "Unknown deal_id")

    work_dir = _deal_work_dir(deal_id)
    pptx_path = work_dir / "offering_summary.pptx"
    generate_deck(deal, str(pptx_path))

    preview_dir = work_dir / "preview"
    preview_dir.mkdir(exist_ok=True)
    for old in preview_dir.glob("*.png"):
        old.unlink()

    render_mode = "libreoffice"
    try:
        subprocess.run(
            [
                "soffice", "--headless", "--norestore",
                "--convert-to", "pdf", "--outdir", str(preview_dir), str(pptx_path),
            ],
            check=True, capture_output=True, timeout=120,
        )
        pdf_path = preview_dir / "offering_summary.pdf"
        if not pdf_path.exists():
            raise RuntimeError("soffice reported success but produced no PDF")
        pdf_doc = fitz.open(str(pdf_path))
        for i, page in enumerate(pdf_doc, start=1):
            pix = page.get_pixmap(matrix=fitz.Matrix(1.6, 1.6))
            pix.save(str(preview_dir / f"slide-{i:02d}.png"))
        pdf_doc.close()
    except Exception:
        # LibreOffice unavailable/broken in this environment -- fall back to
        # a native python-pptx + Pillow re-render so preview still works.
        render_mode = "native-fallback"
        try:
            render_pptx_to_png(str(pptx_path), str(preview_dir))
        except Exception as exc:
            return JSONResponse(
                {"pptx_ready": True, "preview_error": str(exc), "preview_images": []},
            )

    images = sorted(preview_dir.glob("slide-*.png"))
    preview_urls = [f"/api/deals/{deal_id}/preview/{p.name}" for p in images]
    return {"pptx_ready": True, "preview_images": preview_urls, "render_mode": render_mode}


@app.get("/api/deals/{deal_id}/preview/{filename}")
async def get_preview(deal_id: str, filename: str):
    path = _deal_work_dir(deal_id) / "preview" / filename
    if not path.exists():
        raise HTTPException(404, "Preview image not found")
    return FileResponse(str(path))


@app.get("/api/deals/{deal_id}/download")
async def download(deal_id: str):
    path = _deal_work_dir(deal_id) / "offering_summary.pptx"
    if not path.exists():
        raise HTTPException(404, "Deck not generated yet -- call /generate first.")

    # Name the download after the property, the way the firm names these decks
    # (e.g. "Landsby_2pager_draft.pptx"). Falls back to a generic name if the
    # deal is no longer in memory (the registry is per-process).
    deal = DEALS.get(deal_id)
    filename = deck_filename(deal.display_value("address")) if deal else "Property_2pager_draft.pptx"

    return FileResponse(
        str(path),
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        filename=filename,
    )


if FRONTEND_DIR.exists():
    app.mount("/", StaticFiles(directory=str(FRONTEND_DIR), html=True), name="frontend")
