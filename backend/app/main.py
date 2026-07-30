import json
import shutil
import subprocess
import uuid
from dataclasses import asdict
from pathlib import Path

import fitz  # PyMuPDF -- used to rasterize the LibreOffice-rendered PDF preview
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

from .pipeline import build_deal, generate_deck
from .preview_renderer import render_pptx_to_png
from .schema import Deal

BASE_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = BASE_DIR / "run_data"
DATA_DIR.mkdir(exist_ok=True)
FRONTEND_DIR = BASE_DIR.parent / "frontend"

app = FastAPI(title="CRE Deal Screen")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
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


@app.post("/api/deals")
async def create_deal(
    om_file: UploadFile | None = File(default=None),
    xlsx_file: UploadFile | None = File(default=None),
    analyst_notes_json: str = Form(default="{}"),
):
    if om_file is None and xlsx_file is None:
        raise HTTPException(400, "Upload at least an OM PDF or an underwriting Excel model.")

    deal_id = uuid.uuid4().hex[:12]
    work_dir = _deal_work_dir(deal_id)

    om_path = None
    if om_file is not None:
        if not om_file.filename.lower().endswith(".pdf"):
            raise HTTPException(400, "OM file must be a .pdf")
        om_path = str(work_dir / "om.pdf")
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
    return FileResponse(
        str(path),
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        filename="offering_summary.pptx",
    )


if FRONTEND_DIR.exists():
    app.mount("/", StaticFiles(directory=str(FRONTEND_DIR), html=True), name="frontend")
