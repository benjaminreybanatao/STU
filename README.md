# CRE Deal Screen

A demo tool for commercial real estate deal screening. Upload a raw offering
memorandum (OM -- PDF or PPTX) and the firm's underwriting Excel model for a
property; the tool shows a gap analysis of what's present vs. missing, then
generates a real, editable two-page offering summary as a `.pptx` file.

## Architecture

- **Backend** (`backend/`): FastAPI app that does the actual work -- OM
  text/image extraction (`PyMuPDF` for PDFs, `python-pptx` for PPTX decks --
  brokers send both), Excel cell parsing (`openpyxl`), and native PowerPoint
  generation (`python-pptx`).
- **Frontend** (`frontend/`): a single static page (vanilla HTML/JS) that
  uploads files to the backend and renders the gap-analysis checklist,
  extracted photos, and generated deck preview. Served by the backend at `/`.

This can't be a single zero-dependency static HTML file -- real `.xlsx`
parsing, PDF image extraction, and `.pptx` generation all need server-side
libraries a sandboxed browser can't run.

`pptx_builder.py` is measured from four real DivcoWest ("DW") two-pagers. Two
are PPTX files whose financial exhibit is a pasted Excel screenshot; the other
two are native-vector PDFs where the exhibit is real text, so every position,
colour and row label comes from direct measurement of those rather than
inference. All four agree, so this is the firm template rather than one-off
styling.

What that fixes: theme "DW Colors 2021" (green #6AA442, navy #002554, blue
#0175A8), Gandhi Sans / Gandhi Serif, the two-tone title
("TRANSACTION SUMMARY / {NAME}" with the name in green), the green accent rule
and five-bar logo mark, the footer and page numbers, the stacked two-photo
layout, the bulleted one-topic-per-line narrative, and the real exhibit: two
side-by-side tables ("Transaction Overview" and "Sources & Uses at Close") with
navy header bands, grey section headers, green total rows and a blue subtotal
row. Both are rebuilt as native pptx tables so the deck stays editable, unlike
the screenshot in the sources.

The exhibit's row height and type size are derived from the space between the
table and the footer rather than hard-coded, because a declared row height is
only a *minimum*: PowerPoint grows each row to fit its line and ignores
`wrap="none"` on table cells. Sizing the rows to the available space and
setting the type well under that keeps the table on the slide even when Gandhi
Sans isn't installed and a taller fallback is substituted. The reference's
11.2pt/14.6pt pitch is used whenever there's room for it.

Downloads are named after the property, matching how the firm names these decks
(`Landsby_2pager_v1.pptx`) -- see `pptx_builder.deck_filename`.

Two rules the exhibit follows:

- **`$ Per Unit` is never computed.** Per-unit figures are read only from a
  `$ Per Unit` / `$ PSF` column the model states explicitly (see
  `excel_parser._find_per_unit_columns`). Dividing a total by a unit count
  would be inventing a number, so a model without that column leaves the
  column showing the placeholder.
- **Missing line items placeholder rather than disappear.** Sources & Uses rows
  the model doesn't carry render as a visible `TBD — confirm with sponsor`.

Not generated: the Location Overview map, Yield Bridge chart and Ground Lease
Detail slides in the reference decks. Each needs hand-assembled inputs (maps,
comp sets, tenant logos, lease abstracts) this tool has no source for.

## Data model

Every deal tracks three provenance buckets (`backend/app/schema.py`):

- `om_facts` -- parsed out of the OM's text layer (regex/heuristic; PDF via
  PyMuPDF, PPTX via python-pptx).
- `model_facts` -- parsed out of the Excel model via label-adjacent cell
  scanning (search for a label like "Levered IRR", read the value next to
  it). This is layout-tolerant by design since every firm's model is laid
  out differently; pass an explicit `cell_map` to `excel_parser.extract_facts`
  to pin exact cells for a known firm template instead. If the workbook has
  a tab named "Memo Charts", the exhibit's numbers are read from that sheet
  exclusively rather than scanning the whole workbook -- it's meant to hold
  the deck's own figures, so it's a more reliable source than an
  underwriting tab where the same label can appear more than once or sit
  next to an unrelated number.
- `analyst_inputs` -- free text supplied by the user in the UI.

`FIELD_REGISTRY` is the single source of truth for what the two-pager
needs, and each field declares a source priority (e.g. financial metrics
prefer the model, narrative facts prefer the OM). The gap analysis and the
deck builder both read through this same registry, so a number can never
appear on the deck without a traceable source. Anything not found in any
bucket renders as `TBD — confirm with sponsor`, styled in red on the deck,
rather than being invented.

## Image extraction

`om_parser.extract_images` pulls every embedded raster image out of the PDF,
filters out anything under 300x200px or with an extreme aspect ratio (thin
banners/logos), and ranks survivors by pixel area so the largest
photograph-shaped image wins as the slide 1 hero photo. The UI lets you
override the pick from the ranked list.

## Preview rendering

`/generate` tries LibreOffice headless (`soffice --convert-to pdf`, then
rasterized with PyMuPDF) first for a pixel-accurate preview. If that's
unavailable or broken in the deployment environment, it falls back to
`preview_renderer.py`, which walks the saved `.pptx` directly with
`python-pptx` and redraws shapes/tables/pictures onto a Pillow canvas. Either
way the actual downloadable file is the same native `.pptx` -- the fallback
only affects the in-browser preview image, never the deck itself.

## Access control

Every deal handled here can contain confidential OM/underwriting data, so
all `/api/*` routes (uploads, extracted images, previews, downloads) require
a shared access code, set via the `ACCESS_CODE` environment variable. The
frontend shows a gate on load and won't let you into the app without it.
`fetch()` calls send it as an `X-Access-Code` header; plain `<img>`/`<a>`
GETs (images, previews, downloads) can't attach custom headers, so those
carry it as a `?code=` query param instead -- the backend accepts either.

If `ACCESS_CODE` isn't set, the backend falls back to the well-known default
`changeme` and logs a loud warning on startup -- that fallback exists so a
missing env var fails safe-ish (it still requires *a* code) rather than
silently running wide open, but it is **not** real protection. Always set a
real `ACCESS_CODE` before deploying anywhere reachable from the internet.

The static frontend itself (`index.html`/`app.js`/`style.css`) is not gated
-- there's nothing sensitive in it, and the gate has to be loadable before
you've entered a code. Only the API routes that touch deal data are.

## Running it

```bash
cd backend
pip install -r requirements.txt
ACCESS_CODE=your-shared-code python -m uvicorn app.main:app --host 0.0.0.0 --port 8000
```

Then open `http://localhost:8000/` and enter that code.

Test fixtures (`backend/sample_data/`): a synthetic OM PDF with a base/
downside underwriting model, generated by `make_samples.py`. Regenerate with:

```bash
cd backend/sample_data && python3 make_samples.py
```

## Tests

```bash
cd backend
pip install -r requirements-dev.txt
python -m pytest tests/ -q
```

`tests/test_layout_fits.py` is a layout guard. The exhibit tables are placed at
measured absolute coordinates with cell wrapping switched off — faithful to the
reference decks, but it means over-long text spills outside its column instead
of reflowing. The tests assert every shape stays on the slide, every table
cell's text fits the width it was given, and the two side-by-side exhibit
tables never overlap — checked for both a full deal and an OM-only deal (where
every financial cell is a placeholder). Text is measured in DejaVu Sans, which
is wider than Gandhi Sans, so anything that passes has headroom in the real
font.

## Deploying the demo publicly (GitHub Pages + a hosted backend)

GitHub Pages only serves static files -- it cannot run the FastAPI backend.
To get a working demo at a `github.io` URL, the frontend and backend are
deployed separately and wired together with one config value:

1. **Deploy the backend somewhere that runs Python**, e.g. [Render](https://render.com)
   using the included `render.yaml` blueprint (New -> Blueprint -> point at
   this repo). No LibreOffice needed -- `preview_renderer.py`'s Pillow-based
   fallback works fine on a bare Python host. Note the resulting URL, e.g.
   `https://cre-deal-screen-api.onrender.com`. **Set `ACCESS_CODE`** (and
   ideally `ALLOWED_ORIGIN` to your Pages URL) in the service's environment
   variables before pointing any real deal data at it -- see "Access
   control" above.
2. **Point the frontend at it**: edit `frontend/config.js` and set
   `window.API_BASE = "https://cre-deal-screen-api.onrender.com";`, then
   commit and push to `main`.
3. **Turn on GitHub Actions-based Pages** (one-time, manual -- this isn't
   settable via the API): repo Settings -> Pages -> Build and deployment ->
   Source -> **GitHub Actions**. `.github/workflows/pages.yml` is already in
   the repo and publishes `frontend/` on every push to `main` once that's
   set.

After that, `https://<user>.github.io/<repo>/` serves the real interface,
calling the Render-hosted API. Without step 3 flipped, Pages falls back to
its default branch/Jekyll build, which just renders this README instead of
the app -- that's the "why doesn't the interface show up" most people hit.

## API

- `POST /api/deals` -- multipart upload (`om_file`, `xlsx_file`, both
  optional but at least one required; `analyst_notes_json`). Returns
  `deal_id`, `gap_analysis`, and ranked `images`.
- `PATCH /api/deals/{id}/analyst-inputs` -- merge analyst-entered values,
  returns updated `gap_analysis`.
- `POST /api/deals/{id}/select-hero-image` -- reorder extracted images to
  choose the slide 1 hero photo.
- `POST /api/deals/{id}/generate` -- builds the `.pptx` and a PNG preview
  per slide.
- `GET /api/deals/{id}/download` -- the generated `.pptx` file.
