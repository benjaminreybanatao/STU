// Base URL of the CRE Deal Screen backend API.
//
// GitHub Pages only serves static files -- it cannot run the FastAPI backend
// that does the actual PDF/Excel/PPTX processing. So when this frontend is
// hosted on GitHub Pages, point this at a separately-deployed backend
// (Render, Fly.io, Railway, a VM, etc. -- see README.md "Deploying the demo
// publicly"). Leave it as "" when the frontend is served BY the backend
// itself (e.g. running `uvicorn` locally per the README) -- that keeps
// requests same-origin and no CORS setup is needed.
window.API_BASE = "";
