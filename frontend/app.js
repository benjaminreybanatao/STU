const state = {
  dealId: null,
  gapAnalysis: [],
  images: [],
};

const ACCESS_CODE_KEY = "stu_access_code";

// window.API_BASE is set in config.js -- "" for same-origin (backend serves
// this frontend directly), or an absolute URL when this page is hosted
// separately (e.g. GitHub Pages) from the backend.
function apiUrl(path) {
  return (window.API_BASE || "") + path;
}

// Every deal can contain confidential documents, so the backend gates all
// /api/* routes behind a shared access code. fetch() calls send it as a
// header; plain <img>/<a> GETs (images, previews, downloads) can't attach
// custom headers, so those append it as a query param instead -- the
// backend's require_access_code dependency accepts either.
function getAccessCode() {
  return sessionStorage.getItem(ACCESS_CODE_KEY) || "";
}

function mediaUrl(path) {
  const code = getAccessCode();
  const sep = path.includes("?") ? "&" : "?";
  return apiUrl(path) + (code ? `${sep}code=${encodeURIComponent(code)}` : "");
}

async function apiFetch(path, options = {}) {
  const headers = { ...(options.headers || {}), "X-Access-Code": getAccessCode() };
  const res = await fetch(apiUrl(path), { ...options, headers });
  if (res.status === 401) {
    sessionStorage.removeItem(ACCESS_CODE_KEY);
    showAccessGate("Session expired -- please re-enter the access code.");
  }
  return res;
}

const GROUP_LABELS = {
  property: "Property",
  tenancy: "Tenancy",
  pricing: "Pricing",
  capital_stack: "Capital Stack",
  returns: "Returns",
  sources_uses: "Sources & Uses",
  per_unit: "Per-Unit Figures",
  assumptions: "Assumptions",
};

// ---------- access gate ----------

function showAccessGate(errorMessage) {
  const gate = document.getElementById("access-gate");
  const main = document.getElementById("app-main");
  gate.classList.remove("hidden");
  main.classList.add("app-locked");
  const errorEl = document.getElementById("access-error");
  if (errorMessage) {
    errorEl.textContent = errorMessage;
    errorEl.classList.remove("hidden");
  } else {
    errorEl.classList.add("hidden");
  }
  document.getElementById("access-code-input").focus();
}

function hideAccessGate() {
  document.getElementById("access-gate").classList.add("hidden");
  document.getElementById("app-main").classList.remove("app-locked");
}

async function tryAccessCode(code, { onFail } = {}) {
  const res = await fetch(apiUrl("/api/auth/check"), { headers: { "X-Access-Code": code } });
  if (res.ok) {
    sessionStorage.setItem(ACCESS_CODE_KEY, code);
    hideAccessGate();
    return true;
  }
  sessionStorage.removeItem(ACCESS_CODE_KEY);
  if (onFail) onFail();
  return false;
}

document.getElementById("access-form").addEventListener("submit", async (e) => {
  e.preventDefault();
  const input = document.getElementById("access-code-input");
  const btn = e.target.querySelector("button");
  setButtonLoading(btn, true, "Checking…", "Continue");
  const ok = await tryAccessCode(input.value, {
    onFail: () => {
      input.classList.remove("shake");
      void input.offsetWidth;
      input.classList.add("shake");
      document.getElementById("access-error").classList.remove("hidden");
    },
  });
  setButtonLoading(btn, false, "", "Continue");
  if (!ok) input.select();
});

// Validate a stored code on load (it may be stale from a previous session
// or a rotated ACCESS_CODE) rather than trusting it blindly.
(async () => {
  const stored = getAccessCode();
  if (stored) {
    await tryAccessCode(stored, { onFail: () => showAccessGate() });
  } else {
    showAccessGate();
  }
})();

// ---------- small motion helpers ----------

// Reveal a card with its entrance animation. Re-triggerable: strip the
// class first so re-showing a section (e.g. after re-analyzing) replays it.
function revealCard(el) {
  el.classList.remove("hidden");
  el.classList.remove("reveal");
  void el.offsetWidth; // force reflow so the animation restarts
  requestAnimationFrame(() => el.classList.add("reveal"));
}

function setStep(n, { complete = [] } = {}) {
  document.querySelectorAll("#steps .step").forEach((stepEl) => {
    const step = Number(stepEl.dataset.step);
    const isComplete = complete.includes(step);
    stepEl.classList.toggle("active", step === n);
    stepEl.classList.toggle("complete", isComplete);
    const dot = stepEl.querySelector(".step-dot");
    dot.textContent = isComplete ? "✓" : String(step);
  });
}

function setButtonLoading(btn, loading, loadingText, idleText) {
  btn.disabled = loading;
  const label = btn.querySelector(".btn-label") || btn;
  label.innerHTML = loading
    ? `<span class="spinner"></span> ${loadingText}`
    : idleText;
}

// ---------- upload dropzones (click + real drag-and-drop) ----------

function wireDropzone(zoneId, inputEl, filenameId, defaultText) {
  const zone = document.getElementById(zoneId);
  const filenameEl = document.getElementById(filenameId);

  const updateLabel = () => {
    const file = inputEl.files[0];
    filenameEl.textContent = file ? file.name : defaultText;
    filenameEl.classList.toggle("has-file", !!file);
    if (file) {
      zone.classList.remove("flash");
      void zone.offsetWidth;
      zone.classList.add("flash");
    }
  };

  inputEl.addEventListener("change", updateLabel);

  ["dragenter", "dragover"].forEach((evt) =>
    zone.addEventListener(evt, (e) => {
      e.preventDefault();
      zone.classList.add("dragover");
    })
  );
  ["dragleave", "dragend", "drop"].forEach((evt) =>
    zone.addEventListener(evt, () => zone.classList.remove("dragover"))
  );
  zone.addEventListener("drop", (e) => {
    e.preventDefault();
    const file = e.dataTransfer.files[0];
    if (!file) return;
    inputEl.files = e.dataTransfer.files;
    updateLabel();
  });
}

const omFile = document.getElementById("om-file");
const xlsxFile = document.getElementById("xlsx-file");
wireDropzone("om-dropzone", omFile, "om-filename", "Click or drag a .pdf or .pptx here");
wireDropzone("xlsx-dropzone", xlsxFile, "xlsx-filename", "Click or drag a .xlsx here");

document.getElementById("upload-form").addEventListener("submit", async (e) => {
  e.preventDefault();
  if (!omFile.files[0] && !xlsxFile.files[0]) {
    alert("Upload at least an OM PDF or an underwriting model.");
    return;
  }
  const btn = document.getElementById("analyze-btn");
  setButtonLoading(btn, true, "Analyzing…", "Analyze Documents");

  const fd = new FormData();
  if (omFile.files[0]) fd.append("om_file", omFile.files[0]);
  if (xlsxFile.files[0]) fd.append("xlsx_file", xlsxFile.files[0]);
  fd.append("analyst_notes_json", "{}");

  try {
    const res = await apiFetch("/api/deals", { method: "POST", body: fd });
    if (!res.ok) throw new Error(await res.text());
    const data = await res.json();
    state.dealId = data.deal_id;
    state.gapAnalysis = data.gap_analysis;
    state.images = data.images;
    renderGapTable();
    renderImages();
    revealCard(document.getElementById("gap-section"));
    if (state.images.length) {
      revealCard(document.getElementById("images-section"));
    } else {
      document.getElementById("images-section").classList.add("hidden");
    }
    revealCard(document.getElementById("generate-section"));
    setStep(2, { complete: [1] });
  } catch (err) {
    alert("Failed to analyze documents: " + err.message);
  } finally {
    setButtonLoading(btn, false, "", "Analyze Documents");
  }
});

function renderGapTable() {
  const container = document.getElementById("gap-table");
  container.innerHTML = "";

  // Headline counts required fields only. The registry also carries a lot of
  // optional detail (per-unit figures, Sources & Uses line items) that most
  // models don't supply, and folding those in would make the count read alarming
  // when the deck is actually fine.
  const required = state.gapAnalysis.filter((rf) => rf.required);
  const requiredFound = required.filter((rf) => rf.provenance !== "missing").length;
  const requiredMissing = required.length - requiredFound;
  const optionalMissing = state.gapAnalysis.filter(
    (rf) => !rf.required && rf.provenance === "missing"
  ).length;

  const summary = document.getElementById("gap-summary");
  summary.className = "gap-summary" + (requiredMissing > 0 ? " has-missing" : "");

  let text =
    requiredMissing === 0
      ? `All set — every required field was found (${requiredFound}/${required.length}).`
      : `${requiredFound} of ${required.length} required fields found — ${requiredMissing} still need your input below.`;
  if (optionalMissing) {
    text += ` ${optionalMissing} optional field${optionalMissing === 1 ? "" : "s"} will show as “TBD” on the deck.`;
  }
  summary.textContent = text;

  const groups = {};
  for (const rf of state.gapAnalysis) {
    (groups[rf.group] = groups[rf.group] || []).push(rf);
  }
  let rowIndex = 0;
  for (const [group, rows] of Object.entries(groups)) {
    const label = document.createElement("div");
    label.className = "gap-group-label";
    label.textContent = GROUP_LABELS[group] || group;
    container.appendChild(label);

    for (const rf of rows) {
      const row = document.createElement("div");
      row.className = "gap-row";
      row.style.animationDelay = `${Math.min(rowIndex * 25, 400)}ms`;
      rowIndex += 1;

      const labelEl = document.createElement("div");
      labelEl.className = "gap-label";
      labelEl.textContent = rf.label + (rf.required ? "" : " (optional)");
      row.appendChild(labelEl);

      const valueEl = document.createElement("div");
      valueEl.className = "gap-value" + (rf.provenance === "missing" ? " missing" : "");
      valueEl.textContent = rf.value ?? "Not found";
      row.appendChild(valueEl);

      const badge = document.createElement("div");
      badge.className = "badge " + rf.provenance;
      badge.textContent =
        rf.provenance === "om" ? "Found in OM" :
        rf.provenance === "model" ? "Found in Model" :
        rf.provenance === "analyst" ? "Analyst Input" : "Missing";
      row.appendChild(badge);

      container.appendChild(row);

      if (rf.provenance === "missing") {
        const inputRow = document.createElement("div");
        inputRow.className = "missing-input-row";
        const input = document.createElement("input");
        input.type = "text";
        input.placeholder = `Enter ${rf.label}...`;
        input.addEventListener("change", () => submitAnalystInput(rf.key, input.value, input));
        inputRow.appendChild(input);
        container.appendChild(inputRow);
      }
    }
  }
}

async function submitAnalystInput(key, value, inputEl) {
  if (!value) return;
  const res = await apiFetch(`/api/deals/${state.dealId}/analyst-inputs`, {
    method: "PATCH",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ [key]: value }),
  });
  const data = await res.json();
  state.gapAnalysis = data.gap_analysis;
  if (inputEl) inputEl.classList.add("saved");
  renderGapTable();
}

function renderImages() {
  const grid = document.getElementById("images-grid");
  grid.innerHTML = "";
  state.images.forEach((img, idx) => {
    const tile = document.createElement("div");
    tile.className = "img-tile" + (idx === 0 ? " selected" : "");
    tile.style.animationDelay = `${idx * 60}ms`;
    tile.innerHTML = `<img src="${mediaUrl(img.url)}" /> ${idx === 0 ? '<span class="star">&#9733;</span>' : ""}`;
    tile.addEventListener("click", async () => {
      const filename = img.url.split("/").pop();
      await apiFetch(`/api/deals/${state.dealId}/select-hero-image`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ filename }),
      });
      state.images = [img, ...state.images.filter((i) => i !== img)];
      renderImages();
    });
    grid.appendChild(tile);
  });
}

document.getElementById("generate-btn").addEventListener("click", async () => {
  if (!state.dealId) return;
  const btn = document.getElementById("generate-btn");
  const status = document.getElementById("generate-status");
  setButtonLoading(btn, true, "Generating…", "Generate Deck");
  status.innerHTML = '<span class="spinner"></span> Building your deck and rendering a preview — this can take a few seconds.';
  setStep(3, { complete: [1, 2] });

  try {
    const res = await apiFetch(`/api/deals/${state.dealId}/generate`, { method: "POST" });
    const data = await res.json();
    const grid = document.getElementById("preview-grid");
    grid.innerHTML = "";
    if (data.preview_images && data.preview_images.length) {
      data.preview_images.forEach((url, idx) => {
        const img = document.createElement("img");
        img.src = mediaUrl(url);
        img.style.animationDelay = `${idx * 120}ms`;
        grid.appendChild(img);
      });
      status.textContent = "Deck's ready — take a look below, or download the real .pptx.";
    } else {
      status.textContent = "Deck generated. Preview rendering unavailable in this environment (" +
        (data.preview_error || "unknown error") + ") -- download the PPTX to view it.";
    }
    const link = document.getElementById("download-link");
    link.href = mediaUrl(`/api/deals/${state.dealId}/download`);
    link.classList.remove("hidden");
    setStep(3, { complete: [1, 2, 3] });
  } catch (err) {
    status.textContent = "Failed to generate deck: " + err.message;
  } finally {
    setButtonLoading(btn, false, "", "Generate Deck");
  }
});
