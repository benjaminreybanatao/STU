const state = {
  dealId: null,
  gapAnalysis: [],
  images: [],
};

// window.API_BASE is set in config.js -- "" for same-origin (backend serves
// this frontend directly), or an absolute URL when this page is hosted
// separately (e.g. GitHub Pages) from the backend.
function apiUrl(path) {
  return (window.API_BASE || "") + path;
}

const GROUP_LABELS = {
  property: "Property",
  tenancy: "Tenancy",
  pricing: "Pricing",
  capital_stack: "Capital Stack",
  returns: "Returns",
  assumptions: "Assumptions",
};

const omFile = document.getElementById("om-file");
const xlsxFile = document.getElementById("xlsx-file");

omFile.addEventListener("change", () => {
  const el = document.getElementById("om-filename");
  el.textContent = omFile.files[0] ? omFile.files[0].name : "Click or drop a .pdf";
  el.classList.toggle("has-file", !!omFile.files[0]);
});
xlsxFile.addEventListener("change", () => {
  const el = document.getElementById("xlsx-filename");
  el.textContent = xlsxFile.files[0] ? xlsxFile.files[0].name : "Click or drop a .xlsx";
  el.classList.toggle("has-file", !!xlsxFile.files[0]);
});

document.getElementById("upload-form").addEventListener("submit", async (e) => {
  e.preventDefault();
  if (!omFile.files[0] && !xlsxFile.files[0]) {
    alert("Upload at least an OM PDF or an underwriting model.");
    return;
  }
  const btn = document.getElementById("analyze-btn");
  btn.disabled = true;
  btn.textContent = "Analyzing...";

  const fd = new FormData();
  if (omFile.files[0]) fd.append("om_file", omFile.files[0]);
  if (xlsxFile.files[0]) fd.append("xlsx_file", xlsxFile.files[0]);
  fd.append("analyst_notes_json", "{}");

  try {
    const res = await fetch(apiUrl("/api/deals"), { method: "POST", body: fd });
    if (!res.ok) throw new Error(await res.text());
    const data = await res.json();
    state.dealId = data.deal_id;
    state.gapAnalysis = data.gap_analysis;
    state.images = data.images;
    renderGapTable();
    renderImages();
    document.getElementById("gap-section").classList.remove("hidden");
    document.getElementById("images-section").classList.toggle("hidden", state.images.length === 0);
    document.getElementById("generate-section").classList.remove("hidden");
  } catch (err) {
    alert("Failed to analyze documents: " + err.message);
  } finally {
    btn.disabled = false;
    btn.textContent = "Analyze Documents";
  }
});

function renderGapTable() {
  const container = document.getElementById("gap-table");
  container.innerHTML = "";
  const groups = {};
  for (const rf of state.gapAnalysis) {
    (groups[rf.group] = groups[rf.group] || []).push(rf);
  }
  for (const [group, rows] of Object.entries(groups)) {
    const label = document.createElement("div");
    label.className = "gap-group-label";
    label.textContent = GROUP_LABELS[group] || group;
    container.appendChild(label);

    for (const rf of rows) {
      const row = document.createElement("div");
      row.className = "gap-row";

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
        input.addEventListener("change", () => submitAnalystInput(rf.key, input.value));
        inputRow.appendChild(input);
        container.appendChild(inputRow);
      }
    }
  }
}

async function submitAnalystInput(key, value) {
  if (!value) return;
  const res = await fetch(apiUrl(`/api/deals/${state.dealId}/analyst-inputs`), {
    method: "PATCH",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ [key]: value }),
  });
  const data = await res.json();
  state.gapAnalysis = data.gap_analysis;
  renderGapTable();
}

function renderImages() {
  const grid = document.getElementById("images-grid");
  grid.innerHTML = "";
  state.images.forEach((img, idx) => {
    const tile = document.createElement("div");
    tile.className = "img-tile" + (idx === 0 ? " selected" : "");
    tile.innerHTML = `<img src="${apiUrl(img.url)}" /> ${idx === 0 ? '<span class="star">&#9733;</span>' : ""}`;
    tile.addEventListener("click", async () => {
      const filename = img.url.split("/").pop();
      await fetch(apiUrl(`/api/deals/${state.dealId}/select-hero-image`), {
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
  btn.disabled = true;
  btn.textContent = "Generating...";
  status.textContent = "Building deck and rendering preview -- this can take a few seconds.";

  try {
    const res = await fetch(apiUrl(`/api/deals/${state.dealId}/generate`), { method: "POST" });
    const data = await res.json();
    const grid = document.getElementById("preview-grid");
    grid.innerHTML = "";
    if (data.preview_images && data.preview_images.length) {
      data.preview_images.forEach((url) => {
        const img = document.createElement("img");
        img.src = apiUrl(url);
        grid.appendChild(img);
      });
      status.textContent = "";
    } else {
      status.textContent = "Deck generated. Preview rendering unavailable in this environment (" +
        (data.preview_error || "unknown error") + ") -- download the PPTX to view it.";
    }
    const link = document.getElementById("download-link");
    link.href = apiUrl(`/api/deals/${state.dealId}/download`);
    link.classList.remove("hidden");
  } catch (err) {
    status.textContent = "Failed to generate deck: " + err.message;
  } finally {
    btn.disabled = false;
    btn.textContent = "Generate Deck";
  }
});
