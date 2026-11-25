// PANG & PPT Studio - Application Controller (Direct Native Bridge + Fallback)

const API_BASE = (window.location && window.location.origin && window.location.origin.startsWith("http"))
  ? `${window.location.origin}/api`
  : "http://127.0.0.1:8080/api";

const state = {
  currentFile: null,
  processedFile: null,
  lastExportPath: null,
  fileMeta: null,
  activeWorkspace: "optimizer",
  currentPage: 1,
  totalPages: 1,
  activePreset: "balanced",
  viewMode: "single",
  zoom: 100,
  pageObserver: null
};

// DOM References
const elements = {
  fileInput: document.getElementById("file-input-dialog"),
  dropZone: document.getElementById("drop-zone"),
  documentView: document.getElementById("document-view"),
  stageScroll: document.getElementById("stage-scroll"),
  artboardWrapper: document.getElementById("artboard-wrapper"),
  lblFilename: document.getElementById("lbl-filename"),
  lblFilesize: document.getElementById("lbl-filesize"),
  currentPageNum: document.getElementById("current-page-num"),
  totalPagesNum: document.getElementById("total-pages-num"),
  btnPrevPage: document.getElementById("btn-prev-page"),
  btnNextPage: document.getElementById("btn-next-page"),
  btnViewSingle: document.getElementById("btn-view-single"),
  btnViewSplit: document.getElementById("btn-view-split"),
  btnZoomIn: document.getElementById("btn-zoom-in"),
  btnZoomOut: document.getElementById("btn-zoom-out"),
  btnZoomFit: document.getElementById("btn-zoom-fit"),
  zoomValue: document.getElementById("zoom-value"),
  btnExecuteAction: document.getElementById("btn-execute-action"),
  btnActionLabel: document.getElementById("btn-action-label"),
  progressContainer: document.getElementById("process-progress-container"),
  progressFill: document.getElementById("process-progress-fill"),
  progressText: document.getElementById("process-progress-text"),
  exportToast: document.getElementById("export-toast"),
  toastTitle: document.getElementById("toast-title"),
  toastSub: document.getElementById("toast-sub"),
  btnDownloadExport: document.getElementById("btn-download-export"),
  btnOpenLocation: document.getElementById("btn-open-location"),
  btnCloseToast: document.getElementById("btn-close-toast"),
  footerDimensions: document.getElementById("footer-dimensions"),
  footerVectors: document.getElementById("footer-vectors"),
  footerFonts: document.getElementById("footer-fonts"),
  metaArtboard: document.getElementById("meta-artboard"),
  metaPages: document.getElementById("meta-pages"),
  metaFontsCount: document.getElementById("meta-fonts-count"),
  metaImagesCount: document.getElementById("meta-images-count"),
  metaVectorsCount: document.getElementById("meta-vectors-count"),
  metaTransparency: document.getElementById("meta-transparency"),
  fontListContainer: document.getElementById("font-list-container"),
  alertsContainer: document.getElementById("alerts-container"),
  selDpi: document.getElementById("sel-dpi"),
  sliderJpegQuality: document.getElementById("slider-jpeg-quality"),
  valJpegQuality: document.getElementById("val-jpeg-quality"),
  togglePreservePng: document.getElementById("toggle-preserve-png"),
  toggleRepairTransparency: document.getElementById("toggle-repair-transparency"),
  toggleFlattenProblematic: document.getElementById("toggle-flatten-problematic"),
  toggleCleanMetadata: document.getElementById("toggle-clean-metadata"),
  selPptMode: document.getElementById("sel-ppt-mode")
};

const PRESETS = {
  balanced: { dpi: "150", quality: 82, preservePng: true },
  max_quality: { dpi: "300", quality: 92, preservePng: true },
  mobile_friendly: { dpi: "120", quality: 75, preservePng: false },
  smallest_size: { dpi: "96", quality: 65, preservePng: false },
  print_ready: { dpi: "300", quality: 95, preservePng: true }
};

// Bridge Caller helper (calls pywebview js_api if available, else REST API)
async function callBridge(method, ...args) {
  if (window.pywebview && window.pywebview.api && window.pywebview.api[method]) {
    return await window.pywebview.api[method](...args);
  }
  // REST fallback
  return null;
}

function init() {
  setupNavigation();
  setupPresets();
  setupFileHandling();
  setupPreviewControls();
  setupSliders();
  setupTabs();
  setupActionExecution();
  setupExportToast();
}

function setupNavigation() {
  document.querySelectorAll(".mode-btn").forEach(tab => {
    tab.addEventListener("click", () => {
      document.querySelectorAll(".mode-btn").forEach(t => t.classList.remove("active"));
      tab.classList.add("active");
      state.activeWorkspace = tab.dataset.workspace;
      updateWorkspaceUI();
    });
  });
}

function updateWorkspaceUI() {
  switch (state.activeWorkspace) {
    case "optimizer":
      elements.btnActionLabel.textContent = "Compress & Optimize PDF";
      break;
    case "repair":
      elements.btnActionLabel.textContent = "Run PDF → PNG → PDF (Fix Effects & Slash 70% Size)";
      break;
    case "pdf2ppt":
      elements.btnActionLabel.textContent = "Convert to Editable PowerPoint";
      break;
    case "ppt2pdf":
      elements.btnActionLabel.textContent = "Convert PowerPoint to PDF";
      break;
    case "batch":
      elements.btnActionLabel.textContent = "Start Batch Processing";
      break;
  }
}

function setupPresets() {
  document.querySelectorAll(".preset-item").forEach(card => {
    card.addEventListener("click", () => {
      document.querySelectorAll(".preset-item").forEach(c => c.classList.remove("active"));
      card.classList.add("active");
      const presetKey = card.dataset.preset;
      state.activePreset = presetKey;
      const conf = PRESETS[presetKey];
      if (conf) {
        elements.selDpi.value = conf.dpi;
        elements.sliderJpegQuality.value = conf.quality;
        elements.valJpegQuality.textContent = `${conf.quality}%`;
        syncSliderFill(elements.sliderJpegQuality);
        elements.togglePreservePng.checked = conf.preservePng;
      }
    });
  });
}

function setupFileHandling() {
  const fileInput = document.getElementById("file-input-dialog");

  fileInput.addEventListener("change", async (e) => {
    if (e.target.files && e.target.files.length > 0) {
      const file = e.target.files[0];
      fileInput.value = "";
      await handleIncomingFile(file);
    }
  });

  const dropZone = document.getElementById("drop-zone");
  document.body.addEventListener("dragover", (e) => {
    e.preventDefault();
    dropZone.classList.add("drag-over");
  });

  document.body.addEventListener("dragleave", (e) => {
    if (e.target === document.body || e.relatedTarget === null) {
      dropZone.classList.remove("drag-over");
    }
  });

  document.body.addEventListener("drop", async (e) => {
    e.preventDefault();
    dropZone.classList.remove("drag-over");
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      const file = e.dataTransfer.files[0];
      await handleIncomingFile(file);
    }
  });
}

async function handleIncomingFile(file) {
  if (file.path && window.pywebview && window.pywebview.api) {
    await inspectFile(file.path);
    return;
  }

  // Real-time Dock Upload HUD
  const dockHud = document.getElementById("dock-upload-hud");
  const dockFilename = document.getElementById("dock-hud-filename");
  const dockPercentage = document.getElementById("dock-hud-percentage");
  const dockBar = document.getElementById("dock-hud-bar");
  const dockBytes = document.getElementById("dock-hud-bytes");
  const dockSpeed = document.getElementById("dock-hud-speed");

  dockFilename.textContent = `Uploading ${file.name}`;
  dockPercentage.textContent = "0%";
  dockBar.style.width = "0%";
  dockBytes.textContent = `0 MB / ${formatBytes(file.size)}`;
  dockSpeed.textContent = "Streaming to local engine...";
  dockHud.style.display = "flex";

  const startTime = Date.now();

  try {
    const xhr = new XMLHttpRequest();
    xhr.open("POST", `${API_BASE}/upload`, true);
    xhr.setRequestHeader("Content-Type", "application/octet-stream");
    xhr.setRequestHeader("X-Filename", encodeURIComponent(file.name));

    xhr.upload.onprogress = (e) => {
      if (e.lengthComputable) {
        const pct = Math.round((e.loaded / e.total) * 100);
        dockBar.style.width = `${pct}%`;
        dockPercentage.textContent = `${pct}%`;
        dockBytes.textContent = `${formatBytes(e.loaded)} / ${formatBytes(e.total)}`;

        const elapsedSec = (Date.now() - startTime) / 1000;
        if (elapsedSec > 0.05) {
          const speedMB = (e.loaded / (1024 * 1024)) / elapsedSec;
          dockSpeed.textContent = `${speedMB.toFixed(1)} MB/s`;
        }
      }
    };

    xhr.onload = async () => {
      if (xhr.status === 200) {
        dockPercentage.textContent = "100%";
        dockBar.style.width = "100%";
        dockSpeed.textContent = "Analyzing document structure...";
        try {
          const data = JSON.parse(xhr.responseText);
          if (data.success && data.file_path) {
            await inspectFile(data.file_path);
          } else {
            await inspectFile(file.name);
          }
        } catch (ex) {
          console.error("Upload parse error:", ex);
        } finally {
          dockHud.style.display = "none";
        }
      } else {
        dockSpeed.textContent = "Upload failed (" + xhr.status + ")";
        setTimeout(() => { dockHud.style.display = "none"; }, 2500);
      }
    };

    xhr.onerror = () => {
      dockSpeed.textContent = "Connection error";
      setTimeout(() => { dockHud.style.display = "none"; }, 2500);
    };

    // Send binary Blob directly
    xhr.send(file);

  } catch (err) {
    console.error("Upload error:", err);
    dockHud.style.display = "none";
  }
}

async function inspectFile(filePath) {
  try {
    let result = null;
    if (window.pywebview && window.pywebview.api && window.pywebview.api.inspect_file) {
      result = await callBridge("inspect_file", filePath);
    } else {
      const res = await fetch(`${API_BASE}/inspect`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ file_path: filePath })
      });
      result = await res.json();
    }

    const data = (result && result.success && result.data) ? result.data : null;
    if (!data) {
      console.error("Inspect failed:", result);
      return;
    }

    state.currentFile = filePath;
    state.processedFile = null;
    state.fileMeta = data;
    state.currentPage = 1;
    state.totalPages = data.page_count || 1;

    elements.btnViewSingle.classList.add("active");
    elements.btnViewSplit.classList.remove("active");
    state.viewMode = "single";

    elements.lblFilename.textContent = data.file_name;
    elements.lblFilesize.textContent = formatBytes(data.file_size);
    elements.currentPageNum.textContent = state.currentPage;
    elements.totalPagesNum.textContent = state.totalPages;

    elements.metaArtboard.textContent = `${data.dimensions.width} × ${data.dimensions.height} pt`;
    elements.metaPages.textContent = data.page_count;
    elements.metaFontsCount.textContent = data.fonts_count;
    elements.metaImagesCount.textContent = data.images_count;
    elements.metaVectorsCount.textContent = data.vector_paths_count;
    elements.metaTransparency.textContent = data.has_transparency ? "Detected (Alpha/Blend)" : "None";

    elements.footerDimensions.textContent = `Artboard: ${data.dimensions.width} × ${data.dimensions.height} pt`;
    elements.footerVectors.textContent = `Vectors: ${data.vector_paths_count}`;
    elements.footerFonts.textContent = `Fonts: ${data.fonts_count}`;

    elements.fontListContainer.innerHTML = "";
    if (data.fonts && data.fonts.length > 0) {
      data.fonts.forEach(f => {
        const pill = document.createElement("span");
        pill.className = "font-pill";
        pill.textContent = f;
        elements.fontListContainer.appendChild(pill);
      });
    } else {
      elements.fontListContainer.innerHTML = `<span class="font-pill">Embedded Curves / Shapes</span>`;
    }

    elements.alertsContainer.innerHTML = "";
    if (data.has_transparency) {
      const al = document.createElement("div");
      al.className = "studio-alert warning";
      al.textContent = "Adobe Illustrator transparency/opacity detected. 'Repair & Flatten' recommended for mobile PDF viewers.";
      elements.alertsContainer.appendChild(al);
    }
    if (data.problematic_effects && data.problematic_effects.length > 0) {
      const al = document.createElement("div");
      al.className = "studio-alert warning";
      al.textContent = `Complex blend modes detected: ${data.problematic_effects.join(", ")}`;
      elements.alertsContainer.appendChild(al);
    }
    if (!data.has_transparency && (!data.problematic_effects || data.problematic_effects.length === 0)) {
      const al = document.createElement("div");
      al.className = "studio-alert success";
      al.textContent = "Document structure is clean and standard-compliant.";
      elements.alertsContainer.appendChild(al);
    }

    elements.dropZone.style.display = "none";
    elements.documentView.style.display = "flex";

    buildScrollablePages(state.totalPages);

  } catch (err) {
    console.error("Inspect error:", err);
  }
}

function buildScrollablePages(totalPages) {
  if (state.pageObserver) {
    state.pageObserver.disconnect();
    state.pageObserver = null;
  }

  elements.artboardWrapper.innerHTML = "";
  elements.stageScroll.scrollTop = 0;

  for (let p = 1; p <= totalPages; p++) {
    const row = document.createElement("div");
    row.className = "page-row";
    row.dataset.page = String(p);

    const badge = document.createElement("div");
    badge.className = "page-row-badge";
    badge.textContent = `Page ${p} of ${totalPages}`;
    row.appendChild(badge);

    const frames = document.createElement("div");
    frames.className = "page-frames";

    // Original frame
    const origFrame = document.createElement("div");
    origFrame.className = "preview-frame";
    origFrame.innerHTML = `<span class="frame-label">SOURCE</span><div class="preview-placeholder">Loading page ${p}…</div><img class="preview-img preview-orig" data-page="${p}" alt="Page ${p}" style="display:none">`;
    frames.appendChild(origFrame);

    // Processed comparison frame
    const optFrame = document.createElement("div");
    optFrame.className = "preview-frame compare-frame";
    optFrame.style.display = "none";
    optFrame.innerHTML = `<span class="frame-label accent">PROCESSED</span><div class="preview-placeholder preview-opt-placeholder">Run Optimize or Fix Effects to compare</div><img class="preview-img preview-opt" data-page="${p}" alt="Processed Page ${p}" style="display:none">`;
    frames.appendChild(optFrame);

    row.appendChild(frames);
    elements.artboardWrapper.appendChild(row);
  }

  // Lazy load page previews as they enter or approach viewport
  state.pageObserver = new IntersectionObserver((entries) => {
    entries.forEach(entry => {
      if (!entry.isIntersecting) return;
      const row = entry.target;
      const page = parseInt(row.dataset.page, 10);
      loadRowPreview(row, page);
    });
  }, {
    root: elements.stageScroll,
    rootMargin: "600px 0px",
    threshold: 0.01
  });

  elements.artboardWrapper.querySelectorAll(".page-row").forEach(row => {
    state.pageObserver.observe(row);
  });

  updateZoom();
}

async function loadRowPreview(row, page) {
  if (!state.currentFile) return;

  const imgOrig = row.querySelector(".preview-orig");
  const placeholder = row.querySelector(".preview-placeholder");

  if (imgOrig && !imgOrig.dataset.loaded) {
    imgOrig.dataset.loaded = "1";
    imgOrig.onload = () => {
      if (placeholder) placeholder.style.display = "none";
      imgOrig.style.display = "block";
    };

    if (window.pywebview && window.pywebview.api && window.pywebview.api.get_preview_data_url) {
      try {
        const res = await callBridge("get_preview_data_url", state.currentFile, page, 150);
        if (res && res.success && res.data_url) {
          imgOrig.src = res.data_url;
        }
      } catch (e) {
        console.error("Bridge preview error:", e);
      }
    } else {
      imgOrig.src = `${API_BASE}/preview?file=${encodeURIComponent(state.currentFile)}&page=${page}&dpi=150&t=${Date.now()}`;
    }
  }

  // Also load processed comparison if active
  if (state.viewMode === "split" && state.processedFile) {
    const imgOpt = row.querySelector(".preview-opt");
    const optPlaceholder = row.querySelector(".preview-opt-placeholder");
    if (imgOpt && !imgOpt.dataset.loaded) {
      imgOpt.dataset.loaded = "1";
      if (optPlaceholder) optPlaceholder.textContent = `Rendering processed page ${page}…`;
      imgOpt.onload = () => {
        if (optPlaceholder) optPlaceholder.style.display = "none";
        imgOpt.style.display = "block";
      };
      if (window.pywebview && window.pywebview.api && window.pywebview.api.get_preview_data_url) {
        const compRes = await callBridge("get_preview_data_url", state.processedFile, page, 150);
        if (compRes && compRes.success) imgOpt.src = compRes.data_url;
      } else {
        imgOpt.src = `${API_BASE}/preview?file=${encodeURIComponent(state.processedFile)}&page=${page}&dpi=150&t=${Date.now()}`;
      }
    }
  }
}

function scrollToPage(page) {
  const row = elements.artboardWrapper.querySelector(`.page-row[data-page="${page}"]`);
  if (row) {
    row.scrollIntoView({ behavior: "smooth", block: "start" });
  }
}

function setupPreviewControls() {
  // Track current visible page on scroll
  let scrollTick = false;
  elements.stageScroll.addEventListener("scroll", () => {
    if (scrollTick) return;
    scrollTick = true;
    requestAnimationFrame(() => {
      scrollTick = false;
      const rows = elements.artboardWrapper.querySelectorAll(".page-row");
      const containerTop = elements.stageScroll.getBoundingClientRect().top;
      let bestPage = state.currentPage;
      let bestDist = Infinity;
      rows.forEach(r => {
        const dist = Math.abs(r.getBoundingClientRect().top - containerTop - 32);
        if (dist < bestDist) {
          bestDist = dist;
          bestPage = parseInt(r.dataset.page, 10);
        }
      });
      if (bestPage !== state.currentPage) {
        state.currentPage = bestPage;
        elements.currentPageNum.textContent = bestPage;
      }
    });
  });

  elements.btnPrevPage.addEventListener("click", () => {
    if (state.currentPage > 1) {
      state.currentPage--;
      elements.currentPageNum.textContent = state.currentPage;
      scrollToPage(state.currentPage);
    }
  });

  elements.btnNextPage.addEventListener("click", () => {
    if (state.currentPage < state.totalPages) {
      state.currentPage++;
      elements.currentPageNum.textContent = state.currentPage;
      scrollToPage(state.currentPage);
    }
  });

  elements.btnViewSingle.addEventListener("click", () => {
    state.viewMode = "single";
    elements.btnViewSingle.classList.add("active");
    elements.btnViewSplit.classList.remove("active");
    elements.artboardWrapper.querySelectorAll(".compare-frame").forEach(f => {
      f.style.display = "none";
    });
  });

  elements.btnViewSplit.addEventListener("click", () => {
    state.viewMode = "split";
    elements.btnViewSplit.classList.add("active");
    elements.btnViewSingle.classList.remove("active");
    elements.artboardWrapper.querySelectorAll(".compare-frame").forEach(f => {
      f.style.display = "flex";
    });
    if (state.processedFile && state.pageObserver) {
      elements.artboardWrapper.querySelectorAll(".page-row").forEach(row => {
        state.pageObserver.unobserve(row);
        state.pageObserver.observe(row);
      });
    }
  });

  elements.btnZoomIn.addEventListener("click", () => {
    state.zoom = Math.min(300, state.zoom + 25);
    updateZoom();
  });

  elements.btnZoomOut.addEventListener("click", () => {
    state.zoom = Math.max(50, state.zoom - 25);
    updateZoom();
  });

  elements.btnZoomFit.addEventListener("click", () => {
    state.zoom = 100;
    updateZoom();
  });
}

function updateZoom() {
  elements.zoomValue.textContent = `${state.zoom}%`;
  const maxW = Math.round(78 * (state.zoom / 100));
  elements.artboardWrapper.querySelectorAll(".preview-img").forEach(img => {
    img.style.maxWidth = `${maxW}vw`;
  });
}

function syncSliderFill(slider) {
  const min = parseFloat(slider.min || 0);
  const max = parseFloat(slider.max || 100);
  const val = parseFloat(slider.value || 0);
  const pct = ((val - min) / (max - min)) * 100;
  slider.style.setProperty("--val", `${pct}%`);
}

function setupSliders() {
  syncSliderFill(elements.sliderJpegQuality);
  elements.sliderJpegQuality.addEventListener("input", (e) => {
    elements.valJpegQuality.textContent = `${e.target.value}%`;
    syncSliderFill(e.target);
  });
}

function setupTabs() {
  document.querySelectorAll(".itab").forEach(tab => {
    tab.addEventListener("click", () => {
      document.querySelectorAll(".itab").forEach(t => t.classList.remove("active"));
      tab.classList.add("active");
      const targetId = tab.dataset.tab === "controls" ? "itab-controls" : "itab-analysis";
      document.getElementById("itab-controls").style.display = targetId === "itab-controls" ? "flex" : "none";
      document.getElementById("itab-analysis").style.display = targetId === "itab-analysis" ? "flex" : "none";
    });
  });

  document.querySelectorAll(".accordion-header").forEach(hdr => {
    hdr.addEventListener("click", () => {
      const item = hdr.parentElement;
      item.classList.toggle("open");
      const body = item.querySelector(".accordion-body");
      body.style.display = item.classList.contains("open") ? "flex" : "none";
    });
  });
}

function setupActionExecution() {
  elements.btnExecuteAction.addEventListener("click", async () => {
    if (!state.currentFile) return;

    elements.progressContainer.style.display = "flex";
    elements.progressFill.style.width = "40%";
    elements.progressText.textContent = "Processing document...";

    const settings = {
      dpi: parseInt(elements.selDpi.value),
      jpeg_quality: parseInt(elements.sliderJpegQuality.value),
      preserve_png: elements.togglePreservePng.checked,
      clean_metadata: elements.toggleCleanMetadata.checked,
      downsample_images: true,
      deflate_streams: true
    };

    let result = null;

    if (window.pywebview && window.pywebview.api) {
      if (state.activeWorkspace === "optimizer") {
        result = await callBridge("optimize_pdf", state.currentFile, null, settings);
      } else if (state.activeWorkspace === "repair") {
        result = await callBridge("repair_pdf", state.currentFile, null, {
          raster_dpi: parseInt(elements.selDpi.value),
          selective_flatten: elements.toggleFlattenProblematic.checked
        });
      } else if (state.activeWorkspace === "pdf2ppt") {
        result = await callBridge("pdf_to_ppt", state.currentFile, null, elements.selPptMode.value);
      } else if (state.activeWorkspace === "ppt2pdf") {
        result = await callBridge("ppt_to_pdf", state.currentFile, null);
      }
    } else {
      // Localhost HTTP API execution
      let endpoint = "/optimize";
      let postSettings = settings;
      if (state.activeWorkspace === "repair") {
        endpoint = "/repair";
        postSettings = {
          raster_dpi: parseInt(elements.selDpi.value),
          selective_flatten: elements.toggleFlattenProblematic.checked
        };
      } else if (state.activeWorkspace === "pdf2ppt") {
        endpoint = "/pdf2ppt";
        postSettings = { mode: elements.selPptMode.value };
      } else if (state.activeWorkspace === "ppt2pdf") {
        endpoint = "/ppt2pdf";
        postSettings = {};
      }

      const res = await fetch(`${API_BASE}${endpoint}`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          file_path: state.currentFile,
          settings: postSettings
        })
      });
      const data = await res.json();
      if (data.job_id) {
        pollJobStatus(data.job_id);
        return;
      }
    }

    elements.progressFill.style.width = "100%";
    if (result && result.success) {
      elements.progressText.textContent = "Operation Completed!";
      const outPath = result.result && result.result.output_path;
      if (outPath && outPath.endsWith(".pdf")) {
        state.processedFile = outPath;
        elements.artboardWrapper.querySelectorAll(".preview-opt").forEach(img => { delete img.dataset.loaded; });
        elements.btnViewSplit.click();
      }
      showExportNotification(result.result);
    } else {
      elements.progressText.textContent = "Error: " + (result ? result.error : "Failed");
    }

    setTimeout(() => {
      elements.progressContainer.style.display = "none";
    }, 4000);
  });
}

function setupExportToast() {
  elements.btnCloseToast.addEventListener("click", () => {
    elements.exportToast.style.display = "none";
  });

  elements.btnOpenLocation.addEventListener("click", async () => {
    if (!state.lastExportPath) return;
    if (window.pywebview && window.pywebview.api && window.pywebview.api.open_location) {
      await callBridge("open_location", state.lastExportPath);
    } else {
      await fetch(`${API_BASE}/open-location`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ file_path: state.lastExportPath })
      });
    }
  });
}

function showExportNotification(res) {
  if (!res || !res.output_path) return;
  state.lastExportPath = res.output_path;
  const fileName = res.output_path.split(/[\\/]/).pop();

  let details = fileName;
  if (res.new_size) {
    details += ` · ${formatBytes(res.new_size)}`;
    if (res.reduction_percent !== undefined) {
      details += ` (-${res.reduction_percent}%)`;
    }
  } else if (res.output_size) {
    details += ` · ${formatBytes(res.output_size)}`;
  }

  if (!window.pywebview) {
    elements.btnDownloadExport.href = `${API_BASE}/download?file=${encodeURIComponent(res.output_path)}`;
    elements.btnDownloadExport.setAttribute("download", fileName);
    elements.btnDownloadExport.style.display = "inline-flex";
  }

  elements.toastTitle.textContent = "Export Complete";
  elements.toastSub.textContent = details;
  elements.exportToast.style.display = "flex";
}

function pollJobStatus(jobId) {
  if (state.pollTimer) clearInterval(state.pollTimer);

  state.pollTimer = setInterval(async () => {
    try {
      const res = await fetch(`${API_BASE}/job-status?id=${jobId}`);
      const job = await res.json();

      elements.progressFill.style.width = `${job.progress}%`;
      elements.progressText.textContent = `${job.message} (${job.progress}%)`;

      if (job.status === "completed") {
        clearInterval(state.pollTimer);
        elements.progressText.textContent = "Done!";

        if (job.result && job.result.output_path && job.result.output_path.endsWith(".pdf")) {
          state.processedFile = job.result.output_path;
          elements.artboardWrapper.querySelectorAll(".preview-opt").forEach(img => { delete img.dataset.loaded; });
          elements.btnViewSplit.click();
        }

        showExportNotification(job.result);

        setTimeout(() => {
          elements.progressContainer.style.display = "none";
        }, 4000);
      } else if (job.status === "failed") {
        clearInterval(state.pollTimer);
        elements.progressText.textContent = `Error: ${job.error}`;
      }
    } catch (err) {
      console.error("Error polling job status:", err);
    }
  }, 600);
}

function formatBytes(bytes, decimals = 1) {
  if (bytes === 0) return "0 B";
  const k = 1024;
  const dm = decimals < 0 ? 0 : decimals;
  const sizes = ["B", "KB", "MB", "GB"];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(dm)) + " " + sizes[i];
}

window.addEventListener("DOMContentLoaded", init);
