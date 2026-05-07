/* PI 검토 툴 — main.js */
"use strict";

// ── Lucide 아이콘 초기화 ──────────────────────────────────────
document.addEventListener("DOMContentLoaded", () => {
  if (window.lucide) lucide.createIcons();
  init();
});

// ── 상태 ─────────────────────────────────────────────────────
let currentJobId  = null;
let pollTimer     = null;
let _pendingZip   = null;   // File (ZIP mode)
let _pendingFiles = null;   // FileList/Array (folder mode)
let _currentMode  = "zip";  // "zip" | "folder"

// ── 뷰 전환 ──────────────────────────────────────────────────
function showView(name) {
  document.querySelectorAll(".view").forEach(v => v.classList.remove("active"));
  const el = document.getElementById(`${name}-view`);
  if (el) el.classList.add("active");
  if (window.lucide) lucide.createIcons();
}

// ── 업로드 초기화 ─────────────────────────────────────────────
function init() {
  const zoneZip    = document.getElementById("zone-zip");
  const zoneFld    = document.getElementById("zone-folder");
  const fileInput  = document.getElementById("file-input");
  const fldInput   = document.getElementById("folder-input");
  const browseZip  = document.getElementById("browse-zip-btn");
  const browseFld  = document.getElementById("browse-folder-btn");
  const startBtn   = document.getElementById("start-btn");
  const clearBtn   = document.getElementById("clear-file-btn");
  const resetBtn   = document.getElementById("reset-btn");
  const tabZip     = document.getElementById("tab-zip");
  const tabFld     = document.getElementById("tab-folder");

  // 탭 전환
  tabZip.addEventListener("click", () => switchMode("zip"));
  tabFld.addEventListener("click", () => switchMode("folder"));

  // ZIP 영역
  zoneZip.addEventListener("click", e => { if (!e.target.closest("button")) fileInput.click(); });
  browseZip.addEventListener("click", e => { e.stopPropagation(); fileInput.click(); });
  fileInput.addEventListener("change", e => {
    const f = e.target.files[0];
    if (f) onZipSelected(f);
    fileInput.value = "";
  });
  setupZipDrop(zoneZip);

  // 폴더 영역
  zoneFld.addEventListener("click", e => { if (!e.target.closest("button")) fldInput.click(); });
  browseFld.addEventListener("click", e => { e.stopPropagation(); fldInput.click(); });
  fldInput.addEventListener("change", e => {
    const files = Array.from(e.target.files);
    if (files.length) onFolderSelected(files, files[0].webkitRelativePath.split("/")[0] || "폴더");
    fldInput.value = "";
  });
  setupFolderDrop(zoneFld);

  // 시작 / 취소
  startBtn.addEventListener("click", doUpload);
  clearBtn.addEventListener("click", clearSelection);

  resetBtn.addEventListener("click", resetAll);
  showView("upload");
}

// ── 모드 전환 ────────────────────────────────────────────────
function switchMode(mode) {
  _currentMode = mode;
  clearSelection();

  document.getElementById("tab-zip").classList.toggle("active", mode === "zip");
  document.getElementById("tab-folder").classList.toggle("active", mode === "folder");
  document.getElementById("zone-zip").style.display    = mode === "zip"    ? "" : "none";
  document.getElementById("zone-folder").style.display = mode === "folder" ? "" : "none";
}

// ── 선택 처리 ────────────────────────────────────────────────
function onZipSelected(file) {
  if (!file.name.toLowerCase().endsWith(".zip")) {
    alert("ZIP 파일을 선택해주세요.");
    return;
  }
  _pendingZip   = file;
  _pendingFiles = null;
  showPreview(file.name, `${(file.size / 1024 / 1024).toFixed(1)} MB`);
}

function onFolderSelected(files, folderName) {
  _pendingFiles = files;
  _pendingZip   = null;
  const exts = files.map(f => f.name.split(".").pop().toLowerCase());
  const pdfs  = exts.filter(e => e === "pdf").length;
  const xlsx  = exts.filter(e => e === "xlsx" || e === "xls").length;
  showPreview(folderName, `파일 ${files.length}개 (PDF ${pdfs}개, 엑셀 ${xlsx}개)`);
}

function showPreview(name, hint) {
  document.getElementById("preview-name").textContent  = name;
  document.getElementById("preview-count").textContent = hint;
  document.getElementById("file-preview").style.display = "";
  if (window.lucide) lucide.createIcons();
}

function clearSelection() {
  _pendingZip = _pendingFiles = null;
  document.getElementById("file-preview").style.display = "none";
}

// ── 드래그앤드롭 설정 ─────────────────────────────────────────
function setupZipDrop(zone) {
  ["dragenter", "dragover"].forEach(evt =>
    zone.addEventListener(evt, e => { e.preventDefault(); zone.classList.add("drag-over"); })
  );
  ["dragleave", "drop"].forEach(evt =>
    zone.addEventListener(evt, e => { e.preventDefault(); zone.classList.remove("drag-over"); })
  );
  zone.addEventListener("drop", e => {
    const f = e.dataTransfer.files[0];
    if (f) onZipSelected(f);
  });
}

function setupFolderDrop(zone) {
  ["dragenter", "dragover"].forEach(evt =>
    zone.addEventListener(evt, e => { e.preventDefault(); zone.classList.add("drag-over"); })
  );
  ["dragleave", "drop"].forEach(evt =>
    zone.addEventListener(evt, e => { e.preventDefault(); zone.classList.remove("drag-over"); })
  );
  zone.addEventListener("drop", async e => {
    const items = Array.from(e.dataTransfer.items);
    const entry = items[0]?.webkitGetAsEntry?.();
    if (entry?.isDirectory) {
      zone.classList.add("drag-over");
      const files = await readDirEntries(entry);
      zone.classList.remove("drag-over");
      if (files.length) onFolderSelected(files, entry.name);
    } else if (e.dataTransfer.files[0]) {
      // 폴더 탭에 ZIP 드롭 → ZIP 탭으로 자동 전환
      switchMode("zip");
      onZipSelected(e.dataTransfer.files[0]);
    }
  });
}

async function readDirEntries(dirEntry, prefix) {
  prefix = prefix || dirEntry.name;
  const files = [];
  const reader = dirEntry.createReader();
  let batch;
  do {
    batch = await new Promise((res, rej) => reader.readEntries(res, rej));
    for (const entry of batch) {
      if (entry.isFile) {
        const file = await new Promise((res, rej) => entry.file(res, rej));
        // webkitRelativePath는 읽기전용이므로 래퍼로 상대 경로 추가
        Object.defineProperty(file, "webkitRelativePath", {
          value: `${prefix}/${entry.name}`,
          writable: false,
        });
        files.push(file);
      } else if (entry.isDirectory) {
        const sub = await readDirEntries(entry, `${prefix}/${entry.name}`);
        files.push(...sub);
      }
    }
  } while (batch.length > 0);
  return files;
}

// ── 업로드 실행 ──────────────────────────────────────────────
async function doUpload() {
  showView("progress");
  setProgress(0, "파일 업로드 중...");

  const fd = new FormData();

  if (_currentMode === "zip" && _pendingZip) {
    addLog(`파일: ${_pendingZip.name} (${(_pendingZip.size/1024/1024).toFixed(1)} MB)`);
    fd.append("file", _pendingZip);
  } else if (_currentMode === "folder" && _pendingFiles?.length) {
    addLog(`폴더 업로드: ${_pendingFiles.length}개 파일`);
    for (const f of _pendingFiles) {
      fd.append("files", f, f.webkitRelativePath || f.name);
    }
  } else {
    showError("선택된 파일이 없습니다.");
    return;
  }

  try {
    const res  = await fetch("/upload", { method: "POST", body: fd });
    const data = await res.json();
    if (data.error) { showError(data.error); return; }
    currentJobId = data.job_id;
    addLog("업로드 완료, 분석 시작...");
    connectStatus(currentJobId);
  } catch (err) {
    showError("서버 연결에 실패했습니다: " + err.message);
  }
}

// ── 상태 연결 (SSE → polling 폴백) ──────────────────────────
function connectStatus(jobId) {
  if (window.EventSource) {
    const es = new EventSource(`/status/${jobId}`);
    es.onmessage = e => {
      const state = JSON.parse(e.data);
      handleState(state, jobId);
      if (state.status === "done" || state.status === "error") es.close();
    };
    es.onerror = () => { es.close(); startPolling(jobId); };
  } else {
    startPolling(jobId);
  }
}

function startPolling(jobId) {
  if (pollTimer) clearInterval(pollTimer);
  pollTimer = setInterval(async () => {
    try {
      const res   = await fetch(`/status/${jobId}`);
      const state = await res.json();
      handleState(state, jobId);
      if (state.status === "done" || state.status === "error") {
        clearInterval(pollTimer); pollTimer = null;
      }
    } catch (_) {}
  }, 1200);
}

// ── 상태 처리 ─────────────────────────────────────────────────
function handleState(state, jobId) {
  if (state.status === "running") {
    setProgress(state.progress || 0, state.message || "처리 중...");
    if (state.progress > 0) addLog(state.message);
  } else if (state.status === "done") {
    setProgress(100, "분석 완료!");
    addLog("완료! 결과를 불러오는 중...");
    setTimeout(() => showResult(jobId), 800);
  } else if (state.status === "error") {
    showError(state.message || "알 수 없는 오류가 발생했습니다.");
  }
}

// ── 진행 바 ──────────────────────────────────────────────────
function setProgress(pct, msg) {
  const fill  = document.getElementById("progress-fill");
  const pctEl = document.getElementById("progress-pct");
  const msgEl = document.getElementById("progress-message");
  const title = document.getElementById("progress-title");
  if (fill)  fill.style.width   = pct + "%";
  if (pctEl) pctEl.textContent  = pct + "%";
  if (msgEl) msgEl.textContent  = msg;
  if (title) title.textContent  = pct >= 100 ? "완료!" : "분석 중...";
}

// ── 로그 ────────────────────────────────────────────────────
const _loggedLines = new Set();
function addLog(text) {
  if (!text || _loggedLines.has(text)) return;
  _loggedLines.add(text);
  const area = document.getElementById("log-area");
  if (!area) return;
  const line = document.createElement("div");
  line.className = "log-line";
  line.innerHTML = `<span class="log-dot">›</span><span>${escHtml(text)}</span>`;
  area.appendChild(line);
  area.scrollTop = area.scrollHeight;
}

// ── 결과 표시 ────────────────────────────────────────────────
async function showResult(jobId) {
  document.getElementById("dl-excel").href = `/download/${jobId}/excel`;
  document.getElementById("dl-pdf").href   = `/download/${jobId}/pdf`;

  try {
    const res   = await fetch(`/status/${jobId}`);
    const state = await res.json();
    renderStats(state.summary_data || null);
  } catch (_) {
    renderStats(null);
  }

  showView("result");
}

// ── 상태 통계 카드 ────────────────────────────────────────────
const STAT_DEFS = {
  "PI vs 박스내용": [
    { key: "정상",       cls: "badge-ok",     icon: "check-circle-2" },
    { key: "불일치",     cls: "badge-err",    icon: "x-circle" },
    { key: "부분선적",   cls: "badge-part",   icon: "split" },
    { key: "PI항목누락", cls: "badge-err",    icon: "file-x-2" },
    { key: "PI누락",     cls: "badge-err",    icon: "file-minus-2" },
    { key: "OCR필요",    cls: "badge-ocr",    icon: "scan-eye" },
    { key: "확인필요",   cls: "badge-warn",   icon: "alert-triangle" },
  ],
  "PI vs CI": [
    { key: "정상",       cls: "badge-ok",     icon: "check-circle-2" },
    { key: "불일치",     cls: "badge-err",    icon: "x-circle" },
    { key: "부분선적",   cls: "badge-part",   icon: "split" },
    { key: "PI미매칭",   cls: "badge-warn",   icon: "link-2-off" },
    { key: "OCR필요",    cls: "badge-ocr",    icon: "scan-eye" },
    { key: "모델불명",   cls: "badge-purple", icon: "hash" },
  ],
  "PI vs PL": [
    { key: "정상",       cls: "badge-ok",     icon: "check-circle-2" },
    { key: "불일치",     cls: "badge-err",    icon: "x-circle" },
    { key: "부분선적",   cls: "badge-part",   icon: "split" },
    { key: "PI미매칭",   cls: "badge-warn",   icon: "link-2-off" },
    { key: "OCR필요",    cls: "badge-ocr",    icon: "scan-eye" },
    { key: "확인필요",   cls: "badge-warn",   icon: "alert-triangle" },
  ],
};

function renderStats(summaryData) {
  const grid = document.getElementById("stats-grid");
  if (!grid) return;
  grid.innerHTML = "";

  Object.entries(STAT_DEFS).forEach(([section, items]) => {
    const card = document.createElement("div");
    card.className = "stat-card";

    const secEl = document.createElement("div");
    secEl.className = "stat-section";
    secEl.textContent = section;
    card.appendChild(secEl);

    items.forEach(({ key, cls, icon }) => {
      const val = summaryData ? (summaryData[section]?.[key] ?? 0) : "—";
      if (val === 0 || val === "0") return;

      const item = document.createElement("div");
      item.className = "stat-item";
      item.innerHTML = `
        <div class="stat-label">
          <i data-lucide="${icon}" style="width:14px;height:14px;flex-shrink:0"></i>
          <span>${key}</span>
        </div>
        <span class="stat-badge ${cls}">${val}</span>
      `;
      card.appendChild(item);
    });

    grid.appendChild(card);
  });

  if (window.lucide) lucide.createIcons();
}

// ── 에러 표시 ────────────────────────────────────────────────
function showError(msg, title = "오류가 발생했습니다") {
  const view = document.getElementById("progress-view");
  if (!view) return;
  view.querySelectorAll(".error-card").forEach(e => e.remove());
  const el = document.createElement("div");
  el.className = "error-card";
  el.innerHTML = `
    <i data-lucide="circle-x"></i>
    <div>
      <div class="error-title">${escHtml(title)}</div>
      <div class="error-msg">${escHtml(msg)}</div>
    </div>
  `;
  view.insertBefore(el, view.firstChild);
  showView("progress");
  if (window.lucide) lucide.createIcons();
}

// ── 리셋 ────────────────────────────────────────────────────
function resetAll() {
  if (pollTimer) { clearInterval(pollTimer); pollTimer = null; }
  currentJobId = null;
  clearSelection();
  _loggedLines.clear();

  const logArea = document.getElementById("log-area");
  if (logArea) logArea.innerHTML = "";
  setProgress(0, "대기 중...");

  const statsGrid = document.getElementById("stats-grid");
  if (statsGrid) statsGrid.innerHTML = "";

  document.querySelectorAll(".error-card").forEach(e => e.remove());
  showView("upload");
}

// ── 유틸 ────────────────────────────────────────────────────
function escHtml(str) {
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}
