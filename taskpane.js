/* global Office */

const SETTINGS_KEY = "highlightWords";
let highlightWords = [];

Office.onReady((info) => {
  if (info.host !== Office.HostType.Outlook) return;

  wireEvents();
  loadSettings();
  refreshHeaders();
});

/* ---------- Events ---------- */

function wireEvents() {
  document.getElementById("btnRefresh")
    .addEventListener("click", refreshHeaders);

  document.getElementById("btnSettings")
    .addEventListener("click", openSettings);

  document.getElementById("btnSaveSettings")
    .addEventListener("click", saveSettings);

  document.getElementById("btnCloseSettings")
    .addEventListener("click", closeSettings);
}

/* ---------- Core ---------- */

function refreshHeaders() {
  const item = Office.context.mailbox.item;
  const status = document.getElementById("status");

  if (!item || typeof item.getAllInternetHeadersAsync !== "function") {
    status.textContent = "No message selected.";
    renderEmpty();
    return;
  }

  status.textContent = "Reading headers...";

  item.getAllInternetHeadersAsync((result) => {
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      status.textContent = "Error: " + (result.error?.message || "Unknown error");
      renderEmpty();
      return;
    }

    const raw = (result.value || "").trim();

    if (!raw) {
      status.textContent = "No headers returned.";
      renderEmpty("(empty)");
      return;
    }

    renderHeaders(raw);
    status.textContent = "Done.";
  });
}

/* ---------- Rendering ---------- */

function renderHeaders(raw) {
  const container = document.getElementById("output");
  container.innerHTML = "";

  const headers = splitHeaders(raw);

  let rowIndex = 0;

  for (const header of headers) {
    if (!header.trim()) continue; // skip blanks

    const row = document.createElement("div");
    row.className = "header-row " + (rowIndex % 2 ? "row-b" : "row-a");

    let html = escapeHtml(header);

    // highlight words
    if (highlightWords.length > 0) {
      const regex = new RegExp(
        "(" + highlightWords.map(escapeRegex).join("|") + ")",
        "gi"
      );
      html = html.replace(regex, "<mark>$1</mark>");
    }

    // highlight header name (before first colon)
    html = html.replace(/^([^:]+:)/, '<span class="header-name">$1</span>');

    row.innerHTML = html;
    container.appendChild(row);

    rowIndex++;
  }
}

/* ---------- Header splitting (RFC-safe) ---------- */

function splitHeaders(raw) {
  const lines = raw.split(/\r?\n/);
  const headers = [];
  let current = "";

  for (const line of lines) {
    if (/^[ \t]/.test(line)) {
      // continuation line
      current += "\n" + line;
    } else {
      if (current) headers.push(current);
      current = line;
    }
  }

  if (current) headers.push(current);

  return headers;
}

/* ---------- Settings ---------- */

function loadSettings() {
  const stored = Office.context.roamingSettings.get(SETTINGS_KEY);
  highlightWords = Array.isArray(stored) ? stored : [];
}

function openSettings() {
  document.getElementById("settingsPanel").classList.remove("hidden");
  document.getElementById("wordsBox").value = highlightWords.join("\n");
}

function closeSettings() {
  document.getElementById("settingsPanel").classList.add("hidden");
}

function saveSettings() {
  const raw = document.getElementById("wordsBox").value;

  const words = raw
    .split(/\r?\n/)
    .map(w => w.trim())
    .filter(Boolean);

  // dedupe (case-insensitive)
  highlightWords = [...new Set(words)];

  Office.context.roamingSettings.set(SETTINGS_KEY, highlightWords);
  Office.context.roamingSettings.saveAsync();

  document.getElementById("settingsStatus").textContent = "Saved.";
}

/* ---------- Helpers ---------- */

function renderEmpty(text = "") {
  document.getElementById("output").textContent = text;
}

function escapeHtml(text) {
  return text
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
}

function escapeRegex(text) {
  return text.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}
