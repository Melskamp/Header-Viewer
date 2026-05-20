/* global Office */

const SETTINGS_KEY = "highlightWords";

let highlightWords = [];

// -------- Initialization --------

Office.onReady((info) => {
  if (info.host !== Office.HostType.Outlook) {
    return;
  }

  wireEvents();
  loadSettings();
});

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

// -------- Core logic --------

function refreshHeaders() {
  const status = document.getElementById("status");
  const output = document.getElementById("output");

  const item = Office.context.mailbox.item;

  if (!item || typeof item.getAllInternetHeadersAsync !== "function") {
    status.textContent = "No message selected.";
    output.textContent = "";
    return;
  }

  status.textContent = "Reading headers…";

  item.getAllInternetHeadersAsync((result) => {
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      status.textContent = "Failed to read headers: " +
        (result.error && result.error.message);
      output.textContent = "";
      return;
    }

    const headers = result.value || "";

    if (!headers.trim()) {
      status.textContent = "No headers returned for this message.";
      output.textContent = "(empty)";
      return;
    }

    status.textContent = "Headers loaded.";
    renderHeaders(headers);
  });
}

// -------- Rendering --------

function renderHeaders(headersText) {
  const output = document.getElementById("output");

  let html = escapeHtml(headersText);

  if (highlightWords.length > 0) {
    const regex = new RegExp(
      "(" + highlightWords.map(escapeRegex).join("|") + ")",
      "gi"
    );
    html = html.replace(regex, "<mark>$1</mark>");
  }

  output.innerHTML = html;
}

// -------- Settings --------

function loadSettings() {
  const settings = Office.context.roamingSettings;
  const stored = settings.get(SETTINGS_KEY);

  if (Array.isArray(stored)) {
    highlightWords = normalizeWords(stored);
  } else {
    highlightWords = [];
  }
}

function openSettings() {
  document.getElementById("settingsPanel").classList.remove("hidden");
  document.getElementById("wordsBox").value = highlightWords.join("\n");
}

function closeSettings() {
  document.getElementById("settingsPanel").classList.add("hidden");
  document.getElementById("settingsStatus").textContent = "";
}

function saveSettings() {
  const raw = document.getElementById("wordsBox").value || "";
  const words = normalizeWords(raw.split(/\r?\n/));

  const settings = Office.context.roamingSettings;
  settings.set(SETTINGS_KEY, words);

  settings.saveAsync((result) => {
    const status = document.getElementById("settingsStatus");

    if (result.status === Office.AsyncResultStatus.Succeeded) {
      highlightWords = words;
      status.textContent = "Saved.";
    } else {
      status.textContent = "Save failed: " +
        (result.error && result.error.message);
    }
  });
}

// -------- Helpers --------

function normalizeWords(words) {
  const seen = new Set();
  const cleaned = [];

  for (const w of words) {
    const t = String(w).trim();
    if (!t) continue;

    const key = t.toLowerCase();
    if (seen.has(key)) continue;

    seen.add(key);
    cleaned.push(t);
  }

  return cleaned;
}

function escapeHtml(text) {
  return String(text)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll("\"", "&quot;")
    .replaceAll("'", "&#039;");
}

function escapeRegex(text) {
  return String(text).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}