/* global Office */

const SETTINGS_KEY = "highlightWords";

let highlightWords = [];

/** ---------- Initialization ---------- **/
Office.onReady((info) => {
  if (info.host !== Office.HostType.Outlook) {
    return;
  }

  wireEvents();
  loadSettings();

  // Optional: auto-refresh on first open (safe + convenient)
  // Comment out if you want strict manual refresh only.
  // refreshHeaders();
});

function wireEvents() {
  byId("btnRefresh").addEventListener("click", refreshHeaders);
  byId("btnExpandAll").addEventListener("click", () => setAllSections(true));
  byId("btnCollapseAll").addEventListener("click", () => setAllSections(false));
  byId("btnSettings").addEventListener("click", openSettings);
  byId("btnSaveSettings").addEventListener("click", saveSettings);
  byId("btnCloseSettings").addEventListener("click", closeSettings);
}

/** ---------- Core: Read headers ---------- **/
function refreshHeaders() {
  setStatus("Reading headers…");

  const item = Office.context.mailbox.item;

  if (!item || typeof item.getAllInternetHeadersAsync !== "function") {
    renderEmpty("(no message selected)");
    setStatus("Select a message (Reading Pane on), then refresh.");
    return;
  }

  item.getAllInternetHeadersAsync((result) => {
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      renderEmpty("");
      setStatus("Failed to read headers: " + safeError(result.error));
      return;
    }

    const headers = (result.value || "").trim();

    if (!headers) {
      renderEmpty("(no headers returned)");
      setStatus("No headers returned for this message.");
      return;
    }

    setStatus("Headers loaded.");
    renderHeaders(headers);
  });
}

/** ---------- Rendering: Collapsible sections ---------- **/
function renderHeaders(headersText) {
  const output = byId("output");
  output.innerHTML = "";

  const sections = buildHeaderSections(headersText);

  for (const section of sections) {
    output.appendChild(renderSection(section));
  }
}

function renderEmpty(text) {
  const output = byId("output");
  output.innerHTML = "";
  output.textContent = text || "";
}

/**
 * Build sections from the raw header block.
 * - Correctly handles folded header lines (continuations that begin with whitespace).
 * - Routes each full header to a group based on its name.
 */
function buildHeaderSections(rawHeaders) {
  const lines = rawHeaders.split(/\r?\n/);

  const groups = {
    "Message Metadata": [],
    "Authentication": [],
    "Routing": [],
    "Other Headers": []
  };

  let current = "";

  for (const line of lines) {
    if (/^[\s\t]/.test(line)) {
      // Continuation line (header folding)
      current += "\n" + line;
      continue;
    }

    if (current) {
      routeHeader(current, groups);
    }
    current = line;
  }

  if (current) {
    routeHeader(current, groups);
  }

  // Build section objects with sane defaults:
  // - Metadata expanded by default
  // - Everything else collapsed
  return Object.entries(groups)
    .filter(([, headers]) => headers.length > 0)
    .map(([title, headers]) => {
      const content = headers.join("\n");
      return {
        title,
        content,
        expanded: title === "Message Metadata",
        lineCount: content.split(/\r?\n/).length
      };
    });
}

/**
 * Header grouping rules (intentionally pragmatic).
 * Add or refine names here based on your environment.
 */
function routeHeader(header, groups) {
  const name = header.split(":")[0].trim().toLowerCase();

  // Routing / hop tracing
  if (name === "received" || name.startsWith("x-received")) {
    groups["Routing"].push(header);
    return;
  }

  // Auth-related
  if (
    name === "authentication-results" ||
    name === "received-spf" ||
    name === "dkim-signature" ||
    name.startsWith("arc-") ||
    name === "dmarc-filter" ||
    name === "x-ms-exchange-authentication-results"
  ) {
    groups["Authentication"].push(header);
    return;
  }

  // High-value metadata
  if ([
    "from", "to", "cc", "bcc",
    "subject", "date",
    "message-id", "in-reply-to", "references",
    "reply-to", "return-path",
    "mime-version", "content-type", "content-transfer-encoding"
  ].includes(name)) {
    groups["Message Metadata"].push(header);
    return;
  }

  groups["Other Headers"].push(header);
}

function renderSection(section) {
  const container = document.createElement("div");
  container.className = "section";

  const header = document.createElement("div");
  header.className = "section-header";
  header.setAttribute("role", "button");
  header.setAttribute("tabindex", "0");
  header.setAttribute("aria-expanded", section.expanded ? "true" : "false");

  const titleWrap = document.createElement("div");
  titleWrap.className = "section-title";

  const caret = document.createElement("div");
  caret.className = "caret";
  caret.style.transform = section.expanded ? "rotate(90deg)" : "rotate(0deg)";
  caret.setAttribute("aria-hidden", "true");

  const title = document.createElement("div");
  title.textContent = section.title;

  titleWrap.appendChild(caret);
  titleWrap.appendChild(title);

  const meta = document.createElement("div");
  meta.className = "meta";
  meta.textContent = `${section.lineCount} line${section.lineCount === 1 ? "" : "s"}`;

  header.appendChild(titleWrap);
  header.appendChild(meta);

  const body = document.createElement("div");
  body.className = "section-body";
  if (section.expanded) body.classList.add("open");

  const pre = document.createElement("pre");
  // Escape first, then highlight in HTML-safe way
  const escaped = escapeHtml(section.content);
  pre.innerHTML = applyHighlighting(escaped);

  body.appendChild(pre);

  header.addEventListener("click", () => toggleSection(header, body, caret));
  header.addEventListener("keydown", (e) => {
    // Enter / Space toggles for keyboard accessibility
    if (e.key === "Enter" || e.key === " ") {
      e.preventDefault();
      toggleSection(header, body, caret);
    }
  });

  container.appendChild(header);
  container.appendChild(body);

  return container;
}

function toggleSection(headerEl, bodyEl, caretEl) {
  const isOpen = bodyEl.classList.toggle("open");
  headerEl.setAttribute("aria-expanded", isOpen ? "true" : "false");
  caretEl.style.transform = isOpen ? "rotate(90deg)" : "rotate(0deg)";
}

function setAllSections(open) {
  const headers = document.querySelectorAll(".section-header");
  const bodies = document.querySelectorAll(".section-body");
  const carets = document.querySelectorAll(".section-header .caret");

  for (let i = 0; i < bodies.length; i++) {
    const body = bodies[i];
    const header = headers[i];
    const caret = carets[i];

    if (open) body.classList.add("open");
    else body.classList.remove("open");

    if (header) header.setAttribute("aria-expanded", open ? "true" : "false");
    if (caret) caret.style.transform = open ? "rotate(90deg)" : "rotate(0deg)";
  }
}

/** ---------- Highlighting ---------- **/
function applyHighlighting(escapedText) {
  if (!highlightWords.length) return escapedText;

  // Build a single regex that matches any word/phrase, case-insensitive.
  // NOTE: This highlights inside the escaped text, so it's safe to inject <mark>.
  const pattern = highlightWords.map(escapeRegex).join("|");
  if (!pattern) return escapedText;

  const re = new RegExp(`(${pattern})`, "gi");
  return escapedText.replace(re, "<mark>$1</mark>");
}

/** ---------- Settings (RoamingSettings) ---------- **/
function loadSettings() {
  const settings = Office.context.roamingSettings;
  const stored = settings.get(SETTINGS_KEY);

  if (Array.isArray(stored)) {
    highlightWords = normalizeWords(stored);
  } else if (typeof stored === "string") {
    // Accept legacy string formats if needed
    highlightWords = normalizeWords(stored.split(/\r?\n/));
  } else {
    highlightWords = [];
  }
}

function openSettings() {
  byId("settingsPanel").classList.remove("hidden");
  byId("wordsBox").value = highlightWords.join("\n");
  byId("settingsStatus").textContent = "";
}

function closeSettings() {
  byId("settingsPanel").classList.add("hidden");
  byId("settingsStatus").textContent = "";
}

function saveSettings() {
  const raw = byId("wordsBox").value || "";
  const words = normalizeWords(raw.split(/\r?\n/));

  const settings = Office.context.roamingSettings;
  settings.set(SETTINGS_KEY, words);

  settings.saveAsync((result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      highlightWords = words;
      byId("settingsStatus").textContent = `Saved ${words.length} highlight word(s).`;
    } else {
      byId("settingsStatus").textContent = "Save failed: " + safeError(result.error);
    }
  });
}

/** ---------- Utilities ---------- **/
function byId(id) {
  return document.getElementById(id);
}

function setStatus(text) {
  byId("status").textContent = text;
}

function safeError(err) {
  if (!err) return "(unknown error)";
  if (typeof err === "string") return err;
  if (err.message) return err.message;
  try { return JSON.stringify(err); } catch { return "(unserializable error)"; }
}

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