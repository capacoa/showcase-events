/**
 * Showcase Event JSON-LD Generator
 *
 * HOW TO USE:
 * 1. In Google Sheets, go to Extensions > Apps Script
 * 2. Paste this entire script, replacing any existing content
 * 3. Fill in your GitHub Personal Access Token in the GITHUB constant below
 * 4. Save (Ctrl+S), then close the editor
 * 5. Back in your sheet, a new menu "JSON-LD Tools" will appear in the menu bar
 * 6. Click JSON-LD Tools > Generate JSON-LD to preview in the sidebar
 * 7. Click JSON-LD Tools > Push to GitHub to add/update the event in the
 *    central JSON array at capacoa/showcase-events
 */

// ─── GitHub configuration ────────────────────────────────────────────────────
// Paste your Personal Access Token below (fine-grained, Contents: read+write
// on capacoa/showcase-events). Keep this file private — do not share it.

const GITHUB = {
  token:  "YOUR_PERSONAL_ACCESS_TOKEN_HERE",   // ← paste your token here
  owner:  "capacoa",
  repo:   "showcase-events",
  path:   "output/showcase-events.json",
  branch: "main",
};

// ─── Constants: fixed values that never change ──────────────────────────────

const FIXED = {
  context:                 "http://schema.org/",
  type:                    "ConferenceEvent",
  additionalType:          "http://kg.artsdata.ca/resource/Convention",
  about:                   "http://www.wikidata.org/entity/Q184485",
  eventStatus:             "https://schema.org/EventScheduled",
  eventAttendanceMode:     "https://schema.org/OfflineEventAttendanceMode",
  addressCountry:          "CA",
  performerType:           "Thing",
  performerAdditionalType: "http://dbpedia.org/ontology/Agent",
  eventIdBase:             "https://capacoa.ca/data/showcase-events/",
};

// ─── Cell addresses for each input field in the "Form" tab ──────────────────

const CELLS = {
  eventName:      "B2",
  eventUrl:       "B3",
  startDate:      "B4",
  endDate:        "B5",

  venueName:      "B8",
  streetAddress:  "B9",
  city:           "B10",
  province:       "B11",
  postalCode:     "B12",
  venueArtsdataId:"B13",

  artistFirstRow:  19,
  artistNameCol:   2,
  artistSameAsCol: 3,
  artistStopLabel: "Add lines above this",
};

// ─── Helpers ─────────────────────────────────────────────────────────────────

function cellValue(sheet, address) {
  const v = sheet.getRange(address).getValue();
  return v ? String(v).trim() : "";
}

function formatDate(value) {
  if (!value) return "";
  if (value instanceof Date) {
    const y = value.getFullYear();
    const m = String(value.getMonth() + 1).padStart(2, "0");
    const d = String(value.getDate()).padStart(2, "0");
    return `${y}-${m}-${d}`;
  }
  return String(value).trim();
}

/**
 * Converts a string into a URL-safe slug.
 * Handles accented characters (é→e, ê→e, etc.) via NFD normalization.
 * e.g. "Contact East 2026" → "contact-east-2026"
 */
function slugify(text) {
  return text
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "");
}

/**
 * Generates a persistent @id URI for the event.
 * Appends the start year as a suffix only if the year does not already
 * appear anywhere in the event name.
 * e.g. "Contact East 2026", "2026-04-08" → ".../contact-east-2026"
 *      "Contact East",      "2026-04-08" → ".../contact-east-2026"
 */
function generateEventId(eventName, startDate) {
  const year = startDate ? startDate.substring(0, 4) : "";
  const slug = slugify(eventName);
  const needsYear = year && !eventName.includes(year);
  const fullSlug = needsYear ? `${slug}-${year}` : slug;
  return FIXED.eventIdBase + fullSlug;
}

// ─── Build JSON-LD object ─────────────────────────────────────────────────────

function buildJsonLD() {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const formSheet = ss.getSheetByName("Form");

  if (!formSheet) throw new Error('Sheet "Form" not found.');

  const eventName = cellValue(formSheet, CELLS.eventName);
  const eventUrl  = cellValue(formSheet, CELLS.eventUrl);
  const startDate = formatDate(formSheet.getRange(CELLS.startDate).getValue());
  const endDate   = formatDate(formSheet.getRange(CELLS.endDate).getValue());

  const eventId = generateEventId(eventName, startDate);

  const venueName       = cellValue(formSheet, CELLS.venueName);
  const streetAddress   = cellValue(formSheet, CELLS.streetAddress);
  const city            = cellValue(formSheet, CELLS.city);
  const province        = cellValue(formSheet, CELLS.province);
  const postalCode      = cellValue(formSheet, CELLS.postalCode);
  const venueArtsdataId = cellValue(formSheet, CELLS.venueArtsdataId);

  const venueSameAs = venueArtsdataId
    ? (venueArtsdataId.startsWith("http")
        ? venueArtsdataId
        : `http://kg.artsdata.ca/resource/${venueArtsdataId}`)
    : undefined;

  const performers = [];
  let row = CELLS.artistFirstRow;
  while (true) {
    const label  = String(formSheet.getRange(row, 1).getValue() || "").trim();
    const name   = String(formSheet.getRange(row, CELLS.artistNameCol).getValue() || "").trim();
    const sameAs = String(formSheet.getRange(row, CELLS.artistSameAsCol).getValue() || "").trim();

    if (label === CELLS.artistStopLabel) break;
    if (row > 500) break;

    if (name) {
      const sameAsUrl = sameAs
        ? (sameAs.startsWith("http")
            ? sameAs
            : `http://kg.artsdata.ca/resource/${sameAs}`)
        : undefined;

      const performer = {
        "@type":          FIXED.performerType,
        "additionalType": FIXED.performerAdditionalType,
        "name":           name,
      };
      if (sameAsUrl) performer["sameAs"] = sameAsUrl;
      performers.push(performer);
    }
    row++;
  }

  const jsonLd = {
    "@id":                eventId,
    "@context":           FIXED.context,
    "@type":              FIXED.type,
    "additionalType":     FIXED.additionalType,
    "about":              FIXED.about,
    "name":               eventName,
    "startDate":          startDate,
    "location": {
      "@type":   "Place",
      ...(venueSameAs ? { "sameAs": venueSameAs } : {}),
      "name":    venueName,
      "address": {
        "@type":           "PostalAddress",
        "addressLocality": city,
        "addressRegion":   province,
        "postalCode":      postalCode,
        "streetAddress":   streetAddress,
        "addressCountry":  FIXED.addressCountry,
      },
    },
    "endDate":             endDate,
    "eventStatus":         FIXED.eventStatus,
    "eventAttendanceMode": FIXED.eventAttendanceMode,
    "url":                 eventUrl,
  };

  if (performers.length > 0) jsonLd["performer"] = performers;

  return jsonLd;
}

// ─── Sidebar preview ──────────────────────────────────────────────────────────

function generateJsonLD() {
  let output;
  try {
    const jsonLd = buildJsonLD();
    output = '<script type="application/ld+json">\n'
           + JSON.stringify(jsonLd, null, 3)
           + '\n<\/script>';
  } catch (e) {
    output = "Error: " + e.message;
  }

  const escaped = output
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");

  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
      <head>
        <style>
          body { font-family: Arial, sans-serif; font-size: 12px; margin: 10px; }
          textarea {
            width: 100%;
            height: calc(100vh - 80px);
            font-family: "Courier New", monospace;
            font-size: 11px;
            border: 1px solid #ccc;
            padding: 8px;
            box-sizing: border-box;
            resize: none;
            white-space: pre;
          }
          button { margin-bottom: 8px; padding: 6px 12px; font-size: 12px; cursor: pointer; }
        </style>
      </head>
      <body>
        <button onclick="selectAll()">Select all</button>
        <textarea id="output" readonly>${escaped}</textarea>
        <script>
          function selectAll() {
            const ta = document.getElementById("output");
            ta.select();
            ta.setSelectionRange(0, 99999);
          }
          window.onload = selectAll;
        <\/script>
      </body>
    </html>
  `)
  .setTitle("JSON-LD Output")
  .setWidth(400);

  SpreadsheetApp.getUi().showSidebar(html);
}

// ─── GitHub push ──────────────────────────────────────────────────────────────

/**
 * Fetches the current JSON array from GitHub, upserts the current event
 * (matched by @id), and commits the updated file back.
 */
function pushToGitHub() {
  const ui = SpreadsheetApp.getUi();

  if (GITHUB.token === "YOUR_PERSONAL_ACCESS_TOKEN_HERE") {
    ui.alert("Please paste your GitHub Personal Access Token into the GITHUB.token constant in the script before pushing.");
    return;
  }

  let jsonLd;
  try {
    jsonLd = buildJsonLD();
  } catch (e) {
    ui.alert("Error building JSON-LD: " + e.message);
    return;
  }

  const apiUrl = `https://api.github.com/repos/${GITHUB.owner}/${GITHUB.repo}/contents/${GITHUB.path}`;
  const headers = {
    "Authorization": `Bearer ${GITHUB.token}`,
    "Accept":        "application/vnd.github+json",
    "X-GitHub-Api-Version": "2022-11-28",
  };

  // ── Fetch current file ─────────────────────────────────────────────────────
  let currentArray = [];
  let fileSha;
  try {
    const getResponse = UrlFetchApp.fetch(apiUrl, { headers, muteHttpExceptions: true });
    const getStatus   = getResponse.getResponseCode();

    if (getStatus === 200) {
      const fileData  = JSON.parse(getResponse.getContentText());
      fileSha         = fileData.sha;
      const decoded   = Utilities.newBlob(Utilities.base64Decode(fileData.content)).getDataAsString();
      currentArray    = JSON.parse(decoded);
      if (!Array.isArray(currentArray)) currentArray = [];
    } else if (getStatus === 404) {
      // File doesn't exist yet — start with an empty array
      currentArray = [];
    } else {
      ui.alert(`GitHub error fetching file (HTTP ${getStatus}): ${getResponse.getContentText()}`);
      return;
    }
  } catch (e) {
    ui.alert("Error fetching file from GitHub: " + e.message);
    return;
  }

  // ── Upsert: replace existing entry by @id, or append ──────────────────────
  const eventId  = jsonLd["@id"];
  const existing = currentArray.findIndex(e => e["@id"] === eventId);
  if (existing >= 0) {
    currentArray[existing] = jsonLd;   // update in place
  } else {
    currentArray.push(jsonLd);         // new event
  }

  // ── Commit updated file back to GitHub ────────────────────────────────────
  const updatedContent = JSON.stringify(currentArray, null, 3);
  const encoded        = Utilities.base64Encode(Utilities.newBlob(updatedContent).getBytes());
  const commitMessage  = existing >= 0
    ? `Update event: ${jsonLd.name}`
    : `Add event: ${jsonLd.name}`;

  const payload = {
    message: commitMessage,
    content: encoded,
    branch:  GITHUB.branch,
    ...(fileSha ? { sha: fileSha } : {}),
  };

  try {
    const putResponse = UrlFetchApp.fetch(apiUrl, {
      method:  "PUT",
      headers: { ...headers, "Content-Type": "application/json" },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    });
    const putStatus = putResponse.getResponseCode();
    if (putStatus === 200 || putStatus === 201) {
      const action = existing >= 0 ? "updated" : "added";
      ui.alert(`✓ Event "${jsonLd.name}" ${action} successfully in ${GITHUB.path}.`);
    } else {
      ui.alert(`GitHub error committing file (HTTP ${putStatus}): ${putResponse.getContentText()}`);
    }
  } catch (e) {
    ui.alert("Error pushing to GitHub: " + e.message);
  }
}

// ─── Custom menu ──────────────────────────────────────────────────────────────

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("JSON-LD Tools")
    .addItem("Generate JSON-LD", "generateJsonLD")
    .addSeparator()
    .addItem("Push to GitHub", "pushToGitHub")
    .addToUi();
}
