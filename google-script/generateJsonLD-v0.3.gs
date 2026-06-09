/**
 * Showcase Event JSON-LD Generator
 *
 * HOW TO USE:
 * 1. In Google Sheets, go to Extensions > Apps Script
 * 2. Paste this entire script, replacing any existing content
 * 3. Save (Ctrl+S), then close the editor
 * 4. Back in your sheet, a new menu "JSON-LD Tools" will appear in the menu bar
 * 5. Click JSON-LD Tools > Generate JSON-LD
 * 6. A sidebar will open with the full output ready to select and copy
 */

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

// ─── Build JSON-LD object ─────────────────────────────────────────────────────

function buildJsonLD() {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const formSheet = ss.getSheetByName("Form");

  if (!formSheet) throw new Error('Sheet "Form" not found.');

  const eventName = cellValue(formSheet, CELLS.eventName);
  const eventUrl  = cellValue(formSheet, CELLS.eventUrl);
  const startDate = formatDate(formSheet.getRange(CELLS.startDate).getValue());
  const endDate   = formatDate(formSheet.getRange(CELLS.endDate).getValue());

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

// ─── Sidebar ──────────────────────────────────────────────────────────────────

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

  // Escape the output string for safe injection into HTML
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
          body {
            font-family: Arial, sans-serif;
            font-size: 12px;
            margin: 10px;
          }
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
          button {
            margin-bottom: 8px;
            padding: 6px 12px;
            font-size: 12px;
            cursor: pointer;
          }
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
          // Auto-select on load for convenience
          window.onload = selectAll;
        <\/script>
      </body>
    </html>
  `)
  .setTitle("JSON-LD Output")
  .setWidth(400);

  SpreadsheetApp.getUi().showSidebar(html);
}

// ─── Custom menu ──────────────────────────────────────────────────────────────

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("JSON-LD Tools")
    .addItem("Generate JSON-LD", "generateJsonLD")
    .addToUi();
}
