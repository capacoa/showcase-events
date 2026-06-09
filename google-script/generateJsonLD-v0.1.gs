/**
 * Showcase Event JSON-LD Generator Version 1.0
 *
 * HOW TO USE:
 * 1. In Google Sheets, go to Extensions > Apps Script
 * 2. Paste this entire script, replacing any existing content
 * 3. Save (Ctrl+S), then close the editor
 * 4. Back in your sheet, a new menu "JSON-LD Tools" will appear in the menu bar
 * 5. Click JSON-LD Tools > Generate JSON-LD
 * 6. The output will appear in cell A1 of the "JSON-LD" tab
 *
 * The script also auto-runs whenever you edit the Form tab (onEdit trigger).
 * To enable auto-run: in the Apps Script editor, go to Triggers (clock icon on
 * the left) > Add Trigger > choose generateJsonLD, event type = On edit.
 */

// ─── Constants: fixed values that never change ──────────────────────────────

const FIXED = {
  context:            "http://schema.org/",
  type:               "ConferenceEvent",
  additionalType:     "http://kg.artsdata.ca/resource/Convention",
  eventStatus:        "https://schema.org/EventScheduled",
  eventAttendanceMode:"https://schema.org/OfflineEventAttendanceMode",
  addressCountry:     "CA",
};

// ─── Cell addresses for each input field in the "Form" tab ──────────────────
// Update these if you add/move rows in the form.

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

  // Artists start at row 19, columns B (name) and C (Artsdata ID / sameAs URL)
  artistFirstRow: 19,
  artistLastRow:  38,   // extend this if you add more artist rows to the form
  artistNameCol:  2,    // column B
  artistSameAsCol:3,    // column C
};

// ─── Helpers ─────────────────────────────────────────────────────────────────

/**
 * Returns a cell value as a trimmed string, or "" if empty.
 */
function cellValue(sheet, address) {
  const v = sheet.getRange(address).getValue();
  return v ? String(v).trim() : "";
}

/**
 * Formats a JavaScript Date (or date-string) as YYYY-MM-DD.
 * Returns the value as-is if it's already a string.
 */
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
 * Determines whether a sameAs URL points to a Person or an Organization.
 * Rules (extend as needed):
 *   - kg.artsdata.ca/resource/K4-xxx  → Person
 *   - kg.artsdata.ca/resource/K5-xxx  → Person
 *   - kg.artsdata.ca/resource/K12-xxx → Person
 *   - kg.artsdata.ca/resource/K2-xxx  → Person
 *   - everything else                 → Organization
 *
 * You can refine this logic or add an explicit "Type" column to the form
 * if you need more control.
 */
function inferPerformerType(sameAs) {
  const personPrefixes = ["/K4-", "/K5-", "/K12-", "/K2-"];
  if (personPrefixes.some(p => sameAs.includes(p))) return "Person";
  return "Organization";
}

// ─── Main function ────────────────────────────────────────────────────────────

function generateJsonLD() {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const formSheet = ss.getSheetByName("Form");
  const outSheet  = ss.getSheetByName("JSON-LD");

  if (!formSheet) { SpreadsheetApp.getUi().alert('Sheet "Form" not found.'); return; }
  if (!outSheet)  { SpreadsheetApp.getUi().alert('Sheet "JSON-LD" not found.'); return; }

  // ── Read event-level fields ────────────────────────────────────────────────
  const eventName = cellValue(formSheet, CELLS.eventName);
  const eventUrl  = cellValue(formSheet, CELLS.eventUrl);

  const rawStart  = formSheet.getRange(CELLS.startDate).getValue();
  const rawEnd    = formSheet.getRange(CELLS.endDate).getValue();
  const startDate = formatDate(rawStart);
  const endDate   = formatDate(rawEnd);

  // ── Read venue fields ──────────────────────────────────────────────────────
  const venueName       = cellValue(formSheet, CELLS.venueName);
  const streetAddress   = cellValue(formSheet, CELLS.streetAddress);
  const city            = cellValue(formSheet, CELLS.city);
  const province        = cellValue(formSheet, CELLS.province);
  const postalCode      = cellValue(formSheet, CELLS.postalCode);
  const venueArtsdataId = cellValue(formSheet, CELLS.venueArtsdataId);

  // Build venue @id: if the user entered a full URL use it directly,
  // otherwise treat it as a bare Artsdata resource ID.
  const venueId = venueArtsdataId
    ? (venueArtsdataId.startsWith("http")
        ? venueArtsdataId
        : `http://kg.artsdata.ca/resource/${venueArtsdataId}`)
    : undefined;

  // ── Read performers ────────────────────────────────────────────────────────
  const performers = [];
  for (let row = CELLS.artistFirstRow; row <= CELLS.artistLastRow; row++) {
    const name   = String(formSheet.getRange(row, CELLS.artistNameCol).getValue() || "").trim();
    const sameAs = String(formSheet.getRange(row, CELLS.artistSameAsCol).getValue() || "").trim();
    if (!name) continue;   // skip empty rows

    // Build the sameAs URL
    const sameAsUrl = sameAs
      ? (sameAs.startsWith("http")
          ? sameAs
          : `http://kg.artsdata.ca/resource/${sameAs}`)
      : undefined;

    const performer = {
      "@type": inferPerformerType(sameAsUrl || ""),
      "name":  name,
    };
    if (sameAsUrl) performer["sameAs"] = sameAsUrl;
    performers.push(performer);
  }

  // ── Assemble JSON-LD object ────────────────────────────────────────────────
  const jsonLd = {
    "@context":           FIXED.context,
    "@type":              FIXED.type,
    "additionalType":     FIXED.additionalType,
    "name":               eventName,
    "startDate":          startDate,
    "location": {
      "@type":   "Place",
      ...(venueId ? { "@id": venueId } : {}),
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

  // ── Write to JSON-LD tab ───────────────────────────────────────────────────
  const outputCell = outSheet.getRange("A1");
  outputCell.setValue(JSON.stringify(jsonLd, null, 3));

  // Make the cell easy to read: wrap text, wide column, monospace font
  outputCell.setWrap(true);
  outputCell.setFontFamily("Courier New");
  outputCell.setFontSize(10);
  outSheet.setColumnWidth(1, 700);

  SpreadsheetApp.getUi().alert("JSON-LD generated successfully in the JSON-LD tab!");
}

// ─── Custom menu (appears when the spreadsheet opens) ────────────────────────

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("JSON-LD Tools")
    .addItem("Generate JSON-LD", "generateJsonLD")
    .addToUi();
}
