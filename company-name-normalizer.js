// Normalizes company names in a Google Sheet by title-casing the `company` column.
// Intended to run as an Apps Script bound to the spreadsheet used for outreach.

// Adjust these if your sheet/tab/header differ.
const COMPANY_SHEET_NAME = "Emails123";
const COMPANY_COLUMN_HEADER = "company";

function normalizeCompanyNames() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(COMPANY_SHEET_NAME);
  if (!sheet) throw new Error(`Missing sheet: ${COMPANY_SHEET_NAME}`);

  const range = sheet.getDataRange();
  const values = range.getValues();
  if (values.length < 2) return;

  const headers = values[0].map((h) => String(h).trim());
  const colIndex = headers.indexOf(COMPANY_COLUMN_HEADER);
  if (colIndex === -1) {
    throw new Error(`Missing "${COMPANY_COLUMN_HEADER}" column in sheet ${COMPANY_SHEET_NAME}`);
  }

  const updated = values.slice();

  for (let r = 1; r < values.length; r++) {
    const original = String(values[r][colIndex] || "").trim();
    if (!original) continue;

    const normalized = toTitleCaseCompany_(original);
    if (normalized !== values[r][colIndex]) {
      updated[r][colIndex] = normalized;
    }
  }

  range.setValues(updated);
}

// Basic title case plus a few common company-name fixes.
function toTitleCaseCompany_(s) {
  let result = s
    .toLowerCase()
    .split(/\s+/)
    .map((word) => {
      if (!word) return word;
      return word[0].toUpperCase() + word.slice(1);
    })
    .join(" ");

  // Preserve common suffixes / acronyms
  result = result
    .replace(/\bLlc\b/gi, "LLC")
    .replace(/\bInc\b\.?/gi, "Inc.")
    .replace(/\bCo\b\.?/gi, "Co.")
    .replace(/\bAi\b/gi, "AI")
    .replace(/\bUsa\b/gi, "USA");

  return result.trim();
}

