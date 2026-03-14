// Dedup highlighter — run manually before sending
// Scans specified tabs, highlights duplicate emails:
//   RED    = email already exists as SENT somewhere in the spreadsheet
//   PURPLE = email appears 2+ times as PENDING (all occurrences)
// Red takes priority over purple. No data is changed.

const DEDUP_SHEETS = ["Emails123"]; // Add/remove tab names here

const RED = "#f4cccc";
const PURPLE = "#d9d2e9";

function runDedupHighlight() {
  const ss = SpreadsheetApp.getActive();

  // Step 1: collect all rows across all tabs
  // each entry: { sheet, rowIndex (1-based), email, status }
  const allRows = [];

  for (const name of DEDUP_SHEETS) {
    const sheet = ss.getSheetByName(name);
    if (!sheet) {
      Logger.log(`Sheet not found, skipping: ${name}`);
      continue;
    }

    const values = sheet.getDataRange().getValues();
    if (values.length < 2) continue;

    const headers = values[0].map((h) => String(h).trim());
    const emailCol = headers.indexOf("email");
    const statusCol = headers.indexOf("status");

    if (emailCol === -1) throw new Error(`Missing "email" column in: ${name}`);
    if (statusCol === -1) throw new Error(`Missing "status" column in: ${name}`);

    for (let r = 1; r < values.length; r++) {
      const email = String(values[r][emailCol] || "").trim().toLowerCase();
      const status = String(values[r][statusCol] || "").trim().toUpperCase();
      if (!email) continue;

      allRows.push({ sheet, rowIndex: r + 1, email, status });
    }
  }

  // Step 2: build sets
  const sentEmails = new Set(
    allRows.filter((r) => r.status === "SENT").map((r) => r.email)
  );

  // Count PENDING occurrences per email
  const pendingCount = {};
  for (const r of allRows) {
    if (r.status === "PENDING") {
      pendingCount[r.email] = (pendingCount[r.email] || 0) + 1;
    }
  }
  const pendingDupes = new Set(
    Object.entries(pendingCount)
      .filter(([, count]) => count > 1)
      .map(([email]) => email)
  );

  // Step 3: clear existing highlights on all dedup sheets, then apply
  for (const name of DEDUP_SHEETS) {
    const sheet = ss.getSheetByName(name);
    if (!sheet) continue;
    const numRows = sheet.getLastRow();
    const numCols = sheet.getLastColumn();
    if (numRows < 2) continue;
    // Clear only data rows (skip header)
    sheet.getRange(2, 1, numRows - 1, numCols).setBackground(null);
  }

  // Step 4: apply highlights (red takes priority)
  for (const r of allRows) {
    const isRed = sentEmails.has(r.email);
    const isPurple = pendingDupes.has(r.email);

    if (!isRed && !isPurple) continue;

    const color = isRed ? RED : PURPLE;
    const numCols = r.sheet.getLastColumn();
    r.sheet.getRange(r.rowIndex, 1, 1, numCols).setBackground(color);
  }

  Logger.log("Dedup highlight complete.");
}
