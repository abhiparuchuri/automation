// Dedup highlighter — run manually before sending
// Scans specified tabs across multiple spreadsheets, highlights duplicate emails:
//   RED    = email already exists as SENT anywhere across all spreadsheets
//   PURPLE = email appears 2+ times as PENDING across all spreadsheets (all occurrences)
// Red takes priority over purple. No data is changed.

// Config: add each spreadsheet ID and the tab names to scan within it.
// Get the spreadsheet ID from the URL: docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit
// Your friends must share edit access to their spreadsheets with your Google account.
const SPREADSHEETS = [
  { id: "YOUR_SPREADSHEET_ID",         tabs: ["Emails123"] },
  { id: "FRIEND_1_SPREADSHEET_ID",     tabs: ["Emails123"] },
  { id: "FRIEND_2_SPREADSHEET_ID",     tabs: ["Emails123"] },
];

const RED = "#f4cccc";
const PURPLE = "#d9d2e9";

function runDedupHighlight() {
  // Step 1: collect all rows across all spreadsheets and tabs
  // each entry: { sheet, rowIndex (1-based), email, status }
  const allRows = [];

  for (const { id, tabs } of SPREADSHEETS) {
    let ss;
    try {
      ss = SpreadsheetApp.openById(id);
    } catch (e) {
      Logger.log(`Could not open spreadsheet: ${id} — ${e.message}`);
      continue;
    }

    for (const name of tabs) {
      const sheet = ss.getSheetByName(name);
      if (!sheet) {
        Logger.log(`Sheet not found, skipping: ${name} in ${id}`);
        continue;
      }

      const values = sheet.getDataRange().getValues();
      if (values.length < 2) continue;

      const headers = values[0].map((h) => String(h).trim());
      const emailCol = headers.indexOf("email");
      const statusCol = headers.indexOf("status");

      if (emailCol === -1) throw new Error(`Missing "email" column in: ${name} (${id})`);
      if (statusCol === -1) throw new Error(`Missing "status" column in: ${name} (${id})`);

      for (let r = 1; r < values.length; r++) {
        const email = String(values[r][emailCol] || "").trim().toLowerCase();
        const status = String(values[r][statusCol] || "").trim().toUpperCase();
        if (!email) continue;

        allRows.push({ sheet, rowIndex: r + 1, email, status });
      }
    }
  }

  // Step 2: build sets
  const sentEmails = new Set(
    allRows.filter((r) => r.status === "SENT").map((r) => r.email)
  );

  // Count PENDING occurrences per email across all spreadsheets
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

  // Step 3: clear existing highlights on all sheets, then apply fresh
  for (const { id, tabs } of SPREADSHEETS) {
    let ss;
    try {
      ss = SpreadsheetApp.openById(id);
    } catch (e) {
      continue;
    }
    for (const name of tabs) {
      const sheet = ss.getSheetByName(name);
      if (!sheet) continue;
      const numRows = sheet.getLastRow();
      const numCols = sheet.getLastColumn();
      if (numRows < 2) continue;
      sheet.getRange(2, 1, numRows - 1, numCols).setBackground(null);
    }
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
