// Dedup highlighter — run manually before sending
// Scans specified tabs across multiple spreadsheets, highlights + marks duplicates:
//   LIGHT RED    = email already exists as SENT anywhere across all spreadsheets
//                 -> marks any non-SENT rows as SKIPPED with reason
//   LIGHT PURPLE = email appears 2+ times as PENDING across all spreadsheets
//                 -> keeps the earliest send_at, marks the rest as SKIPPED with reason
// Red takes priority over purple.

// Config: add each spreadsheet ID and the tab names to scan within it.
// Get the spreadsheet ID from the URL: docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit
// Your friends must share edit access to their spreadsheets with your Google account.
const SPREADSHEETS = [
  { id: "YOUR_SPREADSHEET_ID",         tabs: ["Emails123"] },
  { id: "FRIEND_1_SPREADSHEET_ID",     tabs: ["Emails123"] },
  { id: "FRIEND_2_SPREADSHEET_ID",     tabs: ["Emails123"] },
];

const LIGHT_RED = "#ffe6e6";
const LIGHT_PURPLE = "#f1e8ff";

function runDedupHighlight() {
  // Step 1: collect all rows across all spreadsheets and tabs
  // each entry: { sheet, rowIndex (1-based), email, status, sendAt, id, tabName }
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
      const skipReasonCol = headers.indexOf("skip_reason");
      const sendAtCol = headers.indexOf("send_at");

      if (emailCol === -1) throw new Error(`Missing "email" column in: ${name} (${id})`);
      if (statusCol === -1) throw new Error(`Missing "status" column in: ${name} (${id})`);
      if (skipReasonCol === -1) throw new Error(`Missing "skip_reason" column in: ${name} (${id})`);
      if (sendAtCol === -1) throw new Error(`Missing "send_at" column in: ${name} (${id})`);

      for (let r = 1; r < values.length; r++) {
        const email = String(values[r][emailCol] || "").trim().toLowerCase();
        const status = String(values[r][statusCol] || "").trim().toUpperCase();
        if (!email) continue;

        const sendAtRaw = values[r][sendAtCol];
        const sendAt =
          Object.prototype.toString.call(sendAtRaw) === "[object Date]" && !isNaN(sendAtRaw)
            ? sendAtRaw
            : new Date(sendAtRaw);
        const sendAtOk = sendAtRaw && Object.prototype.toString.call(sendAt) === "[object Date]" && !isNaN(sendAt);

        allRows.push({
          sheet,
          rowIndex: r + 1,
          email,
          status,
          sendAt: sendAtOk ? sendAt : null,
          id,
          tabName: name,
          emailCol,
          statusCol,
          skipReasonCol,
          sendAtCol,
        });
      }
    }
  }

  // Step 2: build sets
  const sentEmails = new Set(
    allRows.filter((r) => r.status === "SENT").map((r) => r.email)
  );

  // Group PENDING rows by email across all spreadsheets
  const pendingByEmail = new Map();
  for (const r of allRows) {
    if (r.status === "PENDING") {
      const list = pendingByEmail.get(r.email) || [];
      list.push(r);
      pendingByEmail.set(r.email, list);
    }
  }
  const pendingDupes = new Set(
    Array.from(pendingByEmail.entries())
      .filter(([, rows]) => rows.length > 1)
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

  // Step 4a: mark duplicates of SENT as SKIPPED (red takes priority)
  for (const r of allRows) {
    if (!sentEmails.has(r.email)) continue;

    const numCols = r.sheet.getLastColumn();
    r.sheet.getRange(r.rowIndex, 1, 1, numCols).setBackground(LIGHT_RED);

    // If it's already SENT, leave it alone. Otherwise, skip it.
    if (r.status !== "SENT") {
      r.sheet.getRange(r.rowIndex, r.statusCol + 1).setValue("SKIPPED");
      r.sheet.getRange(r.rowIndex, r.skipReasonCol + 1).setValue("DUPLICATE_OF_SENT");
    }
  }

  // Step 4b: for PENDING duplicates (only those not already red), keep earliest send_at, skip rest
  for (const [email, rows] of pendingByEmail.entries()) {
    if (rows.length <= 1) continue;
    if (sentEmails.has(email)) continue; // red already handled

    // Keeper = earliest send_at; null send_at treated as far-future
    const sorted = rows.slice().sort((a, b) => {
      const aTime = a.sendAt ? a.sendAt.getTime() : Number.POSITIVE_INFINITY;
      const bTime = b.sendAt ? b.sendAt.getTime() : Number.POSITIVE_INFINITY;
      return aTime - bTime;
    });
    const keeper = sorted[0];
    const keeperRef = `${keeper.id}:${keeper.tabName}:row${keeper.rowIndex}`;

    for (const r of sorted) {
      const numCols = r.sheet.getLastColumn();
      r.sheet.getRange(r.rowIndex, 1, 1, numCols).setBackground(LIGHT_PURPLE);

      if (r === keeper) continue;
      r.sheet.getRange(r.rowIndex, r.statusCol + 1).setValue("SKIPPED");
      r.sheet.getRange(r.rowIndex, r.skipReasonCol + 1).setValue(`DUPLICATE_PENDING_KEEP:${keeperRef}`);
    }
  }

  Logger.log("Dedup highlight complete.");
}
