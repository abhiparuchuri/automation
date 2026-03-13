// Google Sheets → Gmail email scheduler script (bound to the sheet)

const SHEET_NAME = "Emails123";

// Safer default for many Gmail accounts. Increase slowly if stable.
const MAX_SEND_PER_RUN = 80;

const DEFAULT_SUBJECT =
  "Tool feedback request: Screening candidates for AI fluency";

// Retry policy for transient scrape failures
const MAX_ATTEMPTS = 3;
const RETRY_MINUTES = [10, 60, 6 * 60]; // 10m, 1h, 6h backoff

function runMorningSend() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error(`Missing sheet: ${SHEET_NAME}`);

  const range = sheet.getDataRange();
  const values = range.getValues();
  if (values.length < 2) return;

  const headers = values[0].map((h) => String(h).trim());
  const idx = indexMap_(headers);

  const now = new Date();
  let sentCount = 0;

  // We'll batch all updates and write once at end.
  const updated = values.map((row) => row.slice());

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const status = get_(row, idx, "status");

    if (status === "SENT" || status === "SKIPPED") continue;
    // If previous run died mid-send, leave it for manual review.
    if (status === "SENDING") continue;

    const sendAt = parseDate_(getRaw_(row, idx, "send_at"));
    if (!sendAt || sendAt > now) continue;

    if (sentCount >= MAX_SEND_PER_RUN) break;

    const email = get_(row, idx, "email");
    const firstName = get_(row, idx, "first_name");
    const company = get_(row, idx, "company");
    const jobLink = get_(row, idx, "job_link");

    let roleName = get_(row, idx, "role_name");
    const attempts = parseInt(get_(row, idx, "attempt_count") || "0", 10) || 0;

    // basic validation
    if (!email) {
      set_(updated, r, idx, "status", "SKIPPED");
      set_(updated, r, idx, "skip_reason", "MISSING_EMAIL");
      set_(updated, r, idx, "last_attempt_at", now);
      continue;
    }

    // Attempt scrape if needed
    if (!roleName) {
      set_(updated, r, idx, "last_attempt_at", now);

      const scrape = scrapeRoleName_(jobLink);

      // Track attempts only when we actually try to fetch
      const nextAttempts = attempts + 1;
      set_(updated, r, idx, "attempt_count", nextAttempts);

      if (scrape.roleName) {
        roleName = scrape.roleName;
        set_(updated, r, idx, "role_name", roleName);
        set_(updated, r, idx, "skip_reason", "");
      } else {
        const reason = scrape.reason || "ROLE_NOT_FOUND";
        set_(updated, r, idx, "skip_reason", reason);

        // Retry only for transient reasons
        if (isRetryableReason_(reason) && nextAttempts < MAX_ATTEMPTS) {
          // push send_at forward for retry
          const delayMin =
            RETRY_MINUTES[
              Math.min(nextAttempts - 1, RETRY_MINUTES.length - 1)
            ];
          const retryAt = new Date(now.getTime() + delayMin * 60 * 1000);
          set_(updated, r, idx, "send_at", retryAt);
          // keep as PENDING
          set_(updated, r, idx, "status", "PENDING");
        } else {
          // Permanent skip (or max retries reached)
          set_(updated, r, idx, "status", "SKIPPED");
          if (reason === "ROLE_NOT_FOUND") {
            // You can manually fill role_name then set status=PENDING.
          }
        }
        continue;
      }
    }

    // At this point we have roleName -> mark SENDING before sending
    set_(updated, r, idx, "status", "SENDING");
    set_(updated, r, idx, "last_attempt_at", now);

    // Write the sheet state so a crash doesn’t look like "never attempted".
    sheet.getRange(1, 1, updated.length, updated[0].length).setValues(updated);

    const textBody = renderBody_({ firstName, company, roleName, jobLink });
    const htmlBody = renderHtmlBody_({ firstName, company, roleName, jobLink });

    // Send email (plain text + HTML full-width)
    MailApp.sendEmail(email, DEFAULT_SUBJECT, textBody, {
      htmlBody: htmlBody,
    });

    // Mark sent
    set_(updated, r, idx, "status", "SENT");
    set_(updated, r, idx, "sent_at", new Date());
    set_(updated, r, idx, "skip_reason", "");

    sentCount++;
  }

  // Final batch write
  sheet.getRange(1, 1, updated.length, updated[0].length).setValues(updated);
}

function renderBody_({ firstName, company, roleName, jobLink }) {
  const greeting = firstName ? `Hi ${firstName},` : "Hi,";
  return [
    greeting,
    "",
    "My name is Abhi. I'm a new graduate working on a product to screen skill + AI fluency at the same time. I believe responsible/proficient AI use is crucial. I'm in the building phase and wanted to get feedback from people working in the talent space.",
    "",
    `I noticed that some of the roles at ${company}, such as the ${roleName} role, emphasize AI tools, so I thought you would have some great insight.`,
    "",
    "I'm trying out a demo run and was wondering if you would be willing to chat more and give some feedback. Thank you!",
    "",
    "Just to clarify: not selling a product, just looking for some feedback as I build.",
    "",
    "Best,",
    "Abhi",
    "",
    jobLink || "",
  ].join("\n");
}

function renderHtmlBody_({ firstName, company, roleName, jobLink }) {
  const greeting = firstName ? `Hi ${firstName},` : "Hi,";
  return `
<p>${greeting}</p>

<p>My name is Abhi. I'm a new graduate working on a product to screen skill + AI fluency at the same time. I believe responsible/proficient AI use is crucial. I'm in the building phase and wanted to get feedback from people working in the talent space.</p>

<p>I noticed that some of the roles at ${company}, such as the ${roleName} role, emphasize AI tools, so I thought you would have some great insight.</p>

<p>I'm trying out a demo run and was wondering if you would be willing to chat more and give some feedback. Thank you!</p>

<p>Just to clarify: not selling a product, just looking for some feedback as I build.</p>

<p>Best,<br>
Abhi</p>

<p>${jobLink || ""}</p>
`.trim();
}

function scrapeRoleName_(url) {
  if (!url) return { roleName: "", reason: "MISSING_URL" };

  try {
    const resp = UrlFetchApp.fetch(url, {
      followRedirects: true,
      muteHttpExceptions: true,
      headers: { "User-Agent": "Mozilla/5.0" },
    });

    const code = resp.getResponseCode();
    if (code < 200 || code >= 300)
      return { roleName: "", reason: `HTTP_${code}` };

    const html = resp.getContentText() || "";

    // Parse meta tags regardless of attribute order.
    const meta = parseMetaTags_(html);

    let title = meta["og:title"] || meta["twitter:title"] || "";

    title = title || matchTagText_(html, "h1") || matchTitle_(html);
    title = cleanTitle_(title);

    if (!title) return { roleName: "", reason: "ROLE_NOT_FOUND" };
    return { roleName: title, reason: "" };
  } catch (e) {
    return { roleName: "", reason: "FETCH_FAILED" };
  }
}

function parseMetaTags_(html) {
  const out = {};
  const metaTags = html.match(/<meta\b[^>]*>/gi) || [];
  for (const tag of metaTags) {
    const attrs = {};
    const attrMatches =
      tag.match(
        /\b([a-zA-Z_:][-a-zA-Z0-9_:.]*)\s*=\s*(".*?"|'.*?')/g,
      ) || [];
    for (const a of attrMatches) {
      const m = a.match(
        /^([a-zA-Z_:][-a-zA-Z0-9_:.]*)\s*=\s*("([\s\S]*)"|'([\s\S]*)')$/,
      );
      if (!m) continue;
      const key = m[1].toLowerCase();
      const val = (m[3] !== undefined ? m[3] : m[4]) || "";
      attrs[key] = decodeHtml_(val.trim());
    }
    const name = (attrs["property"] || attrs["name"] || "").toLowerCase();
    const content = attrs["content"] || "";
    if (name && content && !out[name]) out[name] = content;
  }
  return out;
}

function matchTagText_(html, tag) {
  const re = new RegExp(`<${tag}[^>]*>([\\s\\S]*?)<\\/${tag}>`, "i");
  const m = html.match(re);
  if (!m) return "";
  return decodeHtml_(stripTags_(m[1])).trim();
}

function matchTitle_(html) {
  const m = html.match(/<title[^>]*>([\s\S]*?)<\/title>/i);
  return m ? decodeHtml_(stripTags_(m[1])).trim() : "";
}

function cleanTitle_(s) {
  const original = (s || "").replace(/\s+/g, " ").trim();
  if (!original) return "";

  // First remove common board suffixes
  let t = original;
  t = t.replace(/\s+[-|–—]\s+Greenhouse.*$/i, "").trim();
  t = t.replace(/\s+[-|–—]\s+Ashby.*$/i, "").trim();

  // Then handle "Role @ Company"
  t = t.replace(/\s+@\s+.*$/i, "").trim();

  return t || original;
}

function isRetryableReason_(reason) {
  // Retry for fetch/temporary server limits
  if (reason === "FETCH_FAILED") return true;
  if (/^HTTP_(429|500|502|503|504)$/.test(reason)) return true;
  return false;
}

function stripTags_(s) {
  return String(s || "").replace(/<[^>]+>/g, " ");
}

function decodeHtml_(s) {
  return String(s || "")
    .replace(/&amp;/g, "&")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&#(\d+);/g, (_, n) =>
      String.fromCharCode(parseInt(n, 10)),
    );
}

function indexMap_(headers) {
  const map = {};
  headers.forEach((h, i) => (map[h] = i));

  const required = [
    "email",
    "first_name",
    "company",
    "job_link",
    "send_at",
    "role_name",
    "status",
    "skip_reason",
    "attempt_count",
    "last_attempt_at",
    "sent_at",
  ];
  required.forEach((k) => {
    if (map[k] === undefined) throw new Error(`Missing column: ${k}`);
  });

  return map;
}

function get_(row, idx, key) {
  const v = row[idx[key]];
  return v === null || v === undefined ? "" : String(v).trim();
}

function getRaw_(row, idx, key) {
  return row[idx[key]];
}

function set_(updated, r, idx, key, val) {
  updated[r][idx[key]] = val;
}

function parseDate_(v) {
  if (!v) return null;
  if (Object.prototype.toString.call(v) === "[object Date]" && !isNaN(v))
    return v;
  const d = new Date(v);
  return isNaN(d) ? null : d;
}

