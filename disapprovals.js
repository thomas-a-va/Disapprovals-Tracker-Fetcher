const puppeteer = require("puppeteer");
const fs = require("fs");
const path = require("path");

// --------- Load client list from clients.csv ---------
function parseCsvLine(line) {
  const cells = [];
  let current = "";
  let inQuotes = false;

  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (ch === '"') {
      if (inQuotes && line[i + 1] === '"') {
        current += '"';
        i++;
      } else {
        inQuotes = !inQuotes;
      }
    } else if (ch === "," && !inQuotes) {
      cells.push(current);
      current = "";
    } else {
      current += ch;
    }
  }
  cells.push(current);
  return cells.map(c => c.trim());
}

function loadClientsFromCsv() {
  const csvPath = path.join(__dirname, "clients.csv");
  if (!fs.existsSync(csvPath)) {
    throw new Error("clients.csv not found next to disapprovals.js");
  }

  const text = fs.readFileSync(csvPath, "utf8");
  const lines = text.split(/\r?\n/).map(l => l.trim()).filter(Boolean);
  if (lines.length < 2) {
    throw new Error("clients.csv has no data rows");
  }

  const result = [];
  for (let i = 1; i < lines.length; i++) {
    const row = parseCsvLine(lines[i]);
    if (!row.length) continue;
    const [name, idStr, enabledRaw] = row;
    if (!name || !idStr) continue;
    const id = Number(idStr);
    if (!Number.isFinite(id)) continue;

    let enabled = (enabledRaw || "true").toString().toLowerCase();
    enabled = enabled === "true";

    if (!enabled) continue;
    result.push({ name, id });
  }
  return result;
}

/* =========================
   Env / Config
   ========================= */
const EMAIL = process.env.AR_EMAIL;

let CLIENTS = loadClientsFromCsv();

const SINGLE_CLIENT = process.env.AR_CLIENT_ID ? Number(process.env.AR_CLIENT_ID) : null;
const hasSingleOverride = (SINGLE_CLIENT !== null && !Number.isNaN(SINGLE_CLIENT));

if (hasSingleOverride) {
  CLIENTS = CLIENTS.filter(c => c.id === SINGLE_CLIENT);
}

if (!CLIENTS.length) {
  console.error("No clients to run. Exiting.");
  process.exit(0);
}

let VERSION = process.env.AR_VERSION || "";
const PROFILE = process.env.AR_CHROME_PROFILE || "C:\\PPRChrome";
const CHROME  = "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe";

const ADV_PASSWORD    = process.env.AR_PASSWORD || "";
const GOOGLE_PASSWORD = process.env.AR_GOOGLE_PASSWORD || "";
const HEADLESS        = process.env.AR_HEADLESS === "1";

const ADV   = "https://advantage.advertiserreports.com";
const SUPER = "https://superusers.advertiserreports.com";

const ADV_HOME   = `${ADV}/rip/launch/home?localeCode=EN_US`;
const SUPER_HOME = `${SUPER}/rip/launch/home?localeCode=EN_US`;

const CHANGE_URL           = `${SUPER}/rip/login/json/changeClient`;
const GET_DISAPPROVALS_URL = `${ADV}/rip/admin/json/getSearchEngineDisapprovals`;
const APPS_SCRIPT_WEBHOOK = process.env.AR_APPS_SCRIPT_URL || "";
const SNAP_MODE = (process.env.AR_SNAP_MODE || "none").toLowerCase();
const OUTPUT_DIR = path.resolve(__dirname, "outputs");

/* =========================
   Summary Table helpers
   ========================= */
function padRight(str, len) {
  str = String(str ?? "");
  if (str.length >= len) return str;
  return str + " ".repeat(len - str.length);
}

function printSummaryTable(rows) {
  const hasSheets = rows.some(r => r.sheetUrl);
  const hasMoved = rows.some(r => r.moved !== undefined);
  const headers = ["Client", "Fetched", "CSV"];
  if (hasSheets) headers.push("Sheet URL");
  if (hasMoved) headers.push("Moved to Drive");

  const sheetUrlCol = (r) => r.sheetUrl || "-";

  const widths = [
    Math.max(headers[0].length, ...rows.map(r => (r.client || "").length)),
    Math.max(headers[1].length, ...rows.map(r => (r.fetched ? "Yes" : "No").length)),
    Math.max(headers[2].length, ...rows.map(r => (r.saved ? "Yes" : "No").length)),
  ];
  if (hasSheets) {
    widths.push(Math.max("Sheet URL".length, ...rows.map(r => sheetUrlCol(r).length)));
  }
  if (hasMoved) {
    widths.push(Math.max("Moved to Drive".length, ...rows.map(r => (r.moved ? "Yes" : "No").length)));
  }

  const line = (cols) =>
    cols.map((c, i) => padRight(c, widths[i])).join(" | ");

  const sep = widths.map(w => "-".repeat(w)).join("-+-");

  console.log("\n==================== SUMMARY ====================");
  console.log(line(headers));
  console.log(sep);

  for (const r of rows) {
    const cols = [
      r.client || "",
      r.fetched ? "Yes" : "No",
      r.saved ? "Yes" : "No",
    ];
    if (hasSheets) cols.push(sheetUrlCol(r));
    if (hasMoved) cols.push(r.moved ? "Yes" : "No");
    console.log(line(cols));
  }
  console.log("=================================================\n");
}

/* =========================
   Node-level HTTP POST (with cookies, no CORS)
   ========================= */
function nodePost(url, body, cookieHeader) {
  return new Promise((resolve) => {
    const https = require("https");
    const http = require("http");
    const data = JSON.stringify(body);

    const doRequest = (requestUrl, method, sendBody) => {
      const parsed = new URL(requestUrl);
      const mod = parsed.protocol === "https:" ? https : http;
      const isPost = method === "POST";
      const headers = {};
      if (cookieHeader) headers["Cookie"] = cookieHeader;
      if (isPost) {
        headers["Content-Type"] = "application/json";
        headers["Content-Length"] = Buffer.byteLength(sendBody);
      }

      const req = mod.request(requestUrl, { method, headers }, (res) => {
        if (res.statusCode >= 300 && res.statusCode < 400 && res.headers.location) {
          // 307/308 keep the original method; 301/302/303 switch to GET
          const keepPost = res.statusCode === 307 || res.statusCode === 308;
          doRequest(
            res.headers.location,
            keepPost ? "POST" : "GET",
            keepPost ? sendBody : null
          );
          return;
        }
        let chunks = [];
        res.on("data", (c) => chunks.push(c));
        res.on("end", () => {
          const text = Buffer.concat(chunks).toString();
          let json = null;
          try { json = JSON.parse(text); } catch {}
          resolve({ status: res.statusCode, text, json });
        });
      });
      req.on("error", (e) => resolve({ status: -1, text: String(e), json: null }));
      if (isPost && sendBody) req.write(sendBody);
      req.end();
    };

    doRequest(url, "POST", data);
  });
}

/* =========================
   Screenshot Debug (context-safe)
   ========================= */
const SCREEN_DIR = path.resolve(process.cwd(), "screenshots");
if (!fs.existsSync(SCREEN_DIR)) fs.mkdirSync(SCREEN_DIR, { recursive: true });
let shotSeq = 0;

async function snap(page, label) {
  try {
    if (SNAP_MODE === "none") return;
    if (!page || page.isClosed()) return;
    const num = String(++shotSeq).padStart(3, "0");
    const fname = `${num}_${String(label || "").replace(/[^\w.-]+/g, "_")}.png`;
    const full = path.join(SCREEN_DIR, fname);
    await page.screenshot({ path: full, fullPage: true }).catch(() => {});
  } catch {}
}

/* =========================
   Utilities (context-safe waits)
   ========================= */
function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

async function safeWaitForSelector(page, sel, opts = {}) {
  try {
    if (!page || page.isClosed()) return null;
    return await page.waitForSelector(sel, opts);
  } catch {
    return null;
  }
}

async function safeWaitForNavigation(page, opts = {}) {
  try {
    if (!page || page.isClosed()) return null;
    await page.waitForNavigation(opts);
  } catch {}
  return page;
}

async function ensureMCID(page) {
  const cookies = await page.cookies(SUPER);
  const names = cookies.map((c) => c.name);
  console.log("SuperUsers cookies:", names.join(", ") || "(none)");
  return cookies.some((c) => c.name === "MCID");
}

/* =========================
   postInside – fetch inside page so cookies/origin apply
   ========================= */
async function postInside(page, url, body, headers, credOpt) {
  if (!page || page.isClosed()) {
    return { status: -1, text: "Page is closed", json: null };
  }

  try {
    return await page.evaluate(
      async ({ url, body, headers, credOpt }) => {
        try {
          const r = await fetch(url, {
            method: "POST",
            headers: Object.assign(
              {
                "Content-Type": "application/json;charset=UTF-8",
                Accept: "application/json, text/plain, */*",
                "X-Requested-With": "XMLHttpRequest",
              },
              headers || {}
            ),
            body: JSON.stringify(body),
            credentials: credOpt || "include",
            referrerPolicy: "no-referrer-when-downgrade",
          });
          const text = await r.text();
          let json = null;
          try { json = JSON.parse(text); } catch {}
          return { status: r.status, text, json };
        } catch (e) {
          return { status: -1, text: String(e), json: null };
        }
      },
      { url, body, headers, credOpt }
    );
  } catch (e) {
    console.log("postInside outer error:", e?.message || String(e));
    return { status: -1, text: String(e), json: null };
  }
}

/* =========================
   Version detection (from SUPER, same as PPR)
   ========================= */
async function detectVersion(page) {
  try {
    const v1 = await page.evaluate(() => {
      const tryvals = [];
      tryvals.push(window.__APP_CONFIG__ && window.__APP_CONFIG__.version);
      tryvals.push(window.appVersion);
      tryvals.push(window.version);

      const mMeta = document.querySelector("meta[name='app:version']")?.content;
      const mData = document.querySelector("[data-version]")?.getAttribute("data-version");
      if (mMeta) tryvals.push(mMeta);
      if (mData) tryvals.push(mData);

      const re = /\b\d{4}\/\d{2}\/\d{2}\s+\d{2}:\d{2}:\d{2}\s+[A-Z]{2,5}\b/;
      for (const s of Array.from(document.scripts)) {
        const t = s.textContent || "";
        const m = t.match(re);
        if (m) { tryvals.push(m[0]); break; }
      }
      const html = document.documentElement.outerHTML;
      const m2 = html.match(re);
      if (m2) tryvals.push(m2[0]);
      return tryvals.filter(Boolean)[0] || "";
    });
    if (v1) return v1;
  } catch (e) {
    console.log("detectVersion v1 evaluate failed (navigation?):", e.message);
  }

  try {
    const v2 = await page.evaluate(async (home) => {
      try {
        const r = await fetch(home, { credentials: "include" });
        const h = await r.text();
        const re = /\b\d{4}\/\d{2}\/\d{2}\s+\d{2}:\d{2}:\d{2}\s+[A-Z]{2,5}\b/;
        const m = h.match(re);
        if (m) return m[0];
        const m2 = h.match(/"version"\s*:\s*"([^"]+)"/);
        return (m2 && m2[1]) || "";
      } catch { return ""; }
    }, `${SUPER}/`);
    if (v2) return v2;
  } catch (e) {
    console.log("detectVersion v2 evaluate failed (navigation?):", e.message);
  }

  try {
    const v3 = await page.evaluate(async (home) => {
      try {
        const r = await fetch(home, { credentials: "include" });
        const h = await r.text();
        const re = /\b\d{4}\/\d{2}\/\d{2}\s+\d{2}:\d{2}:\d{2}\s+[A-Z]{2,5}\b/;
        const m = h.match(re);
        if (m) return m[0];
        const m2 = h.match(/"version"\s*:\s*"([^"]+)"/);
        return (m2 && m2[1]) || "";
      } catch { return ""; }
    }, SUPER_HOME);
    if (v3) return v3;
  } catch (e) {
    console.log("detectVersion v3 evaluate failed (navigation?):", e.message);
  }

  return "";
}

/* =========================
   App version auto-repair helpers
   ========================= */
async function detectVersionFresh(page) {
  let v = await detectVersion(page);
  if (v) return v;

  const bust = `&_=${Date.now()}`;
  try {
    await page.goto(`${SUPER}/?nocache=1${bust}`, { waitUntil: "domcontentloaded", timeout: 60000 }).catch(()=>{});
    await page.goto(`${SUPER_HOME}&nocache=1${bust}`, { waitUntil: "domcontentloaded", timeout: 60000 }).catch(()=>{});
  } catch {}

  await sleep(2000);

  try {
    const v2 = await page.evaluate(async (home) => {
      try {
        const r = await fetch(home + (home.includes('?') ? '&' : '?') + '_=' + Date.now(), { credentials: "include", cache: "reload" });
        const h = await r.text();
        const re = /\b\d{4}\/\d{2}\/\d{2}\s+\d{2}:\d{2}:\d{2}\s+[A-Z]{2,5}\b/;
        const m = h.match(re);
        if (m) return m[0];
        const m2 = h.match(/"version"\s*:\s*"([^"]+)"/);
        return (m2 && m2[1]) || "";
      } catch { return ""; }
    }, SUPER_HOME);
    if (v2) return v2;
  } catch (e) {
    console.log("detectVersionFresh evaluate failed (navigation?):", e.message);
  }

  await sleep(3000);
  return await detectVersion(page);
}

async function forceAppRefresh(page, snapName) {
  const bust = `&_=${Date.now()}`;
  try {
    await page.goto(`${SUPER}/?refresh=1${bust}`, { waitUntil: "domcontentloaded", timeout: 60000 }).catch(()=>{});
  } catch {}
  try {
    await page.goto(`${SUPER_HOME}&refresh=1${bust}`, { waitUntil: "domcontentloaded", timeout: 60000 }).catch(()=>{});
  } catch {}
  if (snapName) { try { await snap(page, snapName); } catch {} }
}

async function changeClientRobust(page, email, clientId) {
  let version = VERSION;

  if (!version) {
    const maxTries = 5;
    const delayMs = 4000;
    for (let attempt = 1; attempt <= maxTries; attempt++) {
      version = await detectVersionFresh(page);
      console.log(`detectVersionFresh attempt ${attempt}:`, version || "(none)");
      if (version) {
        VERSION = version;
        break;
      }
      if (attempt < maxTries) {
        console.log("No version yet, waiting before retry…");
        await sleep(delayMs);
      }
    }

    if (!version) {
      console.error("changeClientRobust: could not detect app version after retries.");
      return {
        ok: false,
        version: "",
        res: {
          status: 500,
          json: { error: { message: "Could not detect app version (fresh)" } }
        }
      };
    }
  }

  console.log("changeClient(version):", version || "(none)");

  for (let attempt = 1; attempt <= 5; attempt++) {
    const res = await postInside(page, CHANGE_URL, {
      currentLoginName: email,
      version,
      data: { id: clientId },
    });
    const short = (res.text || "").slice(0, 200);
    console.log(`changeClient attempt ${attempt}:`, res.status, short);
    await snap(page, `104_change_client_${res.status}_attempt${attempt}`);

    const code = res?.json?.error?.code;
    if (res.status === 200 && !code) {
      return { ok: true, version, res };
    }

    if (res.status === -1) {
      console.log(`>> changeClient: transient error (context destroyed / navigation), waiting before retry…`);
      await sleep(3000);
      try {
        await page.goto(SUPER_HOME, { waitUntil: "networkidle2", timeout: 60000 }).catch(()=>{});
        await sleep(1000);
      } catch {}
      continue;
    }

    if (code === -32003) {
      console.log(">> changeClient: app version stale; forcing app refresh and re-detecting…");
      await forceAppRefresh(page, `refresh_attempt${attempt}`);

      let newVersion = "";
      const maxVerTries = 3;
      const verDelayMs = 4000;
      for (let vTry = 1; vTry <= maxVerTries; vTry++) {
        newVersion = await detectVersionFresh(page);
        console.log(`detectVersionFresh (stale) attempt ${vTry}:`, newVersion || "(none)");
        if (newVersion) break;
        if (vTry < maxVerTries) {
          console.log("No version yet after refresh, waiting before retry…");
          await sleep(verDelayMs);
        }
      }

      if (newVersion) {
        version = newVersion;
        VERSION = newVersion;
        console.log(">> changeClient: re-detected version:", newVersion);
        continue;
      }

      console.error("changeClientRobust: still no version after refresh retries.");
      return {
        ok: false,
        version: "",
        res: {
          status: 500,
          json: { error: { message: "Could not detect app version after refresh" } }
        }
      };
    }

    return { ok: false, version, res };
  }

  return {
    ok: false,
    version,
    res: {
      status: 200,
      json: { error: { code: -32003, message: "Version still stale after retries" } }
    }
  };
}

/* =========================
   DOM helpers
   ========================= */
async function forceType(page, selector, value) {
  try {
    const el = await page.$(selector);
    if (!el) return false;
    await el.click({ clickCount: 3 }).catch(() => {});
    await page.keyboard.press("Backspace").catch(() => {});
    await page.keyboard.press("Delete").catch(() => {});
    await el.type(value, { delay: 30 }).catch(() => {});
    return true;
  } catch {
    return false;
  }
}

/* =========================
   Navigation handoff helper
   ========================= */
async function waitForNextPageOrNav(currentPage, browser, {
  sameTabWaitUntil = "networkidle2",
  timeout = 30000,
  urlMustMatch = null,
} = {}) {
  if (!currentPage || currentPage.isClosed()) {
    const pages = await browser.pages();
    return pages[0] || null;
  }
  const startUrl = currentPage.url();
  function urlMatches(u) { return urlMustMatch ? urlMustMatch.test(u) : true; }

  const sameTabNav = currentPage
    .waitForNavigation({ waitUntil: sameTabWaitUntil, timeout })
    .then(() => currentPage)
    .catch(() => null);

  const newTarget = browser
    .waitForTarget(t => {
      const u = t.url();
      if (!u || u === "about:blank") return false;
      if (u === startUrl) return false;
      return urlMatches(u);
    }, { timeout })
    .then(t => t.page().catch(() => null))
    .catch(() => null);

  const closedThenFallback = (async () => {
    try { await currentPage.waitForEvent?.("close", { timeout }); } catch { return null; }
    const pages = await browser.pages();
    for (const p of pages) { try { if (urlMatches(p.url())) return p; } catch {} }
    return pages[0] || null;
  })();

  const winner = await Promise.race([sameTabNav, newTarget, closedThenFallback]);
  if (!winner) {
    const pages = await browser.pages();
    for (const p of pages) { const u = p.url(); if (urlMatches(u)) return p; }
    return currentPage;
  }
  return winner;
}

/* =========================
   Dialog auto-accept + modal OK sweeper
   ========================= */
async function wireDialogAutoAccept(browser) {
  const hook = async (p) => {
    try {
      p.on('dialog', async (d) => { try { await d.accept(); } catch {} });
    } catch {}
  };
  for (const p of await browser.pages()) await hook(p);
  browser.on('targetcreated', async (t) => {
    const p = await t.page().catch(() => null);
    if (p) await hook(p);
  });
}

async function clickModalOk(page, {
  labels = ["OK","Ok","Okay","Confirm","Continue","Proceed","Yes"],
  timeoutMs = 8000
} = {}) {
  if (!page || page.isClosed()) return false;
  const t0 = Date.now();

  async function tryClickInFrame(frame) {
    return frame.evaluate((labels) => {
      const norm = (s) => (s || "").replace(/\s+/g, " ").trim().toLowerCase();
      const wanted = labels.map(l => l.toLowerCase());
      const candidates = Array.from(document.querySelectorAll("button, [role='button'], input[type='button'], input[type='submit']"));
      for (const el of candidates) {
        const txt = norm(el.innerText || el.value || el.textContent || el.getAttribute('aria-label'));
        if (!txt) continue;
        if (wanted.some(w => txt === w || txt.includes(w))) {
          const rect = el.getBoundingClientRect();
          const style = window.getComputedStyle(el);
          const visible = rect.width > 1 && rect.height > 1 && style.visibility !== "hidden" && style.display !== "none";
          if (visible) { (el instanceof HTMLElement) && el.click(); return true; }
        }
      }
      const common = document.querySelector(
        ".modal-footer .btn-primary, .modal .btn-primary, .swal2-confirm, .mdc-button--raised, .btn.btn-primary"
      );
      if (common) { (common instanceof HTMLElement) && common.click(); return true; }
      return false;
    }, labels).catch(() => false);
  }

  while (Date.now() - t0 < timeoutMs) {
    if (await tryClickInFrame(page.mainFrame())) return true;
    for (const f of page.frames()) {
      if (f === page.mainFrame()) continue;
      if (await tryClickInFrame(f)) return true;
    }
    await new Promise(r => setTimeout(r, 250));
  }
  return false;
}

/* =========================
   Login flows
   ========================= */
async function handleAdvantageCredentials(page) {
  console.log(">> Advantage login form detected — using autofill");
  await sleep(3000);

  const clickTargets = [
    "button[type='submit']",
    "input[type='submit']",
    "button#submit",
    "button:has-text('Sign In')",
    "button:has-text('Login')",
    "input[value='Sign In']"
  ];

  let clicked = false;
  for (const sel of clickTargets) {
    try {
      const el = await page.$(sel);
      if (el) { await el.click(); clicked = true; break; }
    } catch {}
  }
  if (!clicked) {
    try { await page.keyboard.press("Enter"); } catch {}
  }
  try {
    await page.waitForLoadState("domcontentloaded", { timeout: 20000 });
  } catch {}
  console.log(">> Advantage login submitted (autofill mode)");
}

async function handleGoogleSSO(page, browser) {
  await snap(page, "010_before_google_sso");

  let clickedBtn = false;
  const btn = await safeWaitForSelector(page, "button.gsi-material-button", {
    visible: true,
    timeout: 12000,
  });
  if (btn) {
    await snap(page, "011_google_btn_found");
    try {
      await page.$eval("button.gsi-material-button", (b) => {
        const evOpts = { bubbles: true, cancelable: true, composed: true };
        b.dispatchEvent(new MouseEvent("pointerdown", evOpts));
        b.dispatchEvent(new MouseEvent("mousedown", evOpts));
        b.dispatchEvent(new MouseEvent("mouseup", evOpts));
        b.dispatchEvent(new MouseEvent("click", evOpts));
      });
      clickedBtn = true;
    } catch {}
  }
  if (!clickedBtn) {
    await snap(page, "011_google_btn_not_found_try_text");
    const viaText = await page
      .evaluate(() => {
        const norm = (s) =>
          (s || "").replace(/\s+/g, " ").trim().toLowerCase();
        const spans = Array.from(
          document.querySelectorAll(
            "span.gsi-material-button-contents, button, [role='button']"
          )
        );
        const el = spans.find((n) =>
          norm(n.textContent).includes("sign in with google")
        );
        const btn = el ? el.closest("button") || el : null;
        if (btn) {
          const evOpts = { bubbles: true, cancelable: true, composed: true };
          btn.dispatchEvent(new MouseEvent("pointerdown", evOpts));
          btn.dispatchEvent(new MouseEvent("mousedown", evOpts));
          btn.dispatchEvent(new MouseEvent("mouseup", evOpts));
          btn.dispatchEvent(new MouseEvent("click", evOpts));
          return true;
        }
        return false;
      })
      .catch(() => false);
    await snap(
      page,
      viaText ? "012_google_btn_clicked_via_text" : "012_google_btn_click_failed"
    );
  }

  await sleep(600);
  let googlePage = null;
  try {
    const target = await browser.waitForTarget(
      (t) => /accounts\.google\.com/.test(t.url()),
      { timeout: 15000 }
    );
    googlePage = await target.page();
  } catch {
    if (/accounts\.google\.com/.test(page.url())) googlePage = page;
  }
  if (!googlePage || googlePage.isClosed()) {
    await snap(page, "013_no_google_accounts_detected");
    console.log(">> Google SSO: accounts.google.com did not appear after click.");
    return;
  }
  await snap(googlePage, "014_google_accounts_page");

  let chooserShown = false;
  try {
    const chooser = await safeWaitForSelector(
      googlePage,
      "div[data-identifier], div[data-email], div[role='button'][data-identifier]",
      { timeout: 8000 }
    );
    chooserShown = !!chooser;
  } catch {
    chooserShown = false;
  }

  if (chooserShown) {
    await snap(googlePage, "015_google_chooser_ready");
    const matched = await googlePage
      .$$eval(
        "div[data-identifier], div[data-email], div[role='button'][data-identifier]",
        (nodes, email) => {
          const norm = (s) => String(s || "").trim().toLowerCase();
          const target = norm(email);
          const m = nodes.find((n) => {
            const de =
              n.getAttribute("data-email") ||
              n.getAttribute("data-identifier") ||
              n.textContent;
            return (
              norm(de) === target ||
              norm(n.textContent || "").includes(target)
            );
          });
          if (m) {
            m.click();
            return true;
          }
          return false;
        },
        EMAIL
      )
      .catch(() => false);

    if (matched) {
      const nextPage = await waitForNextPageOrNav(googlePage, browser, {
        sameTabWaitUntil: "domcontentloaded",
        timeout: 40000,
        urlMustMatch:
          /advertiserreports\.com|vendasta|superusers|advantage|roar|accounts\.google\.com/i,
      });
      if (nextPage && !nextPage.isClosed()) {
        await snap(nextPage, "016_after_tile_nav");
        if (!/accounts\.google\.com/i.test(nextPage.url())) return;
        googlePage = nextPage;
      } else {
        return;
      }
    }
  }

  console.log(">> Google SSO: using email/password fallback.");
  if (!googlePage || googlePage.isClosed()) return;

  try {
    if (!/accounts\.google\.com/i.test(googlePage.url())) return;

    const emailField = await safeWaitForSelector(
      googlePage,
      "input[type='email'], input#identifierId",
      { timeout: 10000 }
    );
    if (emailField) {
      const emailSel = (await googlePage
        .$(`input[type='email']`)
        .catch(() => null))
        ? "input[type='email']"
        : "input#identifierId";
      await forceType(googlePage, emailSel, EMAIL);
      await snap(googlePage, "017_email_entered");
      await Promise.race([
        googlePage.click("#identifierNext").catch(() => {}),
        googlePage.click("button[type='submit']").catch(() => {}),
        googlePage.click("div[role='button']#identifierNext").catch(() => {}),
      ]);
      await safeWaitForNavigation(googlePage, {
        waitUntil: "domcontentloaded",
        timeout: 15000,
      });
      if (googlePage && !googlePage.isClosed())
        await snap(googlePage, "018_after_identifier_next");
    }
  } catch {}

  try {
    if (!googlePage || googlePage.isClosed()) return;
    if (!/accounts\.google\.com/i.test(googlePage.url())) return;

    const pwField = await safeWaitForSelector(
      googlePage,
      "input[type='password']",
      { timeout: 15000 }
    );
    if (pwField) {
      if (GOOGLE_PASSWORD) {
        await forceType(googlePage, "input[type='password']", GOOGLE_PASSWORD);
      }
      await snap(googlePage, "019_password_entered");
      await Promise.race([
        googlePage.click("#passwordNext").catch(() => {}),
        googlePage.click("button[type='submit']").catch(() => {}),
        googlePage.click("div[role='button']#passwordNext").catch(() => {}),
      ]);
      await safeWaitForNavigation(googlePage, {
        waitUntil: "networkidle2",
        timeout: 20000,
      });
      if (googlePage && !googlePage.isClosed())
        await snap(googlePage, "020_after_password_next");
    }
  } catch {}
}

async function settleHandsFreeLogin(page, browser) {
  console.log(">> Hands-free login starting…");
  await snap(page, "000_landed_on_advantage");

  const start = Date.now();
  const maxMs = 60 * 1000;

  while (Date.now() - start < maxMs) {
    const url = page.url();

    if (/superusers\.advertiserreports\.com|roar\.advertiserreports\.com|advantage\.advertiserreports\.com/.test(url)) {
      const hasAppShell = await page.evaluate(() => {
        return !!(document.querySelector("app-root, app-shell, nav .user, [data-role='app-container']"));
      }).catch(() => false);
      if (hasAppShell) { await snap(page, "020_app_shell_detected"); break; }
    }

    let seen = null;
    try {
      seen = await Promise.race([
        safeWaitForSelector(page, "button.gsi-material-button",          { timeout: 3000, visible: true }).then(()=> "google"),
        safeWaitForSelector(page, "[aria-label*='Sign in with Google']", { timeout: 3000, visible: true }).then(()=> "google"),
        safeWaitForSelector(page, "button[data-provider='google']",      { timeout: 3000, visible: true }).then(()=> "google"),
        safeWaitForSelector(page, "a[data-provider='google']",           { timeout: 3000, visible: true }).then(()=> "google"),
        safeWaitForSelector(page, "input[type='email']",                 { timeout: 3000, visible: true }).then(()=> "adv"),
        safeWaitForSelector(page, "input[name='email']",                 { timeout: 3000, visible: true }).then(()=> "adv"),
        safeWaitForSelector(page, "input[type='password']",              { timeout: 3000, visible: true }).then(()=> "adv"),
        safeWaitForSelector(page, "button[type='submit']",               { timeout: 3000, visible: true }).then(()=> "adv"),
      ]);
    } catch {}

    if (seen) {
      if (seen === "google") {
        await snap(page, "030_detected_google_option");
        await handleGoogleSSO(page, browser);
      }
      const hasAdvField = await page.$("input[type='email'], input[name='email'], input[type='password']").catch(()=>null);
      if (hasAdvField) {
        console.log(">> Detected Advantage username/password form; attempting login.");
        await handleAdvantageCredentials(page);
      }
    } else {
      await sleep(800);
      await snap(page, "031_waiting_for_login_cues");
    }
  }
}

/* =========================
   Disapprovals response mapping
   ========================= */
const TYPE_MAP = {
  3:  "Segment",
  7:  "Keyword",
  8:  "Phone",
  10: "Callout",
  11: "Sitelink",
  13: "Ad",
  15: "structured snippet",
  20: "unknown",
};

const PLATFORM_HEADERS = [
  "ID", "Campaign ID", "Engine", "Type", "Policy Violation",
  "Reason", "Reason Type", "Text", "Synchronous Disapproval?",
  "Exemptible?", "Runnable?", "Disapproved On", "Last Updated By",
  "Last Updated On", "Valid?", "Resolution?", "% Fulfilled",
  "% Time Used", "Campaign Status",
];

function boolToYesNo(val) {
  if (val === true || val === "true") return "Yes";
  return "No";
}

function fractionToPercent(val) {
  const n = Number(val);
  if (!Number.isFinite(n)) return "";
  return (n * 100).toFixed(2) + "%";
}

function buildReason(obj) {
  const reason = obj.reason ?? "";
  const violatingText = obj.violatingText ?? "";
  if (reason) return reason;
  return "Violating Text: " + violatingText;
}

function transformRow(obj) {
  const typeNum = Number(obj.type);
  return [
    String(obj.id ?? ""),
    String(obj.merchantId ?? ""),
    String(obj.engineName ?? ""),
    TYPE_MAP[typeNum] || "unknown",
    String(obj.policyViolation ?? ""),
    buildReason(obj),
    String(obj.reasonType ?? ""),
    String(obj.text ?? ""),
    boolToYesNo(obj.isSynchronousDisapproval),
    boolToYesNo(obj.isExemptable),
    boolToYesNo(obj.isRunnable),
    String(obj.disapprovedOn ?? ""),
    String(obj.lastUpdatedBy ?? ""),
    String(obj.lastUpdatedOn ?? ""),
    String(obj.validityForDisplay ?? ""),
    String(obj.resolutionForDisplay ?? ""),
    fractionToPercent(obj.fractionFulfilled),
    fractionToPercent(obj.fractionOfTimeUsed),
    String(obj.merchantStatus ?? ""),
  ];
}

function extractObjects(json) {
  if (!json) return null;
  const result = json.result;
  if (result == null) return null;

  if (Array.isArray(result) && result.length > 0 && typeof result[0] === "object" && !Array.isArray(result[0])) {
    return result;
  }

  if (typeof result === "object" && !Array.isArray(result)) {
    for (const key of ["rows", "data", "disapprovals", "items", "tableData"]) {
      const nested = result[key];
      if (Array.isArray(nested) && nested.length > 0 && typeof nested[0] === "object" && !Array.isArray(nested[0])) {
        return nested;
      }
    }
  }

  return null;
}

function mapResponseToRows(json) {
  const objects = extractObjects(json);
  if (!objects || objects.length === 0) return null;

  const dataRows = objects.map(obj => transformRow(obj));
  return [PLATFORM_HEADERS, ...dataRows];
}

/* =========================
   Local CSV conversion (fallback)
   ========================= */
function rowsToCsv(rows) {
  return rows.map(row =>
    row.map(cell => {
      const s = String(cell ?? "");
      if (s.includes(",") || s.includes('"') || s.includes("\n") || s.includes("\r")) {
        return '"' + s.replace(/"/g, '""') + '"';
      }
      return s;
    }).join(",")
  ).join("\n");
}

/* =========================
   Main
   ========================= */
(async () => {
  if (!EMAIL) {
    console.error("Missing AR_EMAIL");
    process.exit(1);
  }

  if (!fs.existsSync(OUTPUT_DIR)) fs.mkdirSync(OUTPUT_DIR, { recursive: true });

  if (APPS_SCRIPT_WEBHOOK) {
    console.log("Google Sheets export enabled (Apps Script webhook configured).");
  } else {
    console.warn("WARNING: AR_APPS_SCRIPT_URL not set — CSV-only mode.");
  }

  console.log(`Disapprovals Tracker – ${CLIENTS.length} client(s) to process`);
  console.log("Clients:", CLIENTS.map(c => `${c.name} [${c.id}]`).join(", "));

  const browser = await puppeteer.launch({
    headless: HEADLESS ? "new" : false,
    executablePath: CHROME,
    userDataDir: PROFILE,
    defaultViewport: { width: 1400, height: 900 },
    args: [
      "--no-sandbox",
      "--disable-features=BlockThirdPartyCookies",
      "--window-size=1400,900",
      "--start-minimized",
    ],
  });

  await wireDialogAutoAccept(browser);

  const summaryRows = [];

  // 1) Advantage -> hands-free login
  const page = await browser.newPage();
  await page.goto(ADV_HOME, { waitUntil: "domcontentloaded" }).catch(() => {});
  await settleHandsFreeLogin(page, browser);
  await clickModalOk(page, { timeoutMs: 4000 });

  // 2) Poll for MCID on SuperUsers (same as PPR)
  let poll = await browser.newPage();
  const maxMs = 10 * 60 * 1000;
  const start = Date.now();
  let has = false;

  while (Date.now() - start < maxMs) {
    try {
      await poll.goto(`${SUPER}/`, { waitUntil: "domcontentloaded", timeout: 60000 }).catch(() => {});
      await poll.goto(SUPER_HOME, { waitUntil: "domcontentloaded", timeout: 60000 }).catch(() => {});
      await snap(poll, "100_poll_super_home");
      has = await ensureMCID(poll);
      if (has) {
        console.log(">> MCID detected. Continuing…");
        await snap(poll, "101_mcid_detected");
        break;
      }
    } catch (e) {
      console.log("Poll error (will retry):", e.message || e);
      try { await poll.close().catch(()=>{}); } catch {}
      poll = await browser.newPage();
      await snap(poll, "099_poll_reopened");
    }
    const remain = Math.ceil((maxMs - (Date.now() - start)) / 1000);
    console.log(`Waiting for SSO to finish… (${remain}s left)`);
    await sleep(3000);
  }
  if (!has) {
    await snap(poll, "102_mcid_timeout");
    console.error("Timed out waiting for MCID.");
    await browser.close();
    process.exit(1);
  }

  await poll.bringToFront().catch(() => {});
  await clickModalOk(poll, { timeoutMs: 4000 });

  // 3) Create a dedicated ADV page for disapprovals API calls
  const advPage = await browser.newPage();
  await advPage.goto(ADV_HOME, { waitUntil: "domcontentloaded", timeout: 60000 }).catch(() => {});
  await snap(advPage, "150_adv_page_loaded");

  // 4) Process each client
  for (const client of CLIENTS) {
    console.log("==================================================");
    console.log(`Disapprovals: ${client.name} [${client.id}]`);

    let fetched = false;
    let saved = false;
    let sheetUrl = "";
    let moved = false;

    // a) changeClient on SuperUsers (updates session to this client)
    const cc = await changeClientRobust(poll, EMAIL, client.id);
    if (!cc.ok) {
      console.error(
        `changeClient failed for clientId=${client.id}. Raw:`,
        JSON.stringify(cc.res?.json || cc.res || {}, null, 2)
      );
      await snap(poll, `change_client_failed_${client.id}`);
      summaryRows.push({ client: client.name, fetched: false, saved: false, sheetUrl: "", moved: false });
      continue;
    }
    VERSION = cc.version;
    await clickModalOk(poll, { timeoutMs: 4000 });

    // b) getSearchEngineDisapprovals from ADV page
    //    Body matches the exact format from the original cURL:
    //    { currentClientId, currentLoginName, version, data: { clientId } }
    const disRes = await postInside(advPage, GET_DISAPPROVALS_URL, {
      currentClientId: client.id,
      currentLoginName: EMAIL,
      version: VERSION,
      data: { clientId: client.id },
    });

    const shortRes = (disRes.text || "").slice(0, 400);
    console.log(`getSearchEngineDisapprovals [${client.id}]: status=${disRes.status}`);
    console.log(`  Response preview: ${shortRes}`);
    await snap(advPage, `disapprovals_${client.id}_${disRes.status}`);

    if (disRes.status !== 200 || !disRes.json) {
      console.error(`API error for ${client.id}: ${disRes.status} ${shortRes}`);
      summaryRows.push({ client: client.name, fetched: false, saved: false, sheetUrl: "", moved: false });
      continue;
    }

    // Check for API-level error
    if (disRes.json.error && disRes.json.error.code) {
      console.error(`API error for ${client.id}: ${JSON.stringify(disRes.json.error)}`);
      summaryRows.push({ client: client.name, fetched: false, saved: false, sheetUrl: "", moved: false });
      continue;
    }

    // c) Map response to rows (header row + data rows)
    const rows = mapResponseToRows(disRes.json);

    const safeName = client.name.replace(/[^\w.-]+/g, "_");
    const outPath = path.join(OUTPUT_DIR, `Disapprovals_${safeName}_${client.id}.csv`);

    if (!rows || rows.length === 0) {
      console.log(`No disapprovals data for ${client.name} [${client.id}]`);
      console.log(`  Response keys: ${JSON.stringify(Object.keys(disRes.json || {}))}`);
      if (disRes.json && disRes.json.result != null) {
        console.log(`  result type: ${typeof disRes.json.result}, isArray: ${Array.isArray(disRes.json.result)}`);
        if (typeof disRes.json.result === "object" && !Array.isArray(disRes.json.result)) {
          console.log(`  result keys: ${JSON.stringify(Object.keys(disRes.json.result))}`);
        }
      }
      fs.writeFileSync(outPath, "No disapprovals found\n", "utf8");
      console.log(`Saved (empty): ${outPath}`);
      summaryRows.push({ client: client.name, fetched: true, saved: true, sheetUrl: "", moved: false });
      continue;
    }

    fetched = true;
    console.log(`Got ${rows.length - 1} disapproval row(s) for ${client.name}`);

    // d) Export to Google Sheets via Apps Script (Node POST with browser cookies)
    if (APPS_SCRIPT_WEBHOOK) {
      try {
        const today = new Date().toISOString().slice(0, 10);
        const sheetTitle = `Disapprovals – ${client.name} – ${today}`;

        const gCookies = await poll.cookies("https://script.google.com");
        const cookieHeader = gCookies.map(c => `${c.name}=${c.value}`).join("; ");
        console.log(`  Google cookies for script.google.com: ${gCookies.length} cookie(s), names: ${gCookies.map(c => c.name).join(", ") || "(none)"}`);

        const postResult = await nodePost(APPS_SCRIPT_WEBHOOK, {
          title: sheetTitle,
          rows,
          clientId: client.id,
          clientName: client.name,
        }, cookieHeader);

        console.log(`Apps Script response [${client.id}]: status=${postResult.status}`);
        console.log(`  Body: ${(postResult.text || "").slice(0, 400)}`);

        if (postResult.json?.success) {
          sheetUrl = postResult.json.url || "";
          moved = !!postResult.json.moved;
          console.log(`Google Sheet created: ${sheetUrl}`);
          console.log(`  Moved to drive: ${moved ? "Yes" : "No"}`);
          if (postResult.json.warning) {
            console.warn(`  Warning: ${postResult.json.warning}`);
          }
        } else {
          console.error(
            `Apps Script failed for ${client.name}: ${postResult.json?.error || postResult.text}`
          );
        }
      } catch (hookErr) {
        console.error(
          `Apps Script call failed for ${client.name}: ${hookErr?.message || hookErr}`
        );
      }
    }

    // f) Save CSV as fallback
    const csvContent = rowsToCsv(rows);
    fs.writeFileSync(outPath, csvContent, "utf8");
    console.log(`CSV fallback saved: ${outPath}`);
    saved = true;

    summaryRows.push({ client: client.name, fetched, saved, sheetUrl, moved });
  }

  // Print summary
  printSummaryTable(summaryRows);

  // Clean up
  await advPage.close().catch(() => {});
  await poll.close().catch(() => {});
  await page.close().catch(() => {});
  await browser.close();
})().catch(async (e) => {
  console.error(e);
  try {
    const fallback = path.join(SCREEN_DIR, "zzz_unhandled_error.txt");
    fs.writeFileSync(fallback, (e && e.stack) ? e.stack : String(e));
    console.log("Saved error stack:", fallback);
  } catch {}
  process.exit(1);
});
