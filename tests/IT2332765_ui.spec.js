const { test, expect } = require("@playwright/test");
const ExcelJS = require("exceljs");
const XLSX = require("xlsx");
const fs = require("fs");
const path = require("path");

const URL = "https://www.swifttranslator.com/";
const EXCEL_NAME = "IT23327658 - ITPM Excel.xlsx";
const OUT_EXCEL_NAME = "IT23327658 - ITPM Excel_EXECUTED.xlsx";

const ROOT = path.join(__dirname, "..");
const EXCEL_PATH = path.join(ROOT, EXCEL_NAME);
const OUT_EXCEL_PATH = path.join(ROOT, OUT_EXCEL_NAME);

function s(v) {
  if (v === null || v === undefined) return "";
  return String(v);
}
function norm(v) {
  return s(v).replace(/\r/g, "").replace(/\s+/g, " ").trim();
}
function canon(v) {
  return norm(v).replace(/\u00A0/g, " ").trim();
}
function crop(v, n = 350) {
  const t = canon(v);
  return t.length > n ? t.slice(0, n - 1) + "…" : t;
}

function pickSheetXlsx(wb) {
  const names = wb.SheetNames || [];
  const exact = names.find((n) => n === " Test cases");
  if (exact) return exact;
  const byName = names.find((n) => norm(n).toLowerCase().includes("test cases"));
  if (byName) return byName;
  const anyTest = names.find((n) => norm(n).toLowerCase().includes("test"));
  return anyTest || names[0];
}

function findHeaderRowAndCols(rows) {
  const max = Math.min(rows.length, 100);
  for (let i = 0; i < max; i++) {
    const r = rows[i] || [];
    const low = r.map((x) => norm(x).toLowerCase());
    const hasTc = low.some((x) => x === "tc id" || x.includes("tc id"));
    const hasInput = low.some((x) => x === "input");
    const hasExpected = low.some((x) => x.includes("expected output"));
    const hasActual = low.some((x) => x.includes("actual output"));
    const hasStatus = low.some((x) => x === "status");
    if (hasTc && hasInput && hasExpected && hasActual && hasStatus) {
      const col = (needle) => low.findIndex((x) => x === needle || x.includes(needle));
      return {
        headerRowIndex0: i,
        tcCol0: col("tc id"),
        inputCol0: low.findIndex((x) => x === "input"),
        expectedCol0: col("expected output"),
        actualCol0: col("actual output"),
        statusCol0: low.findIndex((x) => x === "status")
      };
    }
  }
  return null;
}

function loadUiCasesSync() {
  const wb = XLSX.readFile(EXCEL_PATH);
  const sn = pickSheetXlsx(wb);
  const ws = wb.Sheets[sn];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });

  const meta = findHeaderRowAndCols(rows);
  if (!meta) throw new Error("Header row not found in Test cases sheet");

  const cases = [];
  for (let r0 = meta.headerRowIndex0 + 1; r0 < rows.length; r0++) {
    const row = rows[r0] || [];
    const tcId = canon(row[meta.tcCol0]);
    if (!tcId) continue;
    if (!/^UI_Fun_\d{4}$/.test(tcId)) continue;

    const input = canon(row[meta.inputCol0]);
    const expected = canon(row[meta.expectedCol0]);

    cases.push({
      tcId,
      input,
      expected,
      rowNum1: r0 + 1
    });
  }

  return {
    sheetName: sn,
    cols1: {
      actual: meta.actualCol0 + 1,
      status: meta.statusCol0 + 1
    },
    cases
  };
}

const UI = loadUiCasesSync();
const results = new Map();

async function gotoApp(page) {
  await page.goto(URL, { waitUntil: "domcontentloaded" });
  const input = page.locator('textarea[placeholder*="Input Your Singlish"], textarea').first();
  await expect(input).toBeVisible({ timeout: 45000 });
  return input;
}

async function clickIfExists(page) {
  const btn = page.getByRole("button", { name: /translate|convert|transliterate|submit|generate/i }).first();
  try {
    if ((await btn.count()) > 0) await btn.click({ timeout: 800 }).catch(() => {});
  } catch {}
}

async function readOutputSmart(page) {
  return await page.evaluate(() => {
    const vis = (el) => {
      const r = el.getBoundingClientRect();
      const s = window.getComputedStyle(el);
      return r.width > 0 && r.height > 0 && s.visibility !== "hidden" && s.display !== "none";
    };

    const input =
      document.querySelector('textarea[placeholder*="Input Your Singlish"]') ||
      document.querySelector("textarea");

    const pri = Array.from(
      document.querySelectorAll(
        "textarea[readonly], textarea[disabled], #result, #output, .output, [id*='result'], [id*='output'], [class*='result'], [class*='output'], [aria-label*='output' i]"
      )
    ).filter(vis);

    const getText = (el) => {
      if (!el) return "";
      if (el.tagName === "TEXTAREA" || el.tagName === "INPUT") return (el.value || "").trim();
      return (el.innerText || "").trim();
    };

    for (const el of pri) {
      if (input && el === input) continue;
      const t = getText(el);
      if (t) return t;
    }

    const all = Array.from(document.querySelectorAll("textarea,input,div,span,p,pre,output,section,article")).filter(vis);
    const SIN = /[\u0D80-\u0DFF]/;

    let bestSin = "";
    for (const el of all) {
      if (input && el === input) continue;
      const t = getText(el);
      if (t && SIN.test(t) && t.length > bestSin.length) bestSin = t;
    }
    if (bestSin) return bestSin;

    let best = "";
    for (const el of all) {
      if (input && el === input) continue;
      const t = getText(el);
      if (t && t.length > best.length) best = t;
    }
    return best || "";
  });
}

async function waitForOutput(page, prev, timeoutMs) {
  const t0 = Date.now();
  const p = canon(prev);
  let last = "";
  while (Date.now() - t0 < timeoutMs) {
    const cur = canon(await readOutputSmart(page));
    if (cur && cur !== p && cur !== last) return cur;
    last = cur;
    await page.waitForTimeout(250);
  }
  return canon(await readOutputSmart(page));
}

function parseUiInput(text) {
  const t = canon(text);
  const m = t.match(/type\s+(.+?)\s+then\s+change\s+to\s+(.+)/i);
  if (m) return [m[1].trim(), m[2].trim()];
  if (t) return [t, t + " test"];
  return ["mama pansal yanavaa", "api canteen ekata yamu."];
}

test.describe("SwiftTranslator - UI", () => {
  test.setTimeout(120000);

  test("Excel UI cases loaded", async () => {
    expect(UI.cases.length).toBeGreaterThan(0);
  });

  for (const c of UI.cases) {
    test(c.tcId, async ({ page }) => {
      const input = await gotoApp(page);
      const [firstText, secondText] = parseUiInput(c.input);

      const prev = await readOutputSmart(page);

      await input.click();
      await input.fill(firstText);
      await clickIfExists(page);

      const out1 = await waitForOutput(page, prev, 25000);

      await input.click();
      await input.fill(secondText);
      await clickIfExists(page);

      const out2 = await waitForOutput(page, out1, 25000);

      const changed = canon(out1) && canon(out2) && canon(out1) !== canon(out2);
      const status = changed ? "Pass" : "Fail";
      const actual = `out1=${crop(out1, 140)} | out2=${crop(out2, 140)}`;

      results.set(c.tcId, { actual, status });

      expect(canon(out1).length).toBeGreaterThan(0);
      expect(canon(out2).length).toBeGreaterThan(0);
      expect(changed).toBeTruthy();
    });
  }

  test.afterAll(async () => {
    if (!UI.cases.length) return;

    const base = fs.existsSync(OUT_EXCEL_PATH) ? OUT_EXCEL_PATH : EXCEL_PATH;

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(base);
    const ws = wb.getWorksheet(" Test cases") || wb.worksheets.find(w => norm(w.name).toLowerCase().includes("test")) || wb.worksheets[0];

    for (const c of UI.cases) {
      const res = results.get(c.tcId);
      if (!res) continue;
      const row = ws.getRow(c.rowNum1);
      row.getCell(UI.cols1.actual).value = res.actual;
      row.getCell(UI.cols1.status).value = res.status;
      row.commit();
    }

    await wb.xlsx.writeFile(OUT_EXCEL_PATH);
  });
});
