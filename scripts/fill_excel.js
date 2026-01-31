const XLSX = require("xlsx");
const fs = require("fs");
const path = require("path");

const EXCEL_IN = path.join(__dirname, "..", "IT23327658 - ITPM Excel.xlsx");
const EXCEL_OUT = path.join(__dirname, "..", "IT23327658 - ITPM Excel_EXECUTED.xlsx");
const RESULTS = path.join(__dirname, "..", "test-results", "execution-results.jsonl");
const SHEET = " Test cases";

const HEADER_ROW_INDEX = 4;
const COL_TC = 0;
const COL_ACTUAL = 5;
const COL_STATUS = 6;
const COL_ISSUE = 7;

function norm(v) {
  return (v ?? "").toString().replace(/\r/g, "").trim();
}

function loadResults() {
  if (!fs.existsSync(RESULTS)) throw new Error(`Results not found: ${RESULTS}`);
  const lines = fs.readFileSync(RESULTS, "utf8").split("\n").map(l => l.trim()).filter(Boolean);
  const map = new Map();
  for (const line of lines) {
    const r = JSON.parse(line);
    if (r && r.tcId) map.set(norm(r.tcId), r);
  }
  return map;
}

function setCell(ws, r, c, v) {
  const addr = XLSX.utils.encode_cell({ r, c });
  ws[addr] = ws[addr] || { t: "s", v: "" };
  ws[addr].t = "s";
  ws[addr].v = v;
}

function main() {
  if (!fs.existsSync(EXCEL_IN)) throw new Error(`Excel not found: ${EXCEL_IN}`);

  const wb = XLSX.readFile(EXCEL_IN, { cellStyles: true });
  const ws = wb.Sheets[SHEET];
  if (!ws) throw new Error(`Sheet not found: ${SHEET}`);

  const range = XLSX.utils.decode_range(ws["!ref"]);
  const resMap = loadResults();

  if (norm(ws[XLSX.utils.encode_cell({ r: HEADER_ROW_INDEX, c: COL_TC })]?.v) !== "TC ID")
    throw new Error("Header row not detected at expected index");

  for (let r = HEADER_ROW_INDEX + 1; r <= range.e.r; r++) {
    const tcAddr = XLSX.utils.encode_cell({ r, c: COL_TC });
    const tcId = norm(ws[tcAddr]?.v);
    if (!tcId) continue;

    const rr = resMap.get(tcId);
    if (!rr) continue;

    setCell(ws, r, COL_ACTUAL, norm(rr.actual));
    setCell(ws, r, COL_STATUS, norm(rr.status));

    const issue = norm(rr.status) === "Fail" ? norm(rr.issue || "Fail") : "";
    setCell(ws, r, COL_ISSUE, issue);
  }

  XLSX.writeFile(wb, EXCEL_OUT);
  console.log(EXCEL_OUT);
}

main();
