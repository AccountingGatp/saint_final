import XLSX from "xlsx-js-style";

// ---------------------------------- helpers ----------------------------------

const num = (v) => (typeof v === "number" ? v : Number(v)) || 0;
const round2 = (v) => Math.round((v + Number.EPSILON) * 100) / 100;
const round5 = (v) => Math.round((v + Number.EPSILON) * 1e5) / 1e5;

/** Read the first sheet of an Excel workbook (from a Buffer) as an array of row objects. */
function readExcelBuffer(buffer) {
  const wb = XLSX.read(buffer, { type: "buffer" });
  const ws = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(ws, { defval: null });
}

/**
 * Parse an "Order Time" cell (e.g. "2026/6/18 23:59:41") into { md, iso }:
 *   md  -> "MM/DD"       ("06/18")     group / display key
 *   iso -> "YYYY-MM-DD"  ("2026-06-18") FX API query date
 */
function parseDate(orderTime) {
  if (orderTime === null || orderTime === undefined || String(orderTime).trim() === "") return null;
  const datePart = String(orderTime).trim().split(/\s+/)[0];
  const m = datePart.match(/^(\d{4})[/-](\d{1,2})[/-](\d{1,2})$/);
  if (!m) return null;
  const month = String(Number(m[2])).padStart(2, "0");
  const day = String(Number(m[3])).padStart(2, "0");
  return { md: `${month}/${day}`, iso: `${m[1]}-${month}-${day}` };
}

// --------------------------------- FX (RBA) ---------------------------------

// RBA table F11.1 (Exchange Rates - Daily). Free, keyless, official. Covers the
// last few years of trading days. Rates are quoted as "A$1 = <foreign currency>".
const RBA_CSV_URL = "https://www.rba.gov.au/statistics/tables/csv/f11.1-data.csv";
const RBA_MONTHS = {
  Jan: "01", Feb: "02", Mar: "03", Apr: "04", May: "05", Jun: "06",
  Jul: "07", Aug: "08", Sep: "09", Oct: "10", Nov: "11", Dec: "12",
};

let rbaTablePromise = null; // memoized: fetch + parse the CSV once per process.

/** "15-Jun-2026" -> "2026-06-15", or null if unrecognized. */
function parseRbaDate(cell) {
  const m = String(cell).trim().match(/^(\d{1,2})-([A-Za-z]{3})-(\d{4})$/);
  if (!m) return null;
  const mm = RBA_MONTHS[m[2][0].toUpperCase() + m[2].slice(1).toLowerCase()];
  return mm ? `${m[3]}-${mm}-${m[1].padStart(2, "0")}` : null;
}

/**
 * Load RBA F11.1 into { dates: [isoAsc], byDate: Map(iso -> Map(code -> value)) }.
 * `code` is the foreign-currency code (USD, CNY, EUR, ...) from the "A$1=<code>"
 * header columns; each value is units of that currency per A$1.
 */
async function loadRbaTable() {
  if (!rbaTablePromise) {
    rbaTablePromise = (async () => {
      // RBA blocks requests without a browser-like User-Agent (403 otherwise).
      const res = await fetch(RBA_CSV_URL, {
        headers: { "User-Agent": "Mozilla/5.0 (compatible; SaintFX/1.0)" },
      });
      if (!res.ok) throw new Error(`RBA fetch failed: HTTP ${res.status}`);
      const text = await res.text();
      const lines = text.replace(/^﻿/, "").split(/\r?\n/);

      let codes = null; // column index -> currency code
      const byDate = new Map();
      for (const line of lines) {
        const cells = line.split(",");
        if (cells[0] === "Title") {
          codes = cells.map((c) => {
            const m = /^A\$1=([A-Za-z]{3})$/.exec(c.trim());
            return m ? m[1].toUpperCase() : null;
          });
          continue;
        }
        const iso = parseRbaDate(cells[0]);
        if (!iso || !codes) continue;
        const rowMap = new Map();
        for (let i = 1; i < cells.length; i++) {
          const code = codes[i];
          const v = cells[i]?.trim();
          if (code && v) {
            const n = Number(v);
            if (Number.isFinite(n)) rowMap.set(code, n);
          }
        }
        byDate.set(iso, rowMap);
      }
      if (!byDate.size) throw new Error("RBA table parsed to zero rows.");
      return { dates: [...byDate.keys()].sort(), byDate };
    })().catch((err) => {
      rbaTablePromise = null; // allow a later retry
      throw err;
    });
  }
  return rbaTablePromise;
}

/** Units of `code` per A$1 for a given RBA row map (AUD itself is 1). */
function rbaValue(rowMap, code) {
  return code === "AUD" ? 1 : rowMap.get(code) ?? null;
}

/**
 * base -> to conversion rate for a date, from the RBA F11.1 daily table.
 * Falls back to the most recent trading day on-or-before the requested date
 * (RBA publishes business days only). Null when unavailable.
 *
 * A$1 = value(base) base = value(to) to  =>  1 base = value(to)/value(base) to.
 */
async function fetchRate(isoDate, to = "AUD", base = "USD") {
  try {
    const { dates, byDate } = await loadRbaTable();
    let key = byDate.has(isoDate) ? isoDate : null;
    if (!key) {
      // nearest trading day on-or-before isoDate (dates sorted ascending ISO)
      for (let i = dates.length - 1; i >= 0; i--) {
        if (dates[i] <= isoDate) { key = dates[i]; break; }
      }
    }
    if (!key) return null;
    const rowMap = byDate.get(key);
    const vBase = rbaValue(rowMap, base.toUpperCase());
    const vTo = rbaValue(rowMap, to.toUpperCase());
    if (!vBase || vTo === null) return null;
    return vTo / vBase;
  } catch {
    return null;
  }
}

/** Paint a worksheet's header row (row 1) Excel blue with white bold centered text. */
function styleHeader(worksheet) {
  const style = {
    fill: { patternType: "solid", fgColor: { rgb: "FF4472C4" } },
    font: { bold: true, color: { rgb: "FFFFFFFF" } },
    alignment: { horizontal: "center", vertical: "center" },
  };
  const range = XLSX.utils.decode_range(worksheet["!ref"] || "A1");
  for (let c = range.s.c; c <= range.e.c; c++) {
    const addr = XLSX.utils.encode_cell({ r: 0, c });
    if (!worksheet[addr]) worksheet[addr] = { t: "s", v: "" };
    worksheet[addr].s = style;
  }
}

// --------------------------------- transforms ---------------------------------

/**
 * Sheet1: original rows with "Order ID" (col B) and "Date" (col C) inserted.
 * Order ID = "Clients Order Number" without "#", filled down for blank line items.
 */
function buildSheet1(rows, date, orderColumn = "Clients Order Number") {
  let lastOrderId = "";
  return rows.map((row) => {
    const raw = row[orderColumn];
    if (raw !== null && raw !== undefined && String(raw).trim() !== "") {
      lastOrderId = String(raw).replace(/#/g, "").trim();
    }
    const keys = Object.keys(row);
    const [firstKey, ...restKeys] = keys;
    const out = {};
    out[firstKey] = row[firstKey];
    out["Order ID"] = lastOrderId;
    out["Date"] = date;
    for (const k of restKeys) out[k] = row[k];
    return out;
  });
}

/**
 * Sheet2: per-day summary with target-currency conversion columns and a Grand Total row.
 * Grouped by order date (from "Order Time"); blank Order Time inherits prior date.
 */
async function buildSheet2(rows, currency) {
  const totals = new Map(); // "MM/DD" -> { product, shipping, iso }
  let last = null;
  for (const row of rows) {
    const parsed = parseDate(row["Order Time"]);
    if (parsed) last = parsed;
    const d = parsed || last;
    if (!d) continue;
    if (!totals.has(d.md)) totals.set(d.md, { product: 0, shipping: 0, iso: d.iso });
    const b = totals.get(d.md);
    b.product += num(row["Product Price($)"]);
    b.shipping += num(row["Shipping Cost ($)"]);
  }

  const dates = [...totals.keys()].sort((a, b) => {
    const [am, ad] = a.split("/").map(Number);
    const [bm, bd] = b.split("/").map(Number);
    return am - bm || ad - bd;
  });

  const rates = await Promise.all(dates.map((d) => fetchRate(totals.get(d).iso, currency)));

  const totalCol = `Total in ${currency}`;
  const productCol = `Product ${currency}`;
  const shippingCol = `Shipping ${currency}`;

  const summary = [];
  let gP = 0, gS = 0, gT = 0, gPc = 0, gSc = 0;

  dates.forEach((date, i) => {
    const { product, shipping } = totals.get(date);
    const rate = rates[i];
    if (rate === null) console.warn(`Warning: no FX rate for ${date} (${totals.get(date).iso}).`);
    const total = product + shipping;
    gP += product; gS += shipping;
    if (rate !== null) { gT += total * rate; gPc += product * rate; gSc += shipping * rate; }
    summary.push({
      Date: date,
      "Sum of Product Price($)": round2(product),
      "Sum of Shipping Cost ($)": round2(shipping),
      "Total(product+Shipping)": round2(total),
      "Rate Conversion": rate === null ? "" : round5(rate),
      [totalCol]: rate === null ? "" : round2(total * rate),
      [productCol]: rate === null ? "" : round2(product * rate),
      [shippingCol]: rate === null ? "" : round2(shipping * rate),
    });
  });

  summary.push({
    Date: "Grand Total",
    "Sum of Product Price($)": round2(gP),
    "Sum of Shipping Cost ($)": round2(gS),
    "Total(product+Shipping)": round2(gP + gS),
    "Rate Conversion": "",
    [totalCol]: round2(gT),
    [productCol]: round2(gPc),
    [shippingCol]: round2(gSc),
  });

  return summary;
}

// ----------------------------------- main -----------------------------------

/**
 * Build one workbook (as a Buffer) with three sheets from an uploaded file buffer:
 *   "Input"  - the source rows, unchanged
 *   "Sheet1" - Input + Order ID (B) + Date (C)
 *   "Sheet2" - per-day summary with USD->currency conversion (blue header)
 */
export async function buildWorkbookBuffer(inputBuffer, { date = "06/26", currency = "AUD" } = {}) {
  const rows = readExcelBuffer(inputBuffer);

  const inputSheet = XLSX.utils.json_to_sheet(rows);
  styleHeader(inputSheet);
  const sheet1 = XLSX.utils.json_to_sheet(buildSheet1(rows, date));
  styleHeader(sheet1);
  const summary = await buildSheet2(rows, currency);
  const sheet2 = XLSX.utils.json_to_sheet(summary);
  styleHeader(sheet2);

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, inputSheet, "Input");
  XLSX.utils.book_append_sheet(wb, sheet1, "Sheet1");
  XLSX.utils.book_append_sheet(wb, sheet2, "Sheet2");

  return XLSX.write(wb, { type: "buffer", bookType: "xlsx" });
}

// ------------------------------- import export -------------------------------

const MONTHS = [
  "January", "February", "March", "April", "May", "June",
  "July", "August", "September", "October", "November", "December",
];

// Fixed chart-of-accounts mapping for the accounting import (edit here to change).
const IMPORT_TAX_RATE = "BAS Excluded";
const IMPORT_LINES = [
  { description: "Shipping Cost", accountCode: 310.1, source: "shipping" },
  { description: "Cost of Goods Sold", accountCode: 310, source: "product" },
  { description: "Inventory", accountCode: 631, source: "inventory" },
];
const IMPORT_HEADER = [
  "*Narration", "*Date", "Description", "*AccountCode", "*TaxRate", "*Amount",
  "TrackingName1", "TrackingOption1", "TrackingName2", "TrackingOption2",
];

/**
 * Build the accounting import workbook (as a Buffer) from a previously generated
 * output workbook. Reads "Sheet2" for the per-day AUD totals and recovers each
 * day's year from the original "Order Time" values on the "Input" sheet.
 *
 * Each date produces one balanced journal block:
 *   Shipping Cost (310.1)      =  Shipping <currency>
 *   Cost of Goods Sold (310)   =  Product  <currency>
 *   Inventory (631)            = -(Total in <currency>)
 */
export function buildImportBuffer(generatedBuffer) {
  const wb = XLSX.read(generatedBuffer, { type: "buffer" });
  const sheet2 = wb.Sheets["Sheet2"];
  if (!sheet2) {
    throw new Error("This file has no 'Sheet2'. Upload a processed output workbook.");
  }
  const summary = XLSX.utils.sheet_to_json(sheet2, { defval: null });
  const inputRows = wb.Sheets["Input"]
    ? XLSX.utils.sheet_to_json(wb.Sheets["Input"], { defval: null })
    : [];

  // Recover MM/DD -> year from the original order dates (Sheet2 keeps only MM/DD).
  const yearByMd = new Map();
  let last = null;
  for (const row of inputRows) {
    const parsed = parseDate(row["Order Time"]);
    if (parsed) last = parsed;
    const d = parsed || last;
    if (d && !yearByMd.has(d.md)) yearByMd.set(d.md, d.iso.slice(0, 4));
  }

  // Locate the target-currency columns by prefix ("Product AUD", "Shipping AUD", "Total in AUD").
  const keys = summary.length ? Object.keys(summary[0]) : [];
  const productCol = keys.find((k) => /^Product /.test(k));
  const shippingCol = keys.find((k) => /^Shipping /.test(k));
  const totalCol = keys.find((k) => /^Total in /.test(k));

  const out = [];
  for (const row of summary) {
    const md = row.Date == null ? "" : String(row.Date).trim();
    if (!md || md.toLowerCase() === "grand total") continue;

    const product = round2(num(row[productCol]));
    const shipping = round2(num(row[shippingCol]));
    const total =
      row[totalCol] === "" || row[totalCol] == null
        ? round2(product + shipping)
        : round2(num(row[totalCol]));

    const [mm, dd] = md.split("/");
    const year = yearByMd.get(md) || "";
    const monthName = MONTHS[Number(mm) - 1] || mm;
    const narration = `COGS Supplier Cost Sheet - ${Number(dd)}-${monthName}-${year}`;
    const dateStr = `${dd}/${mm}/${year}`;

    const amounts = { shipping, product, inventory: -total };
    for (const line of IMPORT_LINES) {
      out.push({
        "*Narration": narration,
        "*Date": dateStr,
        Description: line.description,
        "*AccountCode": line.accountCode,
        "*TaxRate": IMPORT_TAX_RATE,
        "*Amount": amounts[line.source],
        TrackingName1: "",
        TrackingOption1: "",
        TrackingName2: "",
        TrackingOption2: "",
      });
    }
  }

  if (!out.length) {
    throw new Error("No dated rows found in 'Sheet2' to build an import from.");
  }

  const ws = XLSX.utils.json_to_sheet(out, { header: IMPORT_HEADER });
  styleHeader(ws);

  const outWb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(outWb, ws, "Import");
  return XLSX.write(outWb, { type: "buffer", bookType: "xlsx" });
}
