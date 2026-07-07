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

/** USD -> target rate for a date from the free, keyless Frankfurter API. Null on failure. */
async function fetchRate(isoDate, to = "AUD", base = "USD") {
  try {
    const res = await fetch(`https://api.frankfurter.dev/v1/${isoDate}?base=${base}&symbols=${to}`);
    if (!res.ok) return null;
    const data = await res.json();
    return data?.rates?.[to] ?? null;
  } catch {
    return null;
  }
}

/** Paint a worksheet's header row (row 1) blue with white bold centered text. */
function styleHeader(worksheet) {
  const style = {
    fill: { patternType: "solid", fgColor: { rgb: "FF4472C4" } },
    font: { bold: true, color: { rgb: "FFFFFFFF" } },
    alignment: { horizontal: "center", vertical: "center" },
  };
  const range = XLSX.utils.decode_range(worksheet["!ref"]);
  for (let c = range.s.c; c <= range.e.c; c++) {
    const addr = XLSX.utils.encode_cell({ r: 0, c });
    if (worksheet[addr]) worksheet[addr].s = style;
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
