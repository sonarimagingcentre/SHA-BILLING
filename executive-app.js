const state = {
  allRecords: [],
  filteredRecords: [],
  vendorRecords: [],
  dateBounds: null,
  fullReconciliation: null,
  charts: {
    weekly: null,
    modality: null,
    insuranceShare: null,
  },
};

const MODALITIES = ["MRI", "CT", "Ultrasound", "X-Ray", "Sedation", "Other"];
const SHA_NAME = "social health insurance fund";
const TOPUP_NAME = "nbi cash customer";
const DEFAULT_SOURCE_PATH = "data/SHA_Billing_Data.xlsx";

const chartColors = {
  sha: "#0b5ba0",
  topup: "#c8720f",
  total: "#12805c",
  mri: "#0b5ba0",
  ct: "#12805c",
  us: "#5caef3",
  xray: "#7fc2ff",
  sedation: "#98abd0",
  other: "#94a3b8",
};

if (typeof Chart !== "undefined") {
  Chart.defaults.font.family = "'Manrope', 'Segoe UI', sans-serif";
  Chart.defaults.color = "#64748b";
}

const refs = {
  fromDate: document.getElementById("fromDate"),
  toDate: document.getElementById("toDate"),
  applyDates: document.getElementById("applyDates"),
  reportingPeriod: document.getElementById("reportingPeriod"),
  dataAsAt: document.getElementById("dataAsAt"),
  summaryPeriod: document.getElementById("summaryPeriod"),
  executiveSummaryList: document.getElementById("executiveSummaryList"),

  kpiShaBilling: document.getElementById("kpiShaBilling"),
  kpiTopup: document.getElementById("kpiTopup"),
  kpiInsuranceTopup: document.getElementById("kpiInsuranceTopup"),
  kpiTotalRevenue: document.getElementById("kpiTotalRevenue"),
  kpiStudies: document.getElementById("kpiStudies"),
  kpiCoverage: document.getElementById("kpiCoverage"),
  kpiPatientPct: document.getElementById("kpiPatientPct"),

  mriStudies: document.getElementById("mriStudies"),
  mriSha: document.getElementById("mriSha"),
  mriCopay: document.getElementById("mriCopay"),
  mriTopup: document.getElementById("mriTopup"),
  mriRevenue: document.getElementById("mriRevenue"),
  mriAvgSha: document.getElementById("mriAvgSha"),
  mriAvgTopup: document.getElementById("mriAvgTopup"),
  mriAvgCopay: document.getElementById("mriAvgCopay"),
  mriShare: document.getElementById("mriShare"),

  ctStudies: document.getElementById("ctStudies"),
  ctSha: document.getElementById("ctSha"),
  ctCopay: document.getElementById("ctCopay"),
  ctTopup: document.getElementById("ctTopup"),
  ctRevenue: document.getElementById("ctRevenue"),
  ctAvgSha: document.getElementById("ctAvgSha"),
  ctAvgTopup: document.getElementById("ctAvgTopup"),
  ctAvgCopay: document.getElementById("ctAvgCopay"),
  ctShare: document.getElementById("ctShare"),

  insuranceTableBody: document.querySelector("#insuranceTable tbody"),
  procedureTableBody: document.querySelector("#procedureTable tbody"),

  recoSourceTotal: document.getElementById("recoSourceTotal"),
  recoShaTotal: document.getElementById("recoShaTotal"),
  recoNbiTotal: document.getElementById("recoNbiTotal"),
  recoMatchedTotal: document.getElementById("recoMatchedTotal"),
  recoUnmatchedShaTotal: document.getElementById("recoUnmatchedShaTotal"),
  recoUnmatchedNbiTotal: document.getElementById("recoUnmatchedNbiTotal"),
  recoOtherTotal: document.getElementById("recoOtherTotal"),
  recoMatchedCount: document.getElementById("recoMatchedCount"),
  matchedPairsBody: document.querySelector("#matchedPairsTable tbody"),
  unmatchedShaBody: document.querySelector("#unmatchedShaTable tbody"),
  unmatchedCashBody: document.querySelector("#unmatchedCashTable tbody"),
};

function initialize() {
  refs.applyDates.addEventListener("click", applyDateRange);
  loadDefaultSource();
}

async function loadDefaultSource() {
  try {
    const records = await loadWorkbookRecords();
    hydrateDashboard(records);
  } catch (error) {
    console.error(error);
    clearDashboard();
  }

  try {
    const vendorRecords = await loadVendorWorkbookRecords();
    state.vendorRecords = vendorRecords;
    refreshDashboard();
  } catch (error) {
    console.warn("Vendor/provider full billing data could not be loaded; that table will be empty.", error);
  }
}

async function loadVendorWorkbookRecords() {
  const inlineBase64 = window.__VENDOR_BILLING_DATA_BASE64;
  if (typeof inlineBase64 === "string" && inlineBase64.length > 100) {
    return parseWorkbookData(base64ToArrayBuffer(inlineBase64));
  }
  throw new Error("Vendor billing data not bundled.");
}

async function loadWorkbookRecords() {
  const inlineBase64 = window.__SHA_BILLING_DATA_BASE64;
  if (typeof inlineBase64 === "string" && inlineBase64.length > 100) {
    try {
      return parseWorkbookData(base64ToArrayBuffer(inlineBase64));
    } catch (error) {
      console.warn("Inline bundled data failed to load.", error);
    }
  }

  try {
    const response = await fetch(encodeURI(DEFAULT_SOURCE_PATH), { cache: "no-store" });
    if (response.ok) {
      const data = await response.arrayBuffer();
      const records = parseWorkbookData(data);
      if (records.length) {
        return records;
      }
    }
  } catch (error) {
    console.warn("Primary workbook fetch failed.", error);
  }

  throw new Error("Bundled dashboard data could not be loaded.");
}

function hydrateDashboard(records) {
  if (!records.length) {
    clearDashboard();
    return;
  }

  state.allRecords = records;
  state.dateBounds = getDateBounds(records);
  // The picker's selectable range (min/max) must span the whole dataset,
  // not just the default window shown on load — otherwise earlier months
  // become impossible to select at all.
  applyDateBounds(state.dateBounds);
  setDateInputs(state.dateBounds.minDate, state.dateBounds.maxDate);
  refreshDashboard();
}

function parseWorkbookData(arrayBuffer) {
  const data = new Uint8Array(arrayBuffer);
  const workbook = XLSX.read(data, { type: "array" });
  const firstSheet = workbook.SheetNames[0];
  const rows = XLSX.utils.sheet_to_json(workbook.Sheets[firstSheet], {
    defval: "",
    raw: false,
    blankrows: false,
  });

  return rows
    .map(mapRecord)
    .filter((record) => record.postingDate && Number.isFinite(record.amount) && record.amount !== 0);
}

function base64ToArrayBuffer(base64) {
  const binary = window.atob(base64);
  const length = binary.length;
  const bytes = new Uint8Array(length);

  for (let index = 0; index < length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }

  return bytes.buffer;
}

function mapRecord(row) {
  const postingDateRaw = getByAlias(row, ["Posting Date", "PostingDate", "Date"]);
  const customerNameRaw = getByAlias(row, [
    "Customer/Vendor Name",
    "Customer/Vendor Name (Insurance Provider)",
    "Customer Name",
    "Insurance Provider",
  ]);
  const patientNameRaw = getByAlias(row, ["Patient Name", "Patient"]);
  const itemNoRaw = getByAlias(row, [
    "Item No.",
    "Item No",
    "Item Number",
    "Item #",
    "Item",
    "Service Code",
    "Item/Service No",
  ]);
  const procedureRaw = getByAlias(row, ["Item/Service Description", "Item Description", "Procedure"]);
  const modalityRaw = getByAlias(row, ["Modality"]);
  const rowTotalRaw = getByAlias(row, ["Row Total", "Row Total (Amount Billed)", "Amount", "Billed Amount"]);
  const invoiceRaw = getByAlias(row, ["Invoice Number", "Invoice", "Invoice No"]);

  const postingDate = parseDateFlexible(postingDateRaw);
  const amount = parseAmount(rowTotalRaw);
  const customerName = cleanText(customerNameRaw);
  const patientName = cleanText(patientNameRaw);
  const itemNo = cleanText(itemNoRaw);
  const procedureName = cleanText(procedureRaw);
  const modality = detectModality(procedureName, cleanText(modalityRaw));
  const invoiceNumber = cleanText(invoiceRaw);

  const customerNorm = normalizeText(customerName);
  const patientNorm = normalizeText(patientName);
  const patientLoose = buildPatientLooseKey(patientName);
  const patientFirstSecond = buildPatientFirstSecondKey(patientName);
  const patientFirstLast = buildPatientFirstLastKey(patientName);
  const patientFirstThree = buildPatientFirstThreeKey(patientName);
  const itemNoNorm = normalizeKey(itemNo);
  const procedureNorm = normalizeProcedure(procedureName);
  const dayKey = postingDate ? toDateOnly(postingDate) : "";
  const weekInfo = postingDate ? getWeekInfo(postingDate) : null;

  return {
    postingDate,
    dayKey,
    customerName,
    customerNorm,
    patientName,
    patientNorm,
    patientLoose,
    patientFirstSecond,
    patientFirstLast,
    patientFirstThree,
    itemNo,
    itemNoNorm,
    procedureName,
    procedureNorm,
    modality,
    amount,
    invoiceNumber,
    weekKey: weekInfo ? weekInfo.weekKey : "Unknown",
    weekLabel: weekInfo ? weekInfo.weekLabel : "Unknown",
    isSha: customerNorm.includes(SHA_NAME),
    isTopup: customerNorm.includes(TOPUP_NAME),
  };
}

function getByAlias(row, aliases) {
  const keys = Object.keys(row);
  for (const alias of aliases) {
    if (row[alias] !== undefined) {
      return row[alias];
    }

    const keyMatch = keys.find((key) => normalizeKey(key) === normalizeKey(alias));
    if (keyMatch) {
      return row[keyMatch];
    }
  }

  return "";
}

function normalizeKey(key) {
  return String(key || "").toLowerCase().replace(/[^a-z0-9]/g, "");
}

function cleanText(value) {
  return String(value || "").trim();
}

function normalizeText(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/[^a-z0-9\s]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function normalizeProcedure(value) {
  return normalizeText(value)
    .replace(/\bwith contrast\b/g, "")
    .replace(/\bwithout contrast\b/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

function parseDateFlexible(value) {
  if (value instanceof Date && !Number.isNaN(value.valueOf())) {
    return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  }

  if (typeof value === "number" && Number.isFinite(value)) {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (!parsed) {
      return null;
    }
    return new Date(parsed.y, parsed.m - 1, parsed.d);
  }

  if (!value) {
    return null;
  }

  const text = String(value).trim();

  // Source spreadsheets render ambiguous NN-NN-YY / NN/NN/YY dates in
  // month-first order (matching Excel's "mm-dd-yy" cell format used
  // throughout this data), so try that interpretation first.
  const numericDate = text.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{2,4})$/);
  if (numericDate) {
    const year = Number(numericDate[3].length === 2 ? `20${numericDate[3]}` : numericDate[3]);

    const tryBuild = (month, day) => {
      const parsed = new Date(year, month, day);
      if (
        !Number.isNaN(parsed.valueOf()) &&
        parsed.getFullYear() === year &&
        parsed.getMonth() === month &&
        parsed.getDate() === day
      ) {
        return parsed;
      }
      return null;
    };

    // Month-first (e.g. 07-01-26 -> July 1)
    const monthFirst = tryBuild(Number(numericDate[1]) - 1, Number(numericDate[2]));
    if (monthFirst) {
      return monthFirst;
    }
    // Fall back to day-first for genuinely day-first sources where the
    // month-first reading was invalid (e.g. 13-06-26 -> 13 June).
    const dayFirst = tryBuild(Number(numericDate[2]) - 1, Number(numericDate[1]));
    if (dayFirst) {
      return dayFirst;
    }
  }

  const direct = new Date(text);
  if (!Number.isNaN(direct.valueOf())) {
    return new Date(direct.getFullYear(), direct.getMonth(), direct.getDate());
  }

  return null;
}

function parseAmount(value) {
  if (typeof value === "number" && Number.isFinite(value)) {
    return value;
  }

  const cleaned = String(value || "")
    .replace(/[^\d.-]/g, "")
    .trim();

  const num = Number(cleaned);
  return Number.isFinite(num) ? num : 0;
}

function detectModality(procedureName, modalityRaw) {
  const text = `${procedureName || ""} ${modalityRaw || ""}`
    .toLowerCase()
    .replace(/[.]/g, "")
    .replace(/\s+/g, " ")
    .trim();

  if (/\bmri\b|magnetic resonance/.test(text)) {
    return "MRI";
  }

  if (/\bct\b|computed tomography|cat scan/.test(text)) {
    return "CT";
  }

  if (/ultra\s*sound|sonograph|\bus\b/.test(text)) {
    return "Ultrasound";
  }

  if (/x\s*-?\s*ray|radiograph|\bcxr\b/.test(text)) {
    return "X-Ray";
  }

  if (/sedation|anaesthesia|anesthesia/.test(text)) {
    return "Sedation";
  }

  return "Other";
}

function getWeekInfo(date) {
  const weekStart = new Date(date.getFullYear(), date.getMonth(), date.getDate() - date.getDay());
  const weekEnd = new Date(weekStart.getFullYear(), weekStart.getMonth(), weekStart.getDate() + 6);
  const weekKey = toDateOnly(weekStart);
  const fmt = (d) => d.toLocaleDateString("en-GB", { day: "numeric", month: "short" });
  const weekLabel = `${fmt(weekStart)} – ${fmt(weekEnd)}`;
  return {
    weekKey,
    weekLabel,
  };
}

function toDateOnly(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function getDateBounds(records) {
  const dates = records.map((x) => x.postingDate).filter(Boolean).sort((a, b) => a - b);
  return {
    minDate: dates[0],
    maxDate: dates[dates.length - 1],
  };
}

function getCurrentMonthBounds(referenceDate = new Date()) {
  const year = referenceDate.getFullYear();
  const month = referenceDate.getMonth();
  return {
    fromDate: new Date(year, month, 1),
    toDate: new Date(year, month + 1, 0),
  };
}

function applyDateBounds(bounds) {
  const min = toDateOnly(bounds.minDate);
  const max = toDateOnly(bounds.maxDate);

  refs.fromDate.min = min;
  refs.fromDate.max = max;
  refs.toDate.min = min;
  refs.toDate.max = max;
}

function setDateInputs(fromDate, toDate) {
  refs.fromDate.value = toDateOnly(fromDate);
  refs.toDate.value = toDateOnly(toDate);
}

function applyDateRange() {
  refreshDashboard();
}

function parseDateInputValue(value, endOfDay = false) {
  if (!value) {
    return null;
  }

  const match = String(value).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (match) {
    const year = Number(match[1]);
    const month = Number(match[2]) - 1;
    const day = Number(match[3]);
    return endOfDay
      ? new Date(year, month, day, 23, 59, 59, 999)
      : new Date(year, month, day);
  }

  const parsed = parseDateFlexible(value);
  if (!parsed) {
    return null;
  }

  if (endOfDay) {
    parsed.setHours(23, 59, 59, 999);
  }

  return parsed;
}

function refreshDashboard() {
  if (!state.allRecords.length) {
    return;
  }

  state.filteredRecords = applyDateFilter(state.allRecords, refs.fromDate.value, refs.toDate.value);
  renderReportingPeriod();
  renderDataAsAt();

  const derived = buildDerivedData(state.filteredRecords, refs.fromDate.value, refs.toDate.value);

  renderKPIs(derived.kpis);
  renderExecutiveModalityCards(derived.modalityStats);
  renderInsuranceTable(derived.insuranceStats);
  renderProcedureTable(derived.procedureStats);
  renderExecutiveSummary(derived);
  renderCharts(derived);
}

function applyDateFilter(records, fromDate, toDate) {
  const from = parseDateInputValue(fromDate);
  const to = parseDateInputValue(toDate, true);

  return records.filter((record) => {
    const time = record.postingDate.getTime();
    const fromOk = from ? time >= from.getTime() : true;
    const toOk = to ? time <= to.getTime() : true;
    return fromOk && toOk;
  });
}

function buildDerivedData(records, fromValue, toValue) {
  // This dashboard now expects a curated, pre-verified data source: every
  // row is already known to be a genuine SHA claim, cash top-up, or
  // insurance co-pay — there is no ambiguity left to resolve by matching
  // patients/exams to each other. Every figure below is a direct sum over
  // the already-correct rows for the selected date range.
  const shaRows = records.filter((r) => r.isSha);
  const cashRows = records.filter((r) => !r.isSha && r.isTopup);
  const insuranceRows = records.filter((r) => !r.isSha && !r.isTopup);

  const kpis = computeKpis(shaRows, cashRows, insuranceRows);
  const modalityStats = computeModalityStats(shaRows, cashRows, insuranceRows, kpis.totalRevenueGenerated);
  const vendorRecordsInRange = applyDateFilter(state.vendorRecords, fromValue, toValue);
  const insuranceStats = computeInsuranceStats(vendorRecordsInRange);
  const procedureStats = computeProcedureStats(shaRows, cashRows, insuranceRows);
  const weeklyStats = computeWeeklyStats(shaRows, cashRows, insuranceRows);

  return {
    kpis,
    modalityStats,
    insuranceStats,
    procedureStats,
    weeklyStats,
    reconciliation: null,
  };
}

function sumAmount(rows) {
  return rows.reduce((sum, r) => sum + r.amount, 0);
}

function computeKpis(shaRows, cashRows, insuranceRows) {
  const totalShaBilling = sumAmount(shaRows);
  const matchedCashTopup = sumAmount(cashRows);
  const insuranceTopupTotal = sumAmount(insuranceRows);
  const totalPatientTopup = matchedCashTopup;
  const totalTopupCombined = matchedCashTopup + insuranceTopupTotal;
  const totalRevenueGenerated = totalShaBilling + totalTopupCombined;
  const totalShaStudies = shaRows.filter((r) => r.amount > 0).length;

  const avgShaPerStudy = totalShaStudies ? totalShaBilling / totalShaStudies : 0;
  const avgTopupPerStudy = totalShaStudies ? totalPatientTopup / totalShaStudies : 0;

  const shaCoveragePct = totalRevenueGenerated ? (totalShaBilling / totalRevenueGenerated) * 100 : 0;
  const patientContributionPct = totalRevenueGenerated ? (totalTopupCombined / totalRevenueGenerated) * 100 : 0;

  return {
    totalShaBilling,
    matchedCashTopup,
    insuranceTopupTotal,
    totalPatientTopup,
    totalTopupCombined,
    totalRevenueGenerated,
    totalShaStudies,
    avgShaPerStudy,
    avgTopupPerStudy,
    shaCoveragePct,
    patientContributionPct,
  };
}

function computeModalityStats(shaRows, cashRows, insuranceRows, totalRevenueGenerated) {
  const seed = MODALITIES.map((modality) => ({
    modality,
    studies: 0,
    shaBilling: 0,
    patientTopup: 0,
    insuranceCopay: 0,
    patientCashTopup: 0,
    totalRevenue: 0,
    avgSha: 0,
    avgTopup: 0,
    avgInsuranceCopay: 0,
    sharePct: 0,
  }));
  const map = Object.fromEntries(seed.map((row) => [row.modality, row]));

  shaRows.forEach((r) => {
    const bucket = map[r.modality] || map.Other;
    if (r.amount > 0) bucket.studies += 1;
    bucket.shaBilling += r.amount;
  });
  cashRows.forEach((r) => {
    const bucket = map[r.modality] || map.Other;
    bucket.patientCashTopup += r.amount;
  });
  insuranceRows.forEach((r) => {
    const bucket = map[r.modality] || map.Other;
    bucket.insuranceCopay += r.amount;
  });

  seed.forEach((row) => {
    row.patientTopup = row.patientCashTopup + row.insuranceCopay;
    row.totalRevenue = row.shaBilling + row.patientTopup;
    row.avgSha = row.studies ? row.shaBilling / row.studies : 0;
    row.avgTopup = row.studies ? row.patientCashTopup / row.studies : 0;
    row.avgInsuranceCopay = row.studies ? row.insuranceCopay / row.studies : 0;
    row.sharePct = totalRevenueGenerated ? (row.totalRevenue / totalRevenueGenerated) * 100 : 0;
  });

  return seed;
}

function computeInsuranceStats(vendorRecords) {
  const map = new Map();

  const ensure = (key, providerName) => {
    if (!map.has(key)) {
      map.set(key, {
        provider: providerName,
        insuranceBilling: 0,
        totalRevenue: 0,
        studies: 0,
        avgBilling: 0,
        sharePct: 0,
      });
    }
    return map.get(key);
  };

  vendorRecords.forEach((r) => {
    const key = r.customerNorm || r.customerName || "unknown";
    const row = ensure(key, r.customerName || "Unknown Provider");
    row.insuranceBilling += r.amount;
    if (r.amount > 0) row.studies += 1;
  });

  const grandTotal = [...map.values()].reduce((sum, r) => sum + r.insuranceBilling, 0) || 1;

  map.forEach((row) => {
    row.totalRevenue = row.insuranceBilling;
    row.avgBilling = row.studies ? row.insuranceBilling / row.studies : 0;
    row.sharePct = (row.totalRevenue / grandTotal) * 100;
  });

  return [...map.values()];
}

function computeProcedureStats(shaRows, cashRows, insuranceRows) {
  const grouped = new Map();

  const ensure = (name) => {
    const key = name.toLowerCase();
    if (!grouped.has(key)) {
      grouped.set(key, { procedureName: name, studies: 0, shaBilling: 0, patientTopup: 0, totalRevenue: 0 });
    }
    return grouped.get(key);
  };

  shaRows.forEach((r) => {
    const item = ensure(r.procedureName);
    if (r.amount > 0) item.studies += 1;
    item.shaBilling += r.amount;
    item.totalRevenue += r.amount;
  });
  cashRows.forEach((r) => {
    const item = ensure(r.procedureName);
    item.patientTopup += r.amount;
    item.totalRevenue += r.amount;
  });
  insuranceRows.forEach((r) => {
    const item = ensure(r.procedureName);
    item.patientTopup += r.amount;
    item.totalRevenue += r.amount;
  });

  return [...grouped.values()].sort((a, b) => b.totalRevenue - a.totalRevenue);
}

function computeWeeklyStats(shaRows, cashRows, insuranceRows) {
  const map = new Map();

  const ensure = (r) => {
    if (!map.has(r.weekKey)) {
      map.set(r.weekKey, { weekKey: r.weekKey, weekLabel: r.weekLabel, shaBilling: 0, patientTopup: 0, totalRevenue: 0, studies: 0 });
    }
    return map.get(r.weekKey);
  };

  shaRows.forEach((r) => {
    const bucket = ensure(r);
    bucket.shaBilling += r.amount;
    bucket.totalRevenue += r.amount;
    if (r.amount > 0) bucket.studies += 1;
  });
  cashRows.forEach((r) => {
    const bucket = ensure(r);
    bucket.patientTopup += r.amount;
    bucket.totalRevenue += r.amount;
  });
  insuranceRows.forEach((r) => {
    const bucket = ensure(r);
    bucket.patientTopup += r.amount;
    bucket.totalRevenue += r.amount;
  });

  return [...map.values()].sort((a, b) => a.weekKey.localeCompare(b.weekKey));
}

function renderKPIs(kpis) {
  if (refs.kpiShaBilling) refs.kpiShaBilling.textContent = formatCurrency(kpis.totalShaBilling);
  if (refs.kpiTopup) refs.kpiTopup.textContent = formatCurrency(kpis.totalPatientTopup);
  if (refs.kpiInsuranceTopup) refs.kpiInsuranceTopup.textContent = formatCurrency(kpis.insuranceTopupTotal || 0);
  if (refs.kpiTotalRevenue) refs.kpiTotalRevenue.textContent = formatCurrency(kpis.totalRevenueGenerated);
  if (refs.kpiStudies) refs.kpiStudies.textContent = formatNumber(kpis.totalShaStudies);
  if (refs.kpiCoverage) refs.kpiCoverage.textContent = `${kpis.shaCoveragePct.toFixed(2)}%`;
  if (refs.kpiPatientPct) refs.kpiPatientPct.textContent = `${kpis.patientContributionPct.toFixed(2)}%`;
}

function renderExecutiveModalityCards(modalityStats) {
  const mri = modalityStats.find((item) => item.modality === "MRI") || emptyModalityRow("MRI");
  const ct = modalityStats.find((item) => item.modality === "CT") || emptyModalityRow("CT");

  fillModalityCard("mri", mri);
  fillModalityCard("ct", ct);
}

function renderInsuranceTable(insuranceStats) {
  const sorted = [...insuranceStats].sort((a, b) => b.totalRevenue - a.totalRevenue);

  if (!insuranceStats.length) {
    if (refs.insuranceTableBody) {
      refs.insuranceTableBody.innerHTML = '<tr><td colspan="5" class="text-center py-4">No insurance records available.</td></tr>';
    }
    return;
  }

  const isShaProvider = (name) => String(name || "").toUpperCase().includes(SHA_NAME.toUpperCase());

  const bodyRows = sorted
    .map((item) => {
      const rowClass = isShaProvider(item.provider) ? "sha-row" : "";
      return `<tr class="${rowClass}">
        <td>${escapeHtml(item.provider)}${isShaProvider(item.provider) ? " <strong>(SHA)</strong>" : ""}</td>
        <td class="text-end">${formatCurrency(item.totalRevenue)}</td>
        <td class="text-end">${formatNumber(item.studies)}</td>
        <td class="text-end">${formatCurrency(item.avgBilling)}</td>
        <td class="text-end">${item.sharePct.toFixed(2)}%</td>
      </tr>`;
    })
    .join("");

  const totals = sorted.reduce(
    (acc, row) => ({
      totalRevenue: acc.totalRevenue + (Number(row.totalRevenue) || 0),
      studies: acc.studies + (Number(row.studies) || 0),
    }),
    { totalRevenue: 0, studies: 0 }
  );

  const totalAvg = totals.studies ? totals.totalRevenue / totals.studies : 0;
  const totalRow = `<tr class="totals-row">
    <td><strong>TOTAL</strong></td>
    <td class="text-end"><strong>${formatCurrency(totals.totalRevenue)}</strong></td>
    <td class="text-end"><strong>${formatNumber(totals.studies)}</strong></td>
    <td class="text-end"><strong>${formatCurrency(totalAvg)}</strong></td>
    <td class="text-end"><strong>100.00%</strong></td>
  </tr>`;

  refs.insuranceTableBody.innerHTML = `${bodyRows}${totalRow}`;
}

function renderProcedureTable(procedureStats) {
  if (!procedureStats.length) {
    refs.procedureTableBody.innerHTML = '<tr><td colspan="5" class="text-center py-4">No procedure records available.</td></tr>';
    return;
  }

  refs.procedureTableBody.innerHTML = procedureStats
    .slice(0, 10)
    .map(
      (item) =>
        `<tr>
          <td>${escapeHtml(item.procedureName)}</td>
          <td class="text-end">${formatNumber(item.studies)}</td>
          <td class="text-end">${formatCurrency(item.shaBilling)}</td>
          <td class="text-end">${formatCurrency(item.patientTopup)}</td>
          <td class="text-end">${formatCurrency(item.totalRevenue)}</td>
        </tr>`
    )
    .join("");
}

function renderReconciliation(reconciliation) {
  if (!refs.recoSourceTotal) {
    return;
  }

  if (!reconciliation || !reconciliation.totals) {
    return;
  }

  const { totals } = reconciliation;
  const otherProviders =
    totals.totalSourceAmount - totals.totalShaAmount - totals.totalCashAmount;

  if (refs.recoSourceTotal) refs.recoSourceTotal.textContent = formatCurrency(totals.totalSourceAmount);
  if (refs.recoShaTotal) refs.recoShaTotal.textContent = formatCurrency(totals.totalShaAmount);
  if (refs.recoNbiTotal) refs.recoNbiTotal.textContent = formatCurrency(totals.totalCashAmount);
  if (refs.recoMatchedTotal) refs.recoMatchedTotal.textContent = formatCurrency(totals.matchedCashAmount);
  if (refs.recoUnmatchedShaTotal) refs.recoUnmatchedShaTotal.textContent = formatCurrency(totals.unmatchedShaAmount);
  if (refs.recoUnmatchedNbiTotal) refs.recoUnmatchedNbiTotal.textContent = formatCurrency(totals.unmatchedCashAmount);
  if (refs.recoOtherTotal) refs.recoOtherTotal.textContent = formatCurrency(otherProviders);
  if (refs.recoMatchedCount) refs.recoMatchedCount.textContent = formatNumber(totals.matchedPairCount);

  if (refs.matchedPairsBody) {
    const rows = reconciliation.matchedPairs || [];
    refs.matchedPairsBody.innerHTML = rows.length
      ? rows
          .map(
            (pair) => `<tr>
              <td>${escapeHtml(pair.sha.patientName)}</td>
              <td>${escapeHtml(pair.sha.dayKey)}</td>
              <td>${escapeHtml(pair.sha.itemNo || "")}</td>
              <td>${escapeHtml(pair.sha.procedureName)}</td>
              <td>${escapeHtml(pair.sha.modality)}</td>
              <td class="text-end">${formatCurrency(pair.sha.amount)}</td>
              <td class="text-end">${formatCurrency(pair.cash.amount)}</td>
              <td>${escapeHtml(pair.rule)}</td>
            </tr>`
          )
          .join("")
      : '<tr><td colspan="8" class="text-center py-3">No matched SHA + cash pairs.</td></tr>';
  }

  if (refs.unmatchedShaBody) {
    const rows = reconciliation.unmatchedShaRecords || [];
    refs.unmatchedShaBody.innerHTML = rows.length
      ? rows
          .map(
            (record) => `<tr>
              <td>${escapeHtml(record.patientName)}</td>
              <td>${escapeHtml(record.dayKey)}</td>
              <td>${escapeHtml(record.itemNo || "")}</td>
              <td>${escapeHtml(record.procedureName)}</td>
              <td>${escapeHtml(record.modality)}</td>
              <td class="text-end">${formatCurrency(record.amount)}</td>
              <td>${escapeHtml(record.matchReason || "No unique non-SHA match")}</td>
            </tr>`
          )
          .join("")
      : '<tr><td colspan="7" class="text-center py-3">No unmatched SHA records.</td></tr>';
  }

  if (refs.unmatchedCashBody) {
    const rows = reconciliation.unmatchedCashRecords || [];
    refs.unmatchedCashBody.innerHTML = rows.length
      ? rows
          .map(
            (record) => `<tr>
              <td>${escapeHtml(record.patientName)}</td>
              <td>${escapeHtml(record.dayKey)}</td>
              <td>${escapeHtml(record.itemNo || "")}</td>
              <td>${escapeHtml(record.procedureName)}</td>
              <td>${escapeHtml(record.modality)}</td>
              <td class="text-end">${formatCurrency(record.amount)}</td>
            </tr>`
          )
          .join("")
      : '<tr><td colspan="6" class="text-center py-3">No unmatched non-SHA candidate records.</td></tr>';
  }
}

function renderCharts(derived) {
  renderWeeklyChart(derived.weeklyStats);
  renderInsuranceShareChart(derived.insuranceStats);
}

function renderReportingPeriod() {
  const from = refs.fromDate?.value ? parseDateInputValue(refs.fromDate.value) : null;
  const to = refs.toDate?.value ? parseDateInputValue(refs.toDate.value) : null;

  const text = from && to
    ? `${formatDateForDisplay(from)} to ${formatDateForDisplay(to)}`
    : "-";

  if (refs.reportingPeriod) {
    refs.reportingPeriod.textContent = text;
  }
  if (refs.summaryPeriod) {
    refs.summaryPeriod.textContent = `Period: ${text}`;
  }
}

function renderDataAsAt() {
  if (!refs.dataAsAt) {
    return;
  }

  const to = refs.toDate?.value ? parseDateInputValue(refs.toDate.value) : null;
  const dateText = to ? formatDateForDisplay(to) : "-";
  refs.dataAsAt.textContent = `Data as at: ${dateText} | Powered by Business Intelligence`;
}

function formatDateForDisplay(date) {
  return date.toLocaleDateString("en-KE", {
    day: "2-digit",
    month: "short",
    year: "numeric",
  });
}

function emptyModalityRow(modality) {
  return {
    modality,
    studies: 0,
    shaBilling: 0,
    patientTopup: 0,
    patientCashTopup: 0,
    insuranceCopay: 0,
    totalRevenue: 0,
    avgSha: 0,
    avgTopup: 0,
    avgPatientTopup: 0,
    avgInsuranceCopay: 0,
    sharePct: 0,
  };
}

function fillModalityCard(prefix, row) {
  if (refs[`${prefix}Studies`]) refs[`${prefix}Studies`].textContent = formatNumber(row.studies);
  if (refs[`${prefix}Sha`]) refs[`${prefix}Sha`].textContent = formatCurrency(row.shaBilling);
  if (refs[`${prefix}Copay`]) refs[`${prefix}Copay`].textContent = formatCurrency(row.insuranceCopay);
  if (refs[`${prefix}Topup`]) refs[`${prefix}Topup`].textContent = formatCurrency(row.patientCashTopup);
  if (refs[`${prefix}Revenue`]) refs[`${prefix}Revenue`].textContent = formatCurrency(row.totalRevenue);
  if (refs[`${prefix}AvgSha`]) refs[`${prefix}AvgSha`].textContent = formatCurrency(row.avgSha);
  if (refs[`${prefix}AvgTopup`]) refs[`${prefix}AvgTopup`].textContent = formatCurrency(row.avgPatientTopup);
  if (refs[`${prefix}AvgCopay`]) refs[`${prefix}AvgCopay`].textContent = formatCurrency(row.avgInsuranceCopay);
  if (refs[`${prefix}Share`]) refs[`${prefix}Share`].textContent = `${row.sharePct.toFixed(2)}%`;
}

function renderExecutiveSummary(derived) {
  if (!refs.executiveSummaryList) {
    return;
  }

  const { kpis, modalityStats, insuranceStats, procedureStats, reconciliation } = derived;
  const mri = modalityStats.find((item) => item.modality === "MRI") || emptyModalityRow("MRI");
  const ct = modalityStats.find((item) => item.modality === "CT") || emptyModalityRow("CT");
  const totalInsuranceCopay = (reconciliation?.matchedPairs || []).reduce((sum, pair) => {
    const customerNorm = String(pair?.cash?.customerNorm || "");
    if (!customerNorm.includes(TOPUP_NAME) && !customerNorm.includes(SHA_NAME)) {
      return sum + (Number(pair?.cash?.amount || 0) || 0);
    }
    return sum;
  }, 0);
  const topInsurer = [...insuranceStats].sort((a, b) => b.totalRevenue - a.totalRevenue)[0];
  const topProcedure = procedureStats[0];

  const items = [
    `Total SHA revenue generated in the selected period was ${formatCurrency(kpis.totalRevenueGenerated)} from ${formatNumber(kpis.totalShaStudies)} SHA studies.`,
    `Total patient top-up (cash) collected was ${formatCurrency(kpis.totalPatientTopup)}, while insurance co-pay/top-up was ${formatCurrency(kpis.insuranceTopupTotal)} (reconciliation co-pay total: ${formatCurrency(totalInsuranceCopay)}).`,
    `MRI contributed ${formatCurrency(mri.totalRevenue)} (${mri.sharePct.toFixed(2)}% of total revenue), while CT contributed ${formatCurrency(ct.totalRevenue)} (${ct.sharePct.toFixed(2)}%).`,
    topInsurer
      ? `${escapeHtml(topInsurer.provider)} was the highest contributing insurer with total revenue of ${formatCurrency(topInsurer.totalRevenue)}.`
      : "No insurance provider records were available for this period.",
    topProcedure
      ? `The highest performing procedure was ${escapeHtml(topProcedure.procedureName)} with total revenue of ${formatCurrency(topProcedure.totalRevenue)}.`
      : "No procedure-level records were available for this period.",
  ];

  refs.executiveSummaryList.innerHTML = items
    .map((line) => `<li class="insight-row"><span class="insight-text">${line}</span></li>`)
    .join("");
}

function renderWeeklyChart(weeklyStats) {
  const labels = weeklyStats.map((x) => x.weekLabel);

  state.charts.weekly = upsertChart(state.charts.weekly, "weeklyChart", {
    type: "bar",
    data: {
      labels,
      datasets: [
        {
          type: "bar",
          label: "SHA Billing",
          data: weeklyStats.map((x) => round2(x.shaBilling)),
          backgroundColor: chartColors.sha,
          borderRadius: 6,
          stack: "revenue",
        },
        {
          type: "bar",
          label: "Patient Top-Up",
          data: weeklyStats.map((x) => round2(x.patientTopup)),
          backgroundColor: chartColors.topup,
          borderRadius: 6,
          stack: "revenue",
        },
        {
          type: "line",
          label: "Total Revenue",
          data: weeklyStats.map((x) => round2(x.totalRevenue)),
          borderColor: chartColors.total,
          backgroundColor: chartColors.total,
          pointRadius: 3,
          pointHoverRadius: 4,
          borderWidth: 2.4,
          tension: 0.3,
          yAxisID: "y",
        },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      interaction: { mode: "index", intersect: false },
      scales: {
        y: {
          beginAtZero: true,
          ticks: { callback: (value) => shortCurrency(value) },
          grid: { color: "rgba(18, 37, 60, 0.05)" },
        },
        x: {
          grid: { display: false },
        },
      },
      plugins: {
        legend: { position: "bottom", labels: { usePointStyle: true, boxWidth: 8, boxHeight: 8, padding: 14, font: { size: 12 } } },
        tooltip: {
          mode: "index",
          intersect: false,
          callbacks: {
            label: (ctx) => ` ${ctx.dataset.label}: ${formatCurrency(ctx.parsed.y)}`,
          },
        },
      },
    },
  });
}

function renderModalityChart(modalityStats) {
  const focusModalities = ["MRI", "CT", "Ultrasound", "X-Ray"];
  const chartRows = modalityStats.filter((x) => focusModalities.includes(x.modality));
  const labels = chartRows.map((x) => x.modality);

  state.charts.modality = upsertChart(state.charts.modality, "modalityChart", {
    type: "bar",
    data: {
      labels,
      datasets: [
        {
          label: "SHA Billing",
          data: chartRows.map((x) => round2(x.shaBilling)),
          backgroundColor: chartColors.sha,
          stack: "modality",
          borderRadius: 6,
        },
        {
          label: "Patient Top-Up",
          data: chartRows.map((x) => round2(x.patientTopup)),
          backgroundColor: chartColors.topup,
          stack: "modality",
          borderRadius: 6,
        },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      scales: {
        y: {
          beginAtZero: true,
          ticks: { callback: (value) => shortCurrency(value) },
          grid: { color: "rgba(18, 37, 60, 0.05)" },
        },
        x: {
          grid: { display: false },
        },
      },
      plugins: {
        legend: { position: "bottom" },
        tooltip: {
          callbacks: {
            footer: (items) => {
              const item = items[0];
              const idx = item.dataIndex;
              return `Total Revenue: ${formatCurrency(modalityStats[idx].totalRevenue)}`;
            },
          },
        },
      },
    },
  });
}

function renderInsuranceShareChart(insuranceStats) {
  const topProviders = [...insuranceStats].sort((a, b) => b.totalRevenue - a.totalRevenue).slice(0, 8);

  state.charts.insuranceShare = upsertChart(state.charts.insuranceShare, "insuranceShareChart", {
    type: "doughnut",
    data: {
      labels: topProviders.map((x) => x.provider),
      datasets: [
        {
          data: topProviders.map((x) => round2(x.totalRevenue)),
          backgroundColor: [
            "#0b5ba0",
            "#2f7dc4",
            "#5c9cd6",
            "#8bbce6",
            "#12805c",
            "#4aa483",
            "#94a3b8",
            "#c3cbd6",
          ],
          borderColor: "#ffffff",
          borderWidth: 2,
        },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      cutout: "68%",
      plugins: {
        legend: {
          position: "bottom",
          labels: { usePointStyle: true, boxWidth: 8, boxHeight: 8, padding: 14, font: { size: 12 } },
        },
        tooltip: {
          callbacks: {
            label: (ctx) => {
              const provider = insuranceStats.find((x) => x.provider === ctx.label);
              const share = provider ? provider.sharePct.toFixed(2) : "0.00";
              return ` ${ctx.label}: ${formatCurrency(ctx.parsed)} (${share}%)`;
            },
          },
        },
      },
    },
  });
}

function upsertChart(current, canvasId, config) {
  if (current) {
    current.destroy();
  }
  return new Chart(document.getElementById(canvasId), config);
}

function clearDashboard() {
  renderKPIs({
    totalShaBilling: 0,
    matchedCashTopup: 0,
    insuranceTopupTotal: 0,
    totalPatientTopup: 0,
    totalTopupCombined: 0,
    totalRevenueGenerated: 0,
    totalShaStudies: 0,
    avgShaPerStudy: 0,
    avgTopupPerStudy: 0,
    shaCoveragePct: 0,
    patientContributionPct: 0,
  });
  renderExecutiveModalityCards([emptyModalityRow("MRI"), emptyModalityRow("CT")], { matchedPairs: [] });
  if (refs.reportingPeriod) refs.reportingPeriod.textContent = "-";
  if (refs.dataAsAt) refs.dataAsAt.textContent = "Data as at: - | Powered by Business Intelligence";
  if (refs.summaryPeriod) refs.summaryPeriod.textContent = "Period: -";
  if (refs.executiveSummaryList) {
    refs.executiveSummaryList.innerHTML = "<li>No data available for the selected period.</li>";
  }

  if (refs.insuranceTableBody) {
    refs.insuranceTableBody.innerHTML = '<tr><td colspan="8" class="text-center py-4">No data available.</td></tr>';
  }
  if (refs.procedureTableBody) {
    refs.procedureTableBody.innerHTML = '<tr><td colspan="5" class="text-center py-4">No data available.</td></tr>';
  }

  Object.keys(state.charts).forEach((key) => {
    if (state.charts[key]) {
      state.charts[key].destroy();
      state.charts[key] = null;
    }
  });
}

function formatCurrency(value) {
  return new Intl.NumberFormat("en-KE", {
    style: "currency",
    currency: "KES",
    maximumFractionDigits: 2,
  }).format(Number(value || 0));
}

function getCanonicalProviderKey(customerNorm) {
  const norm = customerNorm || "unknown provider";
  if (norm.includes(SHA_NAME)) {
    return SHA_NAME;
  }
  if (norm.includes(TOPUP_NAME)) {
    return TOPUP_NAME;
  }
  return norm;
}

function selectTopupCandidates(studies, group) {
  let candidates = studies.filter(
    (study) =>
      isSamePatient(study, group) &&
      study.dayKey === group.dayKey &&
      study.modality === group.modality
  );

  if (candidates.length) {
    return candidates;
  }

  const sameDayStudies = studies.filter(
    (study) => isSamePatient(study, group) && study.dayKey === group.dayKey
  );
  const uniqueModalities = new Set(sameDayStudies.map((study) => study.modality));
  if (sameDayStudies.length && uniqueModalities.size === 1) {
    return sameDayStudies;
  }

  const nearDateSameModality = studies.filter(
    (study) =>
      isSamePatient(study, group) &&
      study.modality === group.modality &&
      daysBetween(study.postingDate, group.postingDate) <= 2
  );

  return nearDateSameModality;
}

function daysBetween(a, b) {
  if (!a || !b) {
    return Number.POSITIVE_INFINITY;
  }

  const ms = Math.abs(a.getTime() - b.getTime());
  return ms / 86400000;
}

function buildPatientLooseKey(name) {
  const norm = normalizeText(name);
  if (!norm) {
    return "";
  }

  const parts = norm.split(" ").filter(Boolean);
  if (parts.length <= 1) {
    return norm;
  }

  return `${parts[0]} ${parts[parts.length - 1]}`;
}

function buildPatientFirstSecondKey(name) {
  const norm = normalizeText(name);
  if (!norm) {
    return "";
  }

  const parts = norm.split(" ").filter(Boolean);
  if (parts.length <= 1) {
    return norm;
  }

  return `${parts[0]} ${parts[1]}`;
}

function buildPatientFirstLastKey(name) {
  return buildPatientLooseKey(name);
}

function buildPatientFirstThreeKey(name) {
  const norm = normalizeText(name);
  if (!norm) {
    return "";
  }

  const parts = norm.split(" ").filter(Boolean);
  if (parts.length < 3) {
    return "";
  }

  return `${parts[0]} ${parts[1]} ${parts[2]}`;
}

function isSamePatientAdvanced(sha, cash) {
  if (sha.patientNorm && cash.patientNorm && sha.patientNorm === cash.patientNorm) {
    return true;
  }

  const shaKeys = new Set(
    [sha.patientFirstThree, sha.patientFirstLast, sha.patientFirstSecond, sha.patientLoose]
      .filter(Boolean)
  );

  const cashKeys = [cash.patientFirstThree, cash.patientFirstLast, cash.patientFirstSecond, cash.patientLoose]
    .filter(Boolean);

  if (cashKeys.some((key) => shaKeys.has(key))) {
    return true;
  }

  return arePatientNamesFuzzyMatch(sha.patientNorm, cash.patientNorm);
}

function arePatientNamesFuzzyMatch(aNorm, bNorm) {
  if (!aNorm || !bNorm) {
    return false;
  }

  const aTokens = normalizeNameTokens(aNorm);
  const bTokens = normalizeNameTokens(bNorm);
  if (!aTokens.length || !bTokens.length) {
    return false;
  }

  const aSorted = [...aTokens].sort().join(" ");
  const bSorted = [...bTokens].sort().join(" ");
  if (aSorted === bSorted) {
    return true;
  }

  const usedB = new Set();
  let matches = 0;

  aTokens.forEach((aToken) => {
    for (let index = 0; index < bTokens.length; index += 1) {
      if (usedB.has(index)) {
        continue;
      }
      if (areNameTokensClose(aToken, bTokens[index])) {
        usedB.add(index);
        matches += 1;
        break;
      }
    }
  });

  const coverageA = matches / aTokens.length;
  const coverageB = matches / bTokens.length;
  return matches >= 2 && Math.min(coverageA, coverageB) >= 0.67;
}

function normalizeNameTokens(normName) {
  return String(normName || "")
    .split(" ")
    .map((token) => token.trim())
    .filter((token) => token.length >= 2);
}

function areNameTokensClose(a, b) {
  if (!a || !b) {
    return false;
  }
  if (a === b) {
    return true;
  }

  const maxLen = Math.max(a.length, b.length);
  if (maxLen < 5) {
    return false;
  }

  const distance = levenshteinDistance(a, b);
  return distance <= 1 || distance / maxLen <= 0.2;
}

function levenshteinDistance(a, b) {
  const left = String(a || "");
  const right = String(b || "");
  const rows = left.length + 1;
  const cols = right.length + 1;
  const dp = Array.from({ length: rows }, () => Array(cols).fill(0));

  for (let i = 0; i < rows; i += 1) {
    dp[i][0] = i;
  }
  for (let j = 0; j < cols; j += 1) {
    dp[0][j] = j;
  }

  for (let i = 1; i < rows; i += 1) {
    for (let j = 1; j < cols; j += 1) {
      const cost = left[i - 1] === right[j - 1] ? 0 : 1;
      dp[i][j] = Math.min(
        dp[i - 1][j] + 1,
        dp[i][j - 1] + 1,
        dp[i - 1][j - 1] + cost
      );
    }
  }

  return dp[left.length][right.length];
}

function isSamePatient(study, group) {
  if (study.patientNorm && group.patientNorm && study.patientNorm === group.patientNorm) {
    return true;
  }

  return Boolean(
    study.patientLoose &&
    group.patientLoose &&
    study.patientLoose === group.patientLoose
  );
}

function shortCurrency(value) {
  const num = Number(value || 0);
  if (Math.abs(num) >= 1000000) {
    return `KES ${(num / 1000000).toFixed(1)}M`;
  }
  if (Math.abs(num) >= 1000) {
    return `KES ${(num / 1000).toFixed(1)}K`;
  }
  return `KES ${num.toFixed(0)}`;
}

function formatNumber(value) {
  return Number(value || 0).toLocaleString("en-US");
}

function round2(value) {
  return Math.round((Number(value) + Number.EPSILON) * 100) / 100;
}

function escapeHtml(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

initialize();
