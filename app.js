const state = {
  allRecords: [],
  filteredRecords: [],
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
  sha: "#0b72c9",
  topup: "#4aa8ff",
  total: "#0a4578",
  mri: "#0b72c9",
  ct: "#2f84d5",
  us: "#5caef3",
  xray: "#7fc2ff",
  sedation: "#98abd0",
  other: "#bac6d4",
};

const refs = {
  fromDate: document.getElementById("fromDate"),
  toDate: document.getElementById("toDate"),
  resetDates: document.getElementById("resetDates"),

  kpiShaBilling: document.getElementById("kpiShaBilling"),
  kpiTopup: document.getElementById("kpiTopup"),
  kpiTotalRevenue: document.getElementById("kpiTotalRevenue"),
  kpiStudies: document.getElementById("kpiStudies"),
  kpiAvgSha: document.getElementById("kpiAvgSha"),
  kpiAvgTopup: document.getElementById("kpiAvgTopup"),
  kpiCoverage: document.getElementById("kpiCoverage"),
  kpiPatientPct: document.getElementById("kpiPatientPct"),

  modalityTableBody: document.querySelector("#modalityTable tbody"),
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
  refs.fromDate.addEventListener("change", refreshDashboard);
  refs.toDate.addEventListener("change", refreshDashboard);
  refs.resetDates.addEventListener("click", resetDateRange);
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
  // Reconcile SHA claims against cash/insurance top-ups across the FULL dataset,
  // once, so a claim and its matching top-up still link even when the top-up
  // was posted on a different date (sometimes days later, sometimes in the
  // following month). Narrowing the date picker must not break this matching.
  state.fullReconciliation = reconcileShaCashRecords(records);
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
  const clinicalRegions = [...extractBodyRegions(procedureNorm)];

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
    clinicalRegions,
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

  // Parse explicit day-first formats before native parsing to avoid MM/DD ambiguity.
  const dayFirst = text.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{2,4})$/);
  if (dayFirst) {
    const day = Number(dayFirst[1]);
    const month = Number(dayFirst[2]) - 1;
    const year = Number(dayFirst[3].length === 2 ? `20${dayFirst[3]}` : dayFirst[3]);
    const parsed = new Date(year, month, day);
    if (
      !Number.isNaN(parsed.valueOf()) &&
      parsed.getFullYear() === year &&
      parsed.getMonth() === month &&
      parsed.getDate() === day
    ) {
      return parsed;
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
  const year = date.getFullYear();
  const month = date.getMonth() + 1;
  const weekNum = Math.floor((date.getDate() - 1) / 7) + 1;
  const weekKey = `${year}-${String(month).padStart(2, "0")}-W${String(weekNum).padStart(2, "0")}`;
  return {
    weekKey,
    weekLabel: `Week ${weekNum}`,
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

function setDateInputs(minDate, maxDate) {
  const min = toDateOnly(minDate);
  const max = toDateOnly(maxDate);

  refs.fromDate.min = min;
  refs.fromDate.max = max;
  refs.toDate.min = min;
  refs.toDate.max = max;

  refs.fromDate.value = min;
  refs.toDate.value = max;
}

function resetDateRange() {
  if (!state.dateBounds) {
    return;
  }
  setDateInputs(state.dateBounds.minDate, state.dateBounds.maxDate);
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

  const derived = buildDerivedData(state.filteredRecords, refs.fromDate.value, refs.toDate.value);
  renderKPIs(derived.kpis);
  renderModalityTable(derived.modalityStats);
  renderInsuranceTable(derived.insuranceStats);
  renderProcedureTable(derived.procedureStats);
  renderReconciliation(derived.reconciliation);
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

function scopeReconciliationToDateRange(fullReconciliation, filteredRecords, fromValue, toValue) {
  const from = parseDateInputValue(fromValue);
  const to = parseDateInputValue(toValue, true);
  const inRange = (date) => {
    if (!date) return false;
    const time = date.getTime();
    const fromOk = from ? time >= from.getTime() : true;
    const toOk = to ? time <= to.getTime() : true;
    return fromOk && toOk;
  };

  // Scope by the SHA claim's own date. A matched top-up keeps its
  // already-computed amount even if the top-up transaction itself was
  // posted on a different date / different month than the SHA claim.
  const shaRecords = fullReconciliation.shaRecords.filter((r) => inRange(r.postingDate));
  const unmatchedShaRecords = fullReconciliation.unmatchedShaRecords.filter((r) => inRange(r.postingDate));
  // Unmatched cash/insurance rows have no linked SHA claim, so they're scoped
  // by their own posting date instead.
  const unmatchedCashRecords = fullReconciliation.unmatchedCashRecords.filter((r) => inRange(r.postingDate));
  const matchedPairs = fullReconciliation.matchedPairs.filter((p) => inRange(p.sha.postingDate));

  const matchedCashAmount = shaRecords.reduce((sum, r) => sum + (Number(r.matchedTopup) || 0), 0);
  const totalShaAmount = shaRecords.reduce((sum, r) => sum + r.amount, 0);
  const unmatchedShaAmount = unmatchedShaRecords.reduce((sum, r) => sum + r.amount, 0);
  const unmatchedCashAmount = unmatchedCashRecords.reduce((sum, r) => sum + r.amount, 0);
  const totalSourceAmount = filteredRecords.reduce((sum, r) => sum + r.amount, 0);
  const totalCashAmount = filteredRecords.filter((r) => !r.isSha).reduce((sum, r) => sum + r.amount, 0);

  return {
    ...fullReconciliation,
    shaRecords,
    unmatchedShaRecords,
    unmatchedCashRecords,
    matchedPairs,
    totals: {
      totalSourceAmount,
      totalShaAmount,
      totalCashAmount,
      matchedCashAmount,
      unmatchedCashAmount,
      unmatchedShaAmount,
      matchedPairCount: matchedPairs.length,
      unmatchedShaCount: unmatchedShaRecords.length,
      unmatchedCashCount: unmatchedCashRecords.length,
    },
  };
}

function buildDerivedData(records, fromValue, toValue) {
  const reconciliation = scopeReconciliationToDateRange(state.fullReconciliation, records, fromValue, toValue);
  const shaStudies = buildShaStudiesFromReconciliation(reconciliation.shaRecords);
  const insuranceStats = computeInsuranceStats(records, reconciliation);
  const insuranceTopupTotal = insuranceStats.reduce((sum, row) => sum + (Number(row.topupSha) || 0), 0);

  const kpis = computeKpis(shaStudies, insuranceTopupTotal);
  const modalityStats = computeModalityStats(shaStudies, kpis.totalRevenueGenerated, insuranceTopupTotal);
  const procedureStats = computeProcedureStats(shaStudies);
  const weeklyStats = computeWeeklyStats(shaStudies);

  return {
    kpis,
    modalityStats,
    insuranceStats,
    procedureStats,
    weeklyStats,
    reconciliation,
  };
}

function reconcileShaCashRecords(records) {
  const shaRecords = records
    .filter((record) => record.isSha)
    .map((record, index) => ({
      ...record,
      sourceIndex: index,
      matchedTopup: 0,
      matchedCashRecord: null,
      matchStatus: "unmatched",
      matchReason: "",
    }));

  const cashRecords = records
    .filter((record) => !record.isSha)
    .map((record, index) => ({
      ...record,
      sourceIndex: index,
      matched: false,
      matchedShaKey: null,
    }));

  const examinations = buildShaExaminations(shaRecords);

  const matchedPairs = [];

  examinations.forEach((exam) => {
    const match = findUniqueCashMatch(exam, cashRecords);
    if (!match) {
      exam.matchStatus = "unmatched";
      exam.matchReason = "No unique non-SHA match on patient+clinical+modality";
      exam.shaRows.forEach((row) => {
        row.matchStatus = "unmatched";
        row.matchReason = exam.matchReason;
      });
      return;
    }

    match.cash.matched = true;
    match.cash.matchedShaKey = exam.examKey;

    exam.matchedCashRecord = match.cash;
    exam.matchStatus = "matched";
    exam.matchReason = "";

    distributeExamTopup(exam.shaRows, match.cash.amount);
    exam.shaRows.forEach((row) => {
      row.matchedCashRecord = match.cash;
      row.matchStatus = "matched";
      row.matchReason = "";
    });

    matchedPairs.push({
      sha: {
        patientName: exam.patientName,
        dayKey: exam.dayKey,
        postingDate: exam.postingDate,
        itemNo: exam.itemNo,
        procedureName: exam.procedureName,
        modality: exam.modality,
        amount: exam.shaAmount,
      },
      cash: match.cash,
      rule: match.rule,
    });
  });

  const unmatchedShaRecords = shaRecords.filter((record) => record.matchStatus !== "matched");
  const unmatchedCashRecords = cashRecords.filter((record) => !record.matched);

  const totalSourceAmount = records.reduce((sum, record) => sum + record.amount, 0);
  const totalShaAmount = shaRecords.reduce((sum, record) => sum + record.amount, 0);
  const totalCashAmount = cashRecords.reduce((sum, record) => sum + record.amount, 0);
  const matchedCashAmount = matchedPairs.reduce((sum, pair) => sum + pair.cash.amount, 0);
  const unmatchedCashAmount = unmatchedCashRecords.reduce((sum, record) => sum + record.amount, 0);
  const unmatchedShaAmount = unmatchedShaRecords.reduce((sum, record) => sum + record.amount, 0);

  return {
    shaRecords,
    cashRecords,
    examinations,
    matchedPairs,
    unmatchedShaRecords,
    unmatchedCashRecords,
    totals: {
      totalSourceAmount,
      totalShaAmount,
      totalCashAmount,
      matchedCashAmount,
      unmatchedCashAmount,
      unmatchedShaAmount,
      matchedPairCount: matchedPairs.length,
      unmatchedShaCount: unmatchedShaRecords.length,
      unmatchedCashCount: unmatchedCashRecords.length,
    },
  };
}

function buildShaExaminations(shaRecords) {
  const map = new Map();

  shaRecords.forEach((row) => {
    const key = `${row.patientNorm}|${row.dayKey}|${row.modality}`;
    if (!map.has(key)) {
      map.set(key, {
        examKey: key,
        patientName: row.patientName,
        patientNorm: row.patientNorm,
        patientLoose: row.patientLoose,
        patientFirstSecond: row.patientFirstSecond,
        patientFirstLast: row.patientFirstLast,
        patientFirstThree: row.patientFirstThree,
        dayKey: row.dayKey,
        postingDate: row.postingDate,
        modality: row.modality,
        itemNo: row.itemNo || "",
        itemNoNorm: row.itemNoNorm || "",
        procedureNorms: [],
        procedureName: "",
        regionSet: new Set(),
        shaAmount: 0,
        shaRows: [],
        matchStatus: "unmatched",
        matchReason: "",
        matchedCashRecord: null,
      });
    }

    const exam = map.get(key);
    exam.shaRows.push(row);
    exam.shaAmount += row.amount;
    if (row.procedureNorm) {
      exam.procedureNorms.push(row.procedureNorm);
    }
    (row.clinicalRegions || []).forEach((region) => exam.regionSet.add(region));
    if (!exam.itemNoNorm && row.itemNoNorm) {
      exam.itemNoNorm = row.itemNoNorm;
      exam.itemNo = row.itemNo;
    }
  });

  return [...map.values()].map((exam) => {
    const uniqueProcedureNames = [...new Set(exam.shaRows.map((row) => row.procedureName).filter(Boolean))];
    exam.procedureName = uniqueProcedureNames.join(" + ");
    exam.procedureNorms = [...new Set(exam.procedureNorms)];
    exam.procedureTokens = extractClinicalTokens(exam.procedureNorms.join(" "));
    return exam;
  });
}

function distributeExamTopup(shaRows, totalTopup) {
  if (!shaRows.length || !totalTopup) {
    return;
  }

  const totalSha = shaRows.reduce((sum, row) => sum + (Number(row.amount) || 0), 0);
  if (totalSha <= 0) {
    const equalShare = totalTopup / shaRows.length;
    shaRows.forEach((row) => {
      row.matchedTopup += equalShare;
    });
    return;
  }

  shaRows.forEach((row) => {
    row.matchedTopup += (totalTopup * (Number(row.amount) || 0)) / totalSha;
  });
}

function findUniqueCashMatch(sha, cashRecords) {
  const available = cashRecords.filter((cash) => !cash.matched && isSamePatientAdvanced(sha, cash));

  if (!available.length) {
    return null;
  }

  const exactName = available.filter((cash) => cash.patientNorm === sha.patientNorm);
  const firstThreeName = available.filter(
    (cash) => sha.patientFirstThree && cash.patientFirstThree && cash.patientFirstThree === sha.patientFirstThree
  );
  const firstLastName = available.filter(
    (cash) => sha.patientFirstLast && cash.patientFirstLast && cash.patientFirstLast === sha.patientFirstLast
  );
  const firstSecondName = available.filter(
    (cash) =>
      sha.patientFirstSecond &&
      cash.patientFirstSecond &&
      cash.patientFirstSecond === sha.patientFirstSecond
  );

  let namePool = available;
  let nameRule = "name-compatible";
  if (exactName.length) {
    namePool = exactName;
    nameRule = "name-exact";
  } else if (firstThreeName.length) {
    namePool = firstThreeName;
    nameRule = "name-first-three";
  } else if (firstLastName.length) {
    namePool = firstLastName;
    nameRule = "name-first-last";
  } else if (firstSecondName.length) {
    namePool = firstSecondName;
    nameRule = "name-first-second";
  }

  const byClinicalExam = namePool.filter((cash) => isClinicalExaminationRelated(sha, cash));
  if (!byClinicalExam.length) {
    return null;
  }

  const hasItem = Boolean(sha.itemNoNorm);
  const hasProcedure = Boolean(sha.procedureNorm || (sha.procedureNorms && sha.procedureNorms.length));
  const hasModality = Boolean(sha.modality && sha.modality !== "Other");

  const byItem = hasItem
    ? byClinicalExam.filter((cash) => Boolean(cash.itemNoNorm) && cash.itemNoNorm === sha.itemNoNorm)
    : [];
  const byProcedure = hasProcedure
    ? byClinicalExam.filter((cash) => isProcedureRelated(sha, cash))
    : [];
  const byModality = hasModality
    ? byClinicalExam.filter((cash) => cash.modality === sha.modality)
    : byClinicalExam;

  const selectUnique = (rows, rule) =>
    rows.length === 1 ? { cash: rows[0], rule: `${nameRule}+${rule}` } : null;
  const intersect = (left, right) => {
    if (!left.length || !right.length) {
      return [];
    }
    const set = new Set(right.map((row) => row.sourceIndex));
    return left.filter((row) => set.has(row.sourceIndex));
  };

  const itemAndProcedure = intersect(byItem, byProcedure);
  const itemAndModality = intersect(byItem, byModality);
  const procedureAndModality = intersect(byProcedure, byModality);

  const orderedCandidates = [
    { rows: intersect(itemAndProcedure, byModality), rule: "patient+clinical+itemno+procedure+modality" },
    { rows: itemAndProcedure, rule: "patient+clinical+itemno+procedure" },
    { rows: procedureAndModality, rule: "patient+clinical+procedure+modality" },
    { rows: itemAndModality, rule: "patient+clinical+itemno+modality" },
    { rows: byItem, rule: "patient+clinical+itemno" },
    { rows: byProcedure, rule: "patient+clinical+procedure" },
    { rows: byModality, rule: "patient+clinical+modality" },
  ];

  for (const candidate of orderedCandidates) {
    if (!candidate.rows.length) {
      continue;
    }

    const unique = selectUnique(candidate.rows, candidate.rule);
    if (unique) {
      return unique;
    }

    const closest = selectClosestByDate(candidate.rows, sha.postingDate);
    if (closest) {
      return {
        cash: closest,
        rule: `${nameRule}+${candidate.rule}+closest-date`,
      };
    }
  }

  return null;
}

function selectClosestByDate(rows, targetDate) {
  if (!rows.length) {
    return null;
  }

  let minDistance = Number.POSITIVE_INFINITY;
  let closest = [];

  rows.forEach((row) => {
    const distance = daysBetween(targetDate, row.postingDate);
    if (distance < minDistance) {
      minDistance = distance;
      closest = [row];
      return;
    }

    if (distance === minDistance) {
      closest.push(row);
    }
  });

  return closest.length === 1 ? closest[0] : null;
}

function isProcedureRelated(sha, cash) {
  if (!cash.procedureNorm) {
    return false;
  }

  const shaNorms = Array.isArray(sha.procedureNorms) ? sha.procedureNorms : [sha.procedureNorm].filter(Boolean);
  if (shaNorms.includes(cash.procedureNorm)) {
    return true;
  }

  const shaTokens = sha.procedureTokens || extractClinicalTokens(shaNorms.join(" "));
  const cashTokens = extractClinicalTokens(cash.procedureNorm);
  if (!shaTokens.size || !cashTokens.size) {
    return false;
  }

  let overlap = 0;
  cashTokens.forEach((token) => {
    if (shaTokens.has(token)) {
      overlap += 1;
    }
  });

  return overlap > 0 && overlap === cashTokens.size;
}

function isClinicalExaminationRelated(sha, cash) {
  const shaRegions = new Set(Array.isArray(sha.regionSet) ? sha.regionSet : [...(sha.regionSet || [])]);
  const cashRegions = new Set(cash.clinicalRegions || []);

  if (shaRegions.size && cashRegions.size) {
    let overlap = 0;
    cashRegions.forEach((region) => {
      if (shaRegions.has(region)) {
        overlap += 1;
      }
    });

    const minSize = Math.min(shaRegions.size, cashRegions.size);
    if (minSize > 0 && overlap / minSize >= 0.6) {
      return true;
    }
  }

  return isProcedureRelated(sha, cash);
}

function extractBodyRegions(text) {
  const norm = normalizeText(text);
  const regions = new Set();

  const patterns = [
    { key: "chest", regex: /\b(chest|thorax|thoracic|lung|lungs|hrct chest)\b/ },
    { key: "abdomen", regex: /\b(abdomen|abdominal|upper abdomen|lower abdomen)\b/ },
    { key: "pelvis", regex: /\b(pelvis|pelvic)\b/ },
    { key: "brain", regex: /\b(brain|head|cranial|skull)\b/ },
    { key: "spine", regex: /\b(spine|spinal|cervical|thoracic spine|lumbar|sacral)\b/ },
    { key: "neck", regex: /\b(neck|carotid)\b/ },
    { key: "upper-limb", regex: /\b(shoulder|arm|elbow|wrist|hand|humerus|forearm)\b/ },
    { key: "lower-limb", regex: /\b(hip|thigh|knee|leg|ankle|foot|femur|tibia|fibula)\b/ },
    { key: "face", regex: /\b(face|facial|sinus|orbit|orbital|jaw|mandible|maxilla)\b/ },
  ];

  patterns.forEach((pattern) => {
    if (pattern.regex.test(norm)) {
      regions.add(pattern.key);
    }
  });

  return regions;
}

function extractClinicalTokens(text) {
  const cleaned = normalizeText(text)
    .replace(/\b(mri|ct|ultrasound|x ray|xray|sedation|scan|study|exam|with|without|and|for|of)\b/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  const set = new Set();
  cleaned.split(" ").forEach((token) => {
    if (token && token.length >= 3) {
      set.add(token);
    }
  });
  return set;
}

function buildShaStudiesFromReconciliation(shaRecords) {
  const shaMap = new Map();

  shaRecords.forEach((record) => {
    const key = `${record.patientNorm}|${record.dayKey}|${record.procedureNorm}`;
    if (!shaMap.has(key)) {
      shaMap.set(key, {
        key,
        postingDate: record.postingDate,
        dayKey: record.dayKey,
        patientNorm: record.patientNorm,
        patientLoose: record.patientLoose,
        patientName: record.patientName || "Unknown Patient",
        procedureName: record.procedureName || "Unknown Procedure",
        modality: record.modality,
        weekKey: record.weekKey,
        weekLabel: record.weekLabel,
        shaBilling: 0,
        patientTopup: 0,
      });
    }

    const study = shaMap.get(key);
    study.shaBilling += record.amount;
    study.patientTopup += record.matchedTopup || 0;
    if (study.modality === "Other" && record.modality !== "Other") {
      study.modality = record.modality;
    }
  });

  return [...shaMap.values()]
    .filter((study) => study.shaBilling > 0)
    .map((study) => ({
    ...study,
    totalRevenue: study.shaBilling + study.patientTopup,
  }));
}

function allocateAmountByBilling(candidates, amount, targetField, weightField) {
  if (!candidates.length || !amount) {
    return;
  }

  const totalWeight = candidates.reduce((sum, item) => sum + (Number(item[weightField]) || 0), 0);
  if (totalWeight > 0) {
    candidates.forEach((item) => {
      item[targetField] += (amount * (Number(item[weightField]) || 0)) / totalWeight;
    });
    return;
  }

  const equalShare = amount / candidates.length;
  candidates.forEach((item) => {
    item[targetField] += equalShare;
  });
}

function computeKpis(shaStudies, insuranceTopupTotal = 0) {
  const totalShaBilling = shaStudies.reduce((sum, s) => sum + s.shaBilling, 0);
  const matchedCashTopup = shaStudies.reduce((sum, s) => sum + s.patientTopup, 0);
  const totalPatientTopup = matchedCashTopup + insuranceTopupTotal;
  const totalRevenueGenerated = totalShaBilling + totalPatientTopup;
  const totalShaStudies = shaStudies.length;

  const avgShaPerStudy = totalShaStudies ? totalShaBilling / totalShaStudies : 0;
  const avgTopupPerStudy = totalShaStudies ? totalPatientTopup / totalShaStudies : 0;

  const shaCoveragePct = totalRevenueGenerated ? (totalShaBilling / totalRevenueGenerated) * 100 : 0;
  const patientContributionPct = totalRevenueGenerated ? (totalPatientTopup / totalRevenueGenerated) * 100 : 0;

  return {
    totalShaBilling,
    matchedCashTopup,
    insuranceTopupTotal,
    totalPatientTopup,
    totalRevenueGenerated,
    totalShaStudies,
    avgShaPerStudy,
    avgTopupPerStudy,
    shaCoveragePct,
    patientContributionPct,
  };
}

function computeModalityStats(shaStudies, totalRevenueGenerated, insuranceTopupTotal = 0) {
  const seed = MODALITIES.map((modality) => ({
    modality,
    studies: 0,
    shaBilling: 0,
    patientTopup: 0,
    totalRevenue: 0,
    avgSha: 0,
    avgTopup: 0,
    sharePct: 0,
  }));

  const map = Object.fromEntries(seed.map((row) => [row.modality, row]));

  shaStudies.forEach((study) => {
    const key = map[study.modality] ? study.modality : "Other";
    const bucket = map[key];

    bucket.studies += 1;
    bucket.shaBilling += study.shaBilling;
    bucket.patientTopup += study.patientTopup;
    bucket.totalRevenue += study.totalRevenue;
  });

  const totalShaBilling = seed.reduce((sum, row) => sum + row.shaBilling, 0);
  if (insuranceTopupTotal > 0 && totalShaBilling > 0) {
    seed.forEach((row) => {
      const allocatedInsuranceTopup = (insuranceTopupTotal * row.shaBilling) / totalShaBilling;
      row.patientTopup += allocatedInsuranceTopup;
      row.totalRevenue += allocatedInsuranceTopup;
    });
  }

  seed.forEach((row) => {
    row.avgSha = row.studies ? row.shaBilling / row.studies : 0;
    row.avgTopup = row.studies ? row.patientTopup / row.studies : 0;
    row.sharePct = totalRevenueGenerated ? (row.totalRevenue / totalRevenueGenerated) * 100 : 0;
  });

  return seed;
}

function computeInsuranceStats(records, reconciliation) {
  const insurerMap = new Map();
  let totalRevenueAllInsurers = 0;

  const ensureInsurer = (insurerKey, providerName) => {
    if (!insurerMap.has(insurerKey)) {
      insurerMap.set(insurerKey, {
        provider: providerName,
        insuranceBilling: 0,
        topupCash: 0,
        topupSha: 0,
        totalRevenue: 0,
        studiesSet: new Set(),
        studies: 0,
        avgBilling: 0,
        sharePct: 0,
        isSha: insurerKey === SHA_NAME,
      });
    }
    return insurerMap.get(insurerKey);
  };

  records.forEach((record) => {
    if (record.isTopup) {
      return;
    }

    const insurerKey = getCanonicalProviderKey(record.customerNorm);
    const insurerAgg = ensureInsurer(insurerKey, record.customerName || "Unknown Provider");
    insurerAgg.insuranceBilling += record.amount;
    insurerAgg.studiesSet.add(`${record.patientNorm}|${record.dayKey}|${record.procedureNorm}`);
  });

  const shaInsurer = ensureInsurer(SHA_NAME, "Social Health Insurance Fund");
  shaInsurer.topupCash += reconciliation?.totals?.matchedCashAmount || 0;

  const rows = [...insurerMap.values()]
    .map((item) => {
      const studies = item.studiesSet.size;
      const totalRevenue = item.insuranceBilling + item.topupCash + item.topupSha;
      totalRevenueAllInsurers += totalRevenue;
      return {
        provider: item.provider,
        insuranceBilling: item.insuranceBilling,
        topupCash: item.topupCash,
        topupSha: item.isSha ? 0 : item.topupSha,
        totalRevenue,
        studies,
        avgBilling: studies ? item.insuranceBilling / studies : 0,
        sharePct: 0,
        isSha: item.isSha,
      };
    })
    .sort((a, b) => b.totalRevenue - a.totalRevenue);

  rows.forEach((row) => {
    row.sharePct = totalRevenueAllInsurers ? (row.totalRevenue / totalRevenueAllInsurers) * 100 : 0;
  });

  return rows;
}

function computeProcedureStats(shaStudies) {
  const grouped = new Map();

  shaStudies.forEach((study) => {
    const key = study.procedureName.toLowerCase();
    if (!grouped.has(key)) {
      grouped.set(key, {
        procedureName: study.procedureName,
        studies: 0,
        shaBilling: 0,
        patientTopup: 0,
        totalRevenue: 0,
      });
    }

    const item = grouped.get(key);
    item.studies += 1;
    item.shaBilling += study.shaBilling;
    item.patientTopup += study.patientTopup;
    item.totalRevenue += study.totalRevenue;
  });

  return [...grouped.values()].sort((a, b) => b.totalRevenue - a.totalRevenue);
}

function computeWeeklyStats(shaStudies) {
  const map = new Map();

  shaStudies.forEach((study) => {
    if (!map.has(study.weekKey)) {
      map.set(study.weekKey, {
        weekKey: study.weekKey,
        weekLabel: study.weekLabel,
        shaBilling: 0,
        patientTopup: 0,
        totalRevenue: 0,
        studies: 0,
      });
    }

    const bucket = map.get(study.weekKey);
    bucket.shaBilling += study.shaBilling;
    bucket.patientTopup += study.patientTopup;
    bucket.totalRevenue += study.totalRevenue;
    bucket.studies += 1;
  });

  return [...map.values()].sort((a, b) => a.weekKey.localeCompare(b.weekKey));
}

function renderKPIs(kpis) {
  refs.kpiShaBilling.textContent = formatCurrency(kpis.totalShaBilling);
  refs.kpiTopup.textContent = formatCurrency(kpis.totalPatientTopup);
  refs.kpiTotalRevenue.textContent = formatCurrency(kpis.totalRevenueGenerated);
  refs.kpiStudies.textContent = formatNumber(kpis.totalShaStudies);
  refs.kpiAvgSha.textContent = formatCurrency(kpis.avgShaPerStudy);
  refs.kpiAvgTopup.textContent = formatCurrency(kpis.avgTopupPerStudy);
  refs.kpiCoverage.textContent = `${kpis.shaCoveragePct.toFixed(2)}%`;
  refs.kpiPatientPct.textContent = `${kpis.patientContributionPct.toFixed(2)}%`;
}

function renderModalityTable(modalityStats) {
  const focusModalities = ["MRI", "CT", "Ultrasound", "X-Ray"];
  const tableRows = modalityStats.filter((item) => focusModalities.includes(item.modality));

  if (!tableRows.length) {
    refs.modalityTableBody.innerHTML = '<tr><td colspan="8" class="text-center py-4">No modality records available.</td></tr>';
    return;
  }

  refs.modalityTableBody.innerHTML = tableRows
    .map(
      (item) =>
        `<tr>
          <td>${escapeHtml(item.modality)}</td>
          <td class="text-end">${formatNumber(item.studies)}</td>
          <td class="text-end">${formatCurrency(item.shaBilling)}</td>
          <td class="text-end">${formatCurrency(item.patientTopup)}</td>
          <td class="text-end">${formatCurrency(item.totalRevenue)}</td>
          <td class="text-end">${formatCurrency(item.avgSha)}</td>
          <td class="text-end">${formatCurrency(item.avgTopup)}</td>
          <td class="text-end">${item.sharePct.toFixed(2)}%</td>
        </tr>`
    )
    .join("");
}

function renderInsuranceTable(insuranceStats) {
  if (!insuranceStats.length) {
    refs.insuranceTableBody.innerHTML = '<tr><td colspan="8" class="text-center py-4">No insurance records available.</td></tr>';
    return;
  }

  refs.insuranceTableBody.innerHTML = insuranceStats
    .map((item) => {
      const rowClass = item.isSha ? "sha-row" : "";
      return `<tr class="${rowClass}">
        <td>${escapeHtml(item.provider)}${item.isSha ? " <strong>(SHA)</strong>" : ""}</td>
        <td class="text-end">${formatCurrency(item.insuranceBilling)}</td>
        <td class="text-end">${formatCurrency(item.topupCash)}</td>
        <td class="text-end">${formatCurrency(item.topupSha)}</td>
        <td class="text-end">${formatCurrency(item.totalRevenue)}</td>
        <td class="text-end">${formatNumber(item.studies)}</td>
        <td class="text-end">${formatCurrency(item.avgBilling)}</td>
        <td class="text-end">${item.sharePct.toFixed(2)}%</td>
      </tr>`;
    })
    .join("");
}

function renderProcedureTable(procedureStats) {
  if (!procedureStats.length) {
    refs.procedureTableBody.innerHTML = '<tr><td colspan="5" class="text-center py-4">No procedure records available.</td></tr>';
    return;
  }

  refs.procedureTableBody.innerHTML = procedureStats
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
  renderModalityChart(derived.modalityStats);
  renderInsuranceShareChart(derived.insuranceStats);
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
      scales: {
        y: {
          beginAtZero: true,
          ticks: { callback: (value) => shortCurrency(value) },
          grid: { color: "rgba(22, 49, 73, 0.08)" },
        },
        x: {
          grid: { display: false },
        },
      },
      plugins: {
        legend: { position: "bottom" },
        tooltip: {
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
          grid: { color: "rgba(22, 49, 73, 0.08)" },
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
  const topProviders = insuranceStats.slice(0, 8);

  state.charts.insuranceShare = upsertChart(state.charts.insuranceShare, "insuranceShareChart", {
    type: "doughnut",
    data: {
      labels: topProviders.map((x) => x.provider),
      datasets: [
        {
          data: topProviders.map((x) => round2(x.totalRevenue)),
          backgroundColor: [
            "#0b72c9",
            "#2f84d5",
            "#57a5ea",
            "#7dbef5",
            "#9ccdf8",
            "#beddf9",
            "#8aa6c3",
            "#b2bfd0",
          ],
          borderColor: "#ffffff",
          borderWidth: 2,
        },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      cutout: "62%",
      plugins: {
        legend: {
          position: "bottom",
          labels: { usePointStyle: true, boxWidth: 10 },
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
    totalPatientTopup: 0,
    totalRevenueGenerated: 0,
    totalShaStudies: 0,
    avgShaPerStudy: 0,
    avgTopupPerStudy: 0,
    shaCoveragePct: 0,
    patientContributionPct: 0,
  });

  refs.modalityTableBody.innerHTML = '<tr><td colspan="8" class="text-center py-4">No data available.</td></tr>';
  refs.insuranceTableBody.innerHTML = '<tr><td colspan="8" class="text-center py-4">No data available.</td></tr>';
  refs.procedureTableBody.innerHTML = '<tr><td colspan="5" class="text-center py-4">No data available.</td></tr>';

  if (refs.matchedPairsBody) {
    refs.matchedPairsBody.innerHTML = '<tr><td colspan="8" class="text-center py-3">No reconciliation data.</td></tr>';
  }
  if (refs.unmatchedShaBody) {
    refs.unmatchedShaBody.innerHTML = '<tr><td colspan="7" class="text-center py-3">No reconciliation data.</td></tr>';
  }
  if (refs.unmatchedCashBody) {
    refs.unmatchedCashBody.innerHTML = '<tr><td colspan="6" class="text-center py-3">No reconciliation data.</td></tr>';
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
