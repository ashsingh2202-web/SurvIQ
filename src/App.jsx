import React, { useState, useMemo, useCallback, useEffect } from "react";
import Papa from "papaparse";
import * as XLSX from "xlsx";
import {
  XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer,
  BarChart, Bar, Cell, AreaChart, Area
} from "recharts";
import { Upload, AlertTriangle, CheckCircle2, TrendingUp, Table2, FileText, ChevronRight, Download, FileSpreadsheet, Layers, ShieldCheck, Info } from "lucide-react";

// =====================================================================
// SEMANTIC COLUMN CLASSIFICATION
// Name-based rules run first (highest precision when column names are
// sane), then value-based statistical fallback for unnamed/ambiguous
// columns. This ordering is the direct fix for "ZIP treated as date" and
// "phone summarized as a numeric measurement" — those columns are now
// recognized as their own semantic types (zip, phone, email, identifier,
// geo) and are structurally excluded from numeric/date pipelines rather
// than filtered after the fact.
// =====================================================================

const isMissing = (v) => v === null || v === undefined || v === "" ||
  (typeof v === "string" && ["NA", "N/A", "NULL", "."].includes(v.trim().toUpperCase()));

function excelSerialToDate(n) {
  const utcDays = Math.floor(n - 25569);
  return new Date(utcDays * 86400 * 1000);
}

function parseDateLoose(v) {
  if (isMissing(v)) return null;
  if (v instanceof Date && !isNaN(v)) return v;
  if (typeof v === "number" && v > 20000 && v < 60000) return excelSerialToDate(v);
  const s = String(v).trim();
  const d = new Date(s);
  if (!isNaN(d) && /\d{4}/.test(s)) return d;
  const m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
  if (m) {
    let [, a, b, y] = m;
    if (y.length === 2) y = "20" + y;
    const d2 = new Date(`${y}-${a.padStart(2, "0")}-${b.padStart(2, "0")}`);
    if (!isNaN(d2)) return d2;
  }
  return null;
}

// Word-boundary name test: matches "record_id" / "Patient ID" / "ID" but
// NOT "Valid" or "Video" — this specific guard is why identifiers won't
// be misfired on ordinary words that merely contain "id" as a substring.
function nameHasWord(name, words) {
  const n = name.toLowerCase();
  return words.some((w) => new RegExp(`(^|[_\\s-])${w}([_\\s-]|$)`).test(n));
}
function nameContains(name, subs) {
  const n = name.toLowerCase();
  return subs.some((s) => n.includes(s));
}

function classifySemanticType(colName, values) {
  const nonMissing = values.filter((v) => !isMissing(v));
  if (nonMissing.length === 0) return { type: "empty" };

  // ---- Name-based rules (checked in priority order) ----
  if (nameContains(colName, ["email", "e-mail"])) return { type: "email" };
  if (nameHasWord(colName, ["phone", "mobile", "telephone", "cell", "fax"])) return { type: "phone" };
  if (nameHasWord(colName, ["zip", "postal", "postcode"])) return { type: "zip" };
  if (nameHasWord(colName, ["lat", "latitude", "lng", "lon", "longitude"])) return { type: "geo" };
  if (nameContains(colName, ["dob", "birthdate", "birthday"]) || (nameContains(colName, ["birth"]) && nameContains(colName, ["date"]))) return { type: "date", subtype: "dob" };
  if (nameContains(colName, ["date"]) || nameHasWord(colName, ["dt"])) return { type: "date", subtype: "event" };
  if (nameHasWord(colName, ["id", "mrn", "uuid", "guid"])) return { type: "identifier" };
  if (nameHasWord(colName, ["age"])) return { type: "numeric", subtype: "age" };
  if (nameHasWord(colName, ["gender", "sex"])) return { type: "categorical", subtype: "sex" };

  // ---- Value-based fallback ----
  const sample = nonMissing.slice(0, 300);
  const numericCount = sample.filter((v) => typeof v === "number" || (v !== "" && !isNaN(parseFloat(v)) && isFinite(v))).length;
  const dateCount = sample.filter((v) => parseDateLoose(v) !== null).length;
  const uniqueRatio = new Set(nonMissing.map(String)).size / nonMissing.length;

  // High-cardinality columns (near-unique per row) behave like identifiers
  // even with no naming hint — this is the "duplicate detection on things
  // that are actually keys" fix.
  if (uniqueRatio > 0.95 && nonMissing.length > 10) return { type: "identifier" };

  if (dateCount / sample.length > 0.85) return { type: "date", subtype: "event" };

  if (numericCount / sample.length > 0.85) {
    const uniqueVals = new Set(nonMissing.map((v) => String(v)));
    if (uniqueVals.size <= 6 && nonMissing.every((v) => Number.isInteger(parseFloat(v)))) {
      return { type: "categorical", numericCoded: true };
    }
    return { type: "numeric" };
  }

  const uniqueVals = new Set(nonMissing.map((v) => String(v).trim()));
  if (uniqueVals.size <= Math.max(20, nonMissing.length * 0.5)) return { type: "categorical" };
  return { type: "text" };
}

// =====================================================================
// Stats helpers (unchanged core math, still auditable against R/SPSS)
// =====================================================================
function minOf(arr) { return arr.reduce((a, b) => (b < a ? b : a), arr[0]); }
function maxOf(arr) { return arr.reduce((a, b) => (b > a ? b : a), arr[0]); }
function mean(arr) { return arr.reduce((a, b) => a + b, 0) / arr.length; }
function median(arr) { const s = [...arr].sort((a, b) => a - b); const mid = Math.floor(s.length / 2); return s.length % 2 ? s[mid] : (s[mid - 1] + s[mid]) / 2; }
function stddev(arr, m) { if (arr.length < 2) return 0; const v = arr.reduce((a, b) => a + (b - m) ** 2, 0) / (arr.length - 1); return Math.sqrt(v); }
function quantile(arr, q) { const s = [...arr].sort((a, b) => a - b); const pos = (s.length - 1) * q; const base = Math.floor(pos); const rest = pos - base; return s[base + 1] !== undefined ? s[base] + rest * (s[base + 1] - s[base]) : s[base]; }
function ageAt(dob, ref) { let age = ref.getFullYear() - dob.getFullYear(); const m = ref.getMonth() - dob.getMonth(); if (m < 0 || (m === 0 && ref.getDate() < dob.getDate())) age--; return age; }

// =====================================================================
// Fuzzy category consistency (COVID19 vs COVID-19, M vs Male, etc.)
// =====================================================================
function levenshtein(a, b) {
  const m = a.length, n = b.length;
  const dp = Array.from({ length: m + 1 }, () => new Array(n + 1).fill(0));
  for (let i = 0; i <= m; i++) dp[i][0] = i;
  for (let j = 0; j <= n; j++) dp[0][j] = j;
  for (let i = 1; i <= m; i++) for (let j = 1; j <= n; j++)
    dp[i][j] = a[i - 1] === b[j - 1] ? dp[i - 1][j - 1] : 1 + Math.min(dp[i - 1][j - 1], dp[i - 1][j], dp[i][j - 1]);
  return dp[m][n];
}
const normalizeCat = (s) => String(s).toUpperCase().replace(/[^A-Z0-9]/g, "");
const SYNONYM_PAIRS = [["M", "MALE"], ["F", "FEMALE"], ["Y", "YES"], ["N", "NO"], ["U", "UNKNOWN"], ["UNK", "UNKNOWN"]];
function findCategoryInconsistencies(levels) {
  const found = [];
  const items = levels.slice(0, 25).map((l) => ({ raw: String(l.level), norm: normalizeCat(l.level) }));
  for (let i = 0; i < items.length; i++) {
    for (let j = i + 1; j < items.length; j++) {
      const a = items[i], b = items[j];
      if (a.norm === b.norm && a.raw !== b.raw) { found.push({ pair: [a.raw, b.raw], reason: "same after removing punctuation/case", similarity: 100 }); continue; }
      const isSynonym = SYNONYM_PAIRS.some(([x, y]) => (a.norm === x && b.norm === y) || (a.norm === y && b.norm === x));
      if (isSynonym) { found.push({ pair: [a.raw, b.raw], reason: "likely coded synonym", similarity: null }); continue; }
      if (a.norm.length >= 4 && b.norm.length >= 4) {
        const dist = levenshtein(a.norm, b.norm);
        const maxLen = Math.max(a.norm.length, b.norm.length);
        const sim = Math.round((1 - dist / maxLen) * 100);
        if (dist <= 1 && sim >= 85) found.push({ pair: [a.raw, b.raw], reason: "possible spelling variant", similarity: sim });
      }
    }
  }
  return found;
}

// =====================================================================
// Sample data
// =====================================================================
function generateSampleData() {
  const rows = [];
  const substances = ["Fentanyl", "Heroin", "Cocaine", "Methamphetamine", "Alcohol", "Benzodiazepines"];
  const sexes = ["Male", "Female"];
  const outcomes = ["Survived", "Died"];
  const zips = ["33301", "33304", "33308", "33312", "33316"];
  let id = 1000;
  for (let m = 0; m < 18; m++) {
    const baseDate = new Date(2025, m, 1);
    const nCases = 18 + Math.round(Math.random() * 14) + (m > 10 ? 6 : 0);
    for (let i = 0; i < nCases; i++) {
      const day = 1 + Math.floor(Math.random() * 27);
      const dt = new Date(baseDate.getFullYear(), baseDate.getMonth(), day);
      const dob = new Date(dt.getFullYear() - (18 + Math.floor(Math.random() * 55)), Math.floor(Math.random() * 12), 1 + Math.floor(Math.random() * 27));
      rows.push({
        Case_ID: "C" + id++,
        Report_Date: dt.toISOString().slice(0, 10),
        DOB: dob.toISOString().slice(0, 10),
        Age: ageAt(dob, dt),
        Sex: sexes[Math.random() > 0.62 ? 0 : 1],
        Substance: substances[Math.floor(Math.random() * substances.length)],
        Zip_Code: zips[Math.floor(Math.random() * zips.length)],
        Phone: `954${100 + Math.floor(Math.random() * 900)}${1000 + Math.floor(Math.random() * 9000)}`,
        Naloxone_Administered: Math.random() > 0.35 ? "Yes" : "No",
        Outcome: outcomes[Math.random() > 0.88 ? 1 : 0],
      });
    }
  }
  rows[3].Age = 214;
  rows[10].Age = "";
  rows[20].Sex = "";
  rows[7].Sex = "M"; // will be flagged against "Male"/"Female" as a synonym
  rows.push({ ...rows[5] });
  rows[15].Phone = "123";
  return rows;
}

// =====================================================================
// File parsing
// =====================================================================
function parseCSV(file) {
  return new Promise((resolve, reject) => {
    Papa.parse(file, {
      header: true, skipEmptyLines: true, dynamicTyping: true,
      complete: (results) => resolve({ rows: results.data, sheetNames: null, activeSheet: null }),
      error: (err) => reject(err),
    });
  });
}
function parseExcel(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(new Uint8Array(e.target.result), { type: "array", cellDates: true });
        const useSheet = wb.SheetNames[0];
        const rows = XLSX.utils.sheet_to_json(wb.Sheets[useSheet], { defval: "", raw: true });
        resolve({ rows, sheetNames: wb.SheetNames, activeSheet: useSheet, workbook: wb });
      } catch (err) { reject(err); }
    };
    reader.onerror = () => reject(new Error("Could not read file"));
    reader.readAsArrayBuffer(file);
  });
}

// Palette
const P = {
  bg: "#F6F7F9", panel: "#FFFFFF", border: "#E1E4EA", borderStrong: "#CBD1DB",
  ink: "#111826", muted: "#606A7D", mutedLight: "#8991A3", navy: "#1E3A5F",
  teal: "#0E6E66", tealSoft: "#E4F1EF", amber: "#9A6324", amberSoft: "#FBF0DE",
  red: "#A32B1F", redSoft: "#FBEAE8",
};
const CHART_COLORS = ["#1E3A5F", "#0E6E66", "#9A6324", "#5B6472", "#2F5E8C", "#3C8577"];
const TYPE_LABEL = { identifier: "identifier", date: "date", numeric: "numeric", categorical: "categorical", phone: "phone", email: "email", zip: "zip", geo: "geo", text: "text", empty: "empty" };

// =====================================================================
// Main
// =====================================================================
export default function SurvIQ() {
  const [rawRows, setRawRows] = useState(null);
  const [fileName, setFileName] = useState("");
  const [fileKind, setFileKind] = useState(null);
  const [sheetNames, setSheetNames] = useState(null);
  const [activeSheet, setActiveSheet] = useState(null);
  const [workbook, setWorkbook] = useState(null);
  const [groupVar, setGroupVar] = useState(null);
  const [trendCol, setTrendCol] = useState(null);
  const [error, setError] = useState(null);
  const [stage, setStage] = useState("upload");

  const columns = useMemo(() => (!rawRows || rawRows.length === 0 ? [] : Object.keys(rawRows[0])), [rawRows]);

  const colTypes = useMemo(() => {
    if (!rawRows) return {};
    const out = {};
    columns.forEach((c) => { out[c] = classifySemanticType(c, rawRows.map((r) => r[c])); });
    return out;
  }, [rawRows, columns]);

  const numericCols = columns.filter((c) => colTypes[c]?.type === "numeric");
  const categoricalCols = columns.filter((c) => colTypes[c]?.type === "categorical");
  const dateCols = columns.filter((c) => colTypes[c]?.type === "date");
  const identifierCols = columns.filter((c) => colTypes[c]?.type === "identifier");
  const phoneCols = columns.filter((c) => colTypes[c]?.type === "phone");
  const emailCols = columns.filter((c) => colTypes[c]?.type === "email");
  const zipCols = columns.filter((c) => colTypes[c]?.type === "zip");

  // Best default trend column: prefer an "event" date over a "dob" date,
  // and require a plausible parse rate before trusting it — this directly
  // replaces "just use the first numeric/parseable column."
  const bestTrendCol = useMemo(() => {
    if (dateCols.length === 0) return null;
    const scored = dateCols.map((c) => {
      const vals = rawRows.map((r) => r[c]).filter((v) => !isMissing(v));
      const parsed = vals.map(parseDateLoose).filter(Boolean);
      const parseRate = vals.length ? parsed.length / vals.length : 0;
      const now = new Date();
      const plausibleRate = parsed.length ? parsed.filter((d) => d.getFullYear() >= 1900 && d.getFullYear() <= now.getFullYear() + 1).length / parsed.length : 0;
      const isEvent = colTypes[c]?.subtype === "event";
      return { col: c, parseRate, plausibleRate, isEvent };
    }).filter((s) => s.parseRate >= 0.85 && s.plausibleRate >= 0.85);
    scored.sort((a, b) => (b.isEvent - a.isEvent) || (b.parseRate - a.parseRate));
    return scored.length ? scored[0].col : null;
  }, [dateCols, rawRows, colTypes]);

  useEffect(() => { setTrendCol(bestTrendCol); }, [bestTrendCol]);

  const cleanedRows = useMemo(() => {
    if (!rawRows) return [];
    return rawRows.map((r) => {
      const copy = { ...r };
      columns.forEach((c) => {
        if (colTypes[c]?.subtype === "age") {
          const v = parseFloat(r[c]);
          if (!isMissing(r[c]) && (v < 0 || v > 120)) copy[c] = null;
        }
      });
      return copy;
    });
  }, [rawRows, columns, colTypes]);

  // ---- Validation / issue engine ----
  const issues = useMemo(() => {
    if (!rawRows) return [];
    const found = [];
    const now = new Date();

    // 1. Duplicate identifiers — explicit, with row numbers, per identifier column
    identifierCols.forEach((idCol) => {
      const seen = new Map();
      rawRows.forEach((r, i) => { const v = r[idCol]; if (isMissing(v)) return; if (!seen.has(v)) seen.set(v, []); seen.get(v).push(i + 1); });
      const dupEntries = [...seen.entries()].filter(([, idxs]) => idxs.length > 1);
      if (dupEntries.length > 0) {
        const example = dupEntries[0];
        found.push({ level: "error", label: `Duplicate ${idCol} values`, detail: `${dupEntries.length} value(s) repeated — e.g. "${example[0]}" appears at rows ${example[1].join(", ")}` });
      }
    });
    // Duplicate full rows (only when no identifier column exists to key on)
    if (identifierCols.length === 0) {
      const rowStrs = rawRows.map((r) => JSON.stringify(r));
      const dupRows = rowStrs.length - new Set(rowStrs).size;
      if (dupRows > 0) found.push({ level: "warn", label: "Duplicate rows", detail: `${dupRows} exact duplicate row(s) found` });
    }
    // Duplicate emails / phones (info — may be legitimate shared contact)
    emailCols.forEach((c) => {
      const seen = new Map();
      rawRows.forEach((r) => { const v = r[c]; if (isMissing(v)) return; seen.set(v, (seen.get(v) || 0) + 1); });
      const dups = [...seen.values()].filter((n) => n > 1).length;
      if (dups > 0) found.push({ level: "info", label: `Repeated email addresses in "${c}"`, detail: `${dups} address(es) used by more than one row` });
    });
    phoneCols.forEach((c) => {
      const seen = new Map();
      rawRows.forEach((r) => { const v = r[c]; if (isMissing(v)) return; seen.set(v, (seen.get(v) || 0) + 1); });
      const dups = [...seen.values()].filter((n) => n > 1).length;
      if (dups > 0) found.push({ level: "info", label: `Repeated phone numbers in "${c}"`, detail: `${dups} number(s) used by more than one row` });
    });

    // 2. Missingness (severity-scored)
    columns.forEach((c) => {
      const missingCount = rawRows.filter((r) => isMissing(r[c])).length;
      if (missingCount > 0) {
        const pct = (missingCount / rawRows.length) * 100;
        found.push({ level: pct > 20 ? "warn" : "info", label: `Missing values in "${c}"`, detail: `${missingCount} of ${rawRows.length} rows (${pct.toFixed(0)}%)` });
      }
    });

    // 3. Implausible ages
    columns.forEach((c) => {
      if (colTypes[c]?.subtype === "age") {
        const bad = rawRows.filter((r) => { const v = parseFloat(r[c]); return !isMissing(r[c]) && (v < 0 || v > 120); });
        if (bad.length > 0) found.push({ level: "error", label: `Implausible values in "${c}"`, detail: `${bad.length} row(s) outside a plausible 0–120 range`, fixable: c });
      }
    });

    // 4. Unparseable dates
    dateCols.forEach((c) => {
      const bad = rawRows.filter((r) => !isMissing(r[c]) && parseDateLoose(r[c]) === null);
      if (bad.length > 0) found.push({ level: "warn", label: `Unparseable dates in "${c}"`, detail: `${bad.length} row(s) could not be read as a date` });
    });

    // 5. Format validation: email
    emailCols.forEach((c) => {
      const emailRe = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
      const bad = rawRows.filter((r) => !isMissing(r[c]) && !emailRe.test(String(r[c]).trim()));
      if (bad.length > 0) found.push({ level: "warn", label: `Invalid email format in "${c}"`, detail: `${bad.length} row(s) don't match a standard email pattern` });
    });

    // 6. Format validation: phone (10 or 11 digits after stripping formatting)
    phoneCols.forEach((c) => {
      const bad = rawRows.filter((r) => {
        if (isMissing(r[c])) return false;
        const digits = String(r[c]).replace(/\D/g, "");
        return digits.length !== 10 && digits.length !== 11;
      });
      if (bad.length > 0) found.push({ level: "warn", label: `Invalid phone format in "${c}"`, detail: `${bad.length} row(s) aren't 10–11 digits` });
    });

    // 7. Format validation: ZIP, plus the specific leading-zero warning that
    // happens when a ZIP column gets read from Excel as a number.
    zipCols.forEach((c) => {
      const zipRe = /^\d{5}(-\d{4})?$/;
      const rowsWithVal = rawRows.filter((r) => !isMissing(r[c]));
      const bad = rowsWithVal.filter((r) => !zipRe.test(String(r[c]).trim()));
      const shortNumeric = rowsWithVal.filter((r) => typeof r[c] === "number" && String(r[c]).length < 5);
      if (bad.length > 0) found.push({ level: "warn", label: `Invalid ZIP format in "${c}"`, detail: `${bad.length} row(s) aren't 5 (or 5+4) digits` });
      if (shortNumeric.length > 0) found.push({ level: "info", label: `Possible leading zeros lost in "${c}"`, detail: `${shortNumeric.length} value(s) are under 5 digits — likely because the source spreadsheet stored this column as a number` });
    });

    // 8. Cross-field validation: DOB vs stated Age
    const dobCol = dateCols.find((c) => colTypes[c]?.subtype === "dob");
    const ageCol = numericCols.find((c) => colTypes[c]?.subtype === "age");
    if (dobCol && ageCol) {
      const refCol = trendCol || dobCol;
      let mismatches = 0;
      rawRows.forEach((r) => {
        const dob = parseDateLoose(r[dobCol]);
        const stated = parseFloat(r[ageCol]);
        if (!dob || isNaN(stated)) return;
        const ref = refCol && !isMissing(r[refCol]) ? (parseDateLoose(r[refCol]) || now) : now;
        const computed = ageAt(dob, ref);
        if (Math.abs(computed - stated) > 1) mismatches++;
      });
      if (mismatches > 0) found.push({ level: "error", label: `${ageCol} doesn't match ${dobCol}`, detail: `${mismatches} row(s) where stated age differs from age computed from date of birth by more than 1 year` });
    }

    // 9. Cross-field validation: DOB after event date (impossible)
    const eventCol = dateCols.find((c) => colTypes[c]?.subtype === "event");
    if (dobCol && eventCol) {
      let bad = 0;
      rawRows.forEach((r) => {
        const dob = parseDateLoose(r[dobCol]);
        const ev = parseDateLoose(r[eventCol]);
        if (dob && ev && dob > ev) bad++;
      });
      if (bad > 0) found.push({ level: "error", label: `${dobCol} occurs after ${eventCol}`, detail: `${bad} row(s) have a birth date later than the event date — likely swapped or mistyped fields` });
    }

    // 10. Future event dates
    if (eventCol) {
      const future = rawRows.filter((r) => { const d = parseDateLoose(r[eventCol]); return d && d > now; });
      if (future.length > 0) found.push({ level: "warn", label: `Future dates in "${eventCol}"`, detail: `${future.length} row(s) are dated after today — verify these aren't data-entry errors` });
    }

    // 11. Inconsistent categories (fuzzy match)
    categoricalCols.forEach((c) => {
      const counts = {};
      cleanedRows.forEach((r) => { const v = r[c]; if (!isMissing(v)) counts[v] = (counts[v] || 0) + 1; });
      const levels = Object.entries(counts).map(([level, n]) => ({ level, n }));
      const inconsistencies = findCategoryInconsistencies(levels);
      inconsistencies.forEach((inc) => {
        found.push({ level: "info", label: `Possible inconsistent coding in "${c}"`, detail: `"${inc.pair[0]}" and "${inc.pair[1]}" (${inc.reason}${inc.similarity ? `, ${inc.similarity}% similar` : ""})` });
      });
    });

    return found;
  }, [rawRows, columns, colTypes, identifierCols, emailCols, phoneCols, zipCols, dateCols, numericCols, categoricalCols, cleanedRows, trendCol]);

  const descriptives = useMemo(() => {
    const out = { numeric: {}, categorical: {} };
    numericCols.forEach((c) => {
      const vals = cleanedRows.map((r) => parseFloat(r[c])).filter((v) => !isNaN(v));
      if (vals.length === 0) return;
      const m = mean(vals);
      out.numeric[c] = { n: vals.length, mean: m, sd: stddev(vals, m), median: median(vals), q1: quantile(vals, 0.25), q3: quantile(vals, 0.75), min: minOf(vals), max: maxOf(vals) };
    });
    categoricalCols.forEach((c) => {
      const counts = {}; let total = 0;
      cleanedRows.forEach((r) => { const v = r[c]; if (isMissing(v)) return; counts[v] = (counts[v] || 0) + 1; total++; });
      out.categorical[c] = { total, levels: Object.entries(counts).sort((a, b) => b[1] - a[1]).map(([level, n]) => ({ level, n, pct: ((n / total) * 100).toFixed(1) })) };
    });
    return out;
  }, [cleanedRows, numericCols, categoricalCols]);

  const table1 = useMemo(() => {
    if (!groupVar) return null;
    const groups = descriptives.categorical[groupVar]?.levels.map((l) => l.level) || [];
    if (groups.length === 0 || groups.length > 6) return null;
    const rowsOut = [];
    numericCols.forEach((c) => {
      if (c === groupVar) return;
      const byGroup = {};
      groups.forEach((g) => {
        const vals = cleanedRows.filter((r) => String(r[groupVar]) === String(g)).map((r) => parseFloat(r[c])).filter((v) => !isNaN(v));
        byGroup[g] = vals.length ? { mean: mean(vals), sd: stddev(vals, mean(vals)) } : null;
      });
      rowsOut.push({ variable: c, type: "numeric", byGroup });
    });
    categoricalCols.forEach((c) => {
      if (c === groupVar) return;
      const levels = descriptives.categorical[c]?.levels.map((l) => l.level) || [];
      levels.slice(0, 5).forEach((lvl) => {
        const byGroup = {};
        groups.forEach((g) => {
          const denom = cleanedRows.filter((r) => String(r[groupVar]) === String(g)).length;
          const num = cleanedRows.filter((r) => String(r[groupVar]) === String(g) && String(r[c]) === String(lvl)).length;
          byGroup[g] = denom ? { n: num, pct: ((num / denom) * 100).toFixed(1) } : null;
        });
        rowsOut.push({ variable: `${c}: ${lvl}`, type: "categorical", byGroup });
      });
    });
    return { groups, rows: rowsOut };
  }, [groupVar, cleanedRows, numericCols, categoricalCols, descriptives]);

  const trend = useMemo(() => {
    if (!trendCol) return null;
    const counts = {};
    cleanedRows.forEach((r) => { const d = parseDateLoose(r[trendCol]); if (!d) return; const key = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`; counts[key] = (counts[key] || 0) + 1; });
    const sortedKeys = Object.keys(counts).sort();
    return { column: trendCol, data: sortedKeys.map((k) => ({ month: k, count: counts[k] })) };
  }, [cleanedRows, trendCol]);

  // ---- Data quality scorecard ----
  const qualityScore = useMemo(() => {
    if (!rawRows) return null;
    const totalCells = rawRows.length * columns.length;
    const missingCells = columns.reduce((sum, c) => sum + rawRows.filter((r) => isMissing(r[c])).length, 0);
    const completeness = totalCells ? 1 - missingCells / totalCells : 1;

    const errorIssues = issues.filter((i) => i.level === "error").length;
    const warnIssues = issues.filter((i) => i.level === "warn").length;
    const validity = Math.max(0, 1 - (errorIssues * 2 + warnIssues) / Math.max(10, columns.length * 2));

    const dupIdCount = issues.filter((i) => i.label.startsWith("Duplicate")).length;
    const uniqueness = Math.max(0, 1 - dupIdCount * 0.15);

    const inconsistentCount = issues.filter((i) => i.label.includes("inconsistent coding")).length;
    const consistency = Math.max(0, 1 - inconsistentCount * 0.1);

    return { completeness, validity, uniqueness, consistency };
  }, [rawRows, columns, issues]);

  const loadParsed = (parsed, name, kind) => {
    setRawRows(parsed.rows); setSheetNames(parsed.sheetNames || null); setActiveSheet(parsed.activeSheet || null);
    setWorkbook(parsed.workbook || null); setFileName(name); setFileKind(kind); setGroupVar(null);
  };

  const handleFile = useCallback((file) => {
    setError(null);
    const ext = file.name.split(".").pop().toLowerCase();
    setStage("analyzing");
    if (ext === "csv") {
      parseCSV(file).then((parsed) => { if (!parsed.rows?.length) throw new Error("empty"); loadParsed(parsed, file.name, "csv"); setTimeout(() => setStage("results"), 500); })
        .catch(() => { setError("This CSV appears to be empty or couldn't be read."); setStage("upload"); });
    } else if (ext === "xlsx" || ext === "xls") {
      parseExcel(file).then((parsed) => { if (!parsed.rows?.length) throw new Error("empty"); loadParsed(parsed, file.name, "excel"); setTimeout(() => setStage("results"), 500); })
        .catch((err) => { setError("Couldn't read this Excel file: " + err.message); setStage("upload"); });
    } else {
      setError("Unsupported file type. Please upload a .csv, .xlsx, or .xls file.");
      setStage("upload");
    }
  }, []);

  const switchSheet = (name) => {
    if (!workbook) return;
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[name], { defval: "", raw: true });
    setRawRows(rows); setActiveSheet(name); setGroupVar(null);
  };
  const loadSample = () => {
    setError(null); setStage("analyzing");
    setTimeout(() => { loadParsed({ rows: generateSampleData(), sheetNames: null, activeSheet: null }, "sample_overdose_surveillance.csv", "csv"); setStage("results"); }, 550);
  };
  const onDrop = (e) => { e.preventDefault(); const file = e.dataTransfer.files?.[0]; if (file) handleFile(file); };

  const downloadReport = () => {
    let out = `SurvIQ Summary Report\nFile: ${fileName}${activeSheet ? ` (sheet: ${activeSheet})` : ""}\nGenerated: ${new Date().toLocaleString()}\nRows: ${rawRows.length}  Columns: ${columns.length}\n\n`;
    if (qualityScore) out += `DATA QUALITY SCORE\nCompleteness: ${(qualityScore.completeness * 100).toFixed(1)}%\nValidity: ${(qualityScore.validity * 100).toFixed(1)}%\nUniqueness: ${(qualityScore.uniqueness * 100).toFixed(1)}%\nConsistency: ${(qualityScore.consistency * 100).toFixed(1)}%\n\n`;
    out += `VALIDATION ISSUES (${issues.length})\n`;
    issues.forEach((i) => (out += `- [${i.level.toUpperCase()}] ${i.label}: ${i.detail}\n`));
    out += `\nDESCRIPTIVE STATISTICS\n`;
    Object.entries(descriptives.numeric).forEach(([c, d]) => { out += `${c}: n=${d.n}, mean=${d.mean.toFixed(2)}, SD=${d.sd.toFixed(2)}, median=${d.median.toFixed(1)}, IQR=${d.q1.toFixed(1)}-${d.q3.toFixed(1)}, range=${d.min}-${d.max}\n`; });
    Object.entries(descriptives.categorical).forEach(([c, d]) => { out += `${c} (n=${d.total}):\n`; d.levels.forEach((l) => (out += `   ${l.level}: ${l.n} (${l.pct}%)\n`)); });
    if (trend) { out += `\nTREND (by ${trend.column}, monthly counts)\n`; trend.data.forEach((t) => (out += `${t.month}: ${t.count}\n`)); }
    const blob = new Blob([out], { type: "text/plain" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a"); a.href = url; a.download = "surviq_summary_report.txt"; a.click();
    URL.revokeObjectURL(url);
  };
  const reset = () => { setRawRows(null); setFileName(""); setFileKind(null); setSheetNames(null); setActiveSheet(null); setWorkbook(null); setGroupVar(null); setTrendCol(null); setError(null); setStage("upload"); };

  return (
    <div style={{ minHeight: "100vh", background: P.bg, fontFamily: "'IBM Plex Sans', ui-sans-serif, system-ui", color: P.ink }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500;600&family=Source+Serif+4:opsz,wght@8..60,500;8..60,600&display=swap');
        * { box-sizing: border-box; }
        .mono { font-family: 'IBM Plex Mono', monospace; }
        .serif { font-family: 'Source Serif 4', Georgia, serif; }
        ::selection { background: ${P.tealSoft}; }
        .fade-in { animation: fadeIn 0.4s ease both; }
        @keyframes fadeIn { from { opacity: 0; transform: translateY(4px); } to { opacity: 1; transform: translateY(0); } }
        input, select, button { font-family: inherit; }
      `}</style>

      <header style={{ padding: "17px 32px", display: "flex", alignItems: "center", justifyContent: "space-between", borderBottom: `1px solid ${P.border}`, background: P.panel }}>
        <div style={{ display: "flex", alignItems: "center", gap: 11 }}>
          <div style={{ width: 30, height: 30, borderRadius: 6, background: P.navy, display: "flex", alignItems: "center", justifyContent: "center" }}>
            <span className="serif" style={{ color: "#fff", fontSize: 15, fontWeight: 600 }}>S</span>
          </div>
          <span className="serif" style={{ fontSize: 20, fontWeight: 600, letterSpacing: "-0.01em" }}>SurvIQ</span>
          <span className="mono" style={{ fontSize: 10.5, color: P.mutedLight, letterSpacing: "0.05em", marginLeft: 2, paddingTop: 2 }}>SURVEILLANCE ANALYTICS</span>
        </div>
        {rawRows && <button onClick={reset} className="mono" style={{ fontSize: 11.5, background: "transparent", border: `1px solid ${P.borderStrong}`, borderRadius: 6, padding: "7px 13px", color: P.ink, cursor: "pointer" }}>NEW UPLOAD</button>}
      </header>

      {stage === "upload" && <UploadScreen onFile={handleFile} onSample={loadSample} onDrop={onDrop} error={error} />}
      {stage === "analyzing" && <AnalyzingScreen fileName={fileName} />}
      {stage === "results" && rawRows && (
        <ResultsScreen
          rawRows={rawRows} fileName={fileName} fileKind={fileKind} sheetNames={sheetNames} activeSheet={activeSheet} onSwitchSheet={switchSheet}
          columns={columns} colTypes={colTypes} issues={issues} qualityScore={qualityScore}
          descriptives={descriptives} numericCols={numericCols} categoricalCols={categoricalCols}
          identifierCols={identifierCols} phoneCols={phoneCols} emailCols={emailCols} zipCols={zipCols}
          groupVar={groupVar} setGroupVar={setGroupVar} table1={table1}
          trend={trend} dateCols={dateCols} trendCol={trendCol} setTrendCol={setTrendCol}
          onDownload={downloadReport}
        />
      )}
    </div>
  );
}

// =====================================================================
// Screens
// =====================================================================
function UploadScreen({ onFile, onSample, onDrop, error }) {
  const [dragOver, setDragOver] = useState(false);
  return (
    <div style={{ maxWidth: 780, margin: "0 auto", padding: "68px 24px 40px" }}>
      <div className="fade-in" style={{ textAlign: "center", marginBottom: 40 }}>
        <div className="mono" style={{ fontSize: 11, letterSpacing: "0.1em", color: P.teal, marginBottom: 14, fontWeight: 600 }}>NEW PROJECT</div>
        <h1 className="serif" style={{ fontSize: 36, margin: "0 0 14px", lineHeight: 1.2, fontWeight: 600 }}>Upload your dataset. Get a validated first-pass analysis in seconds.</h1>
        <p style={{ fontSize: 15, color: P.muted, maxWidth: 520, margin: "0 auto", lineHeight: 1.65 }}>
          CSV and Excel supported. SurvIQ classifies every column by what it actually represents —
          identifier, date, contact info, measurement, or category — then validates, cross-checks, and summarizes accordingly.
        </p>
      </div>
      <label onDragOver={(e) => { e.preventDefault(); setDragOver(true); }} onDragLeave={() => setDragOver(false)} onDrop={(e) => { setDragOver(false); onDrop(e); }}
        style={{ display: "block", border: `1.5px dashed ${dragOver ? P.teal : P.borderStrong}`, borderRadius: 12, padding: "42px 24px", textAlign: "center", cursor: "pointer", background: dragOver ? P.tealSoft : P.panel, transition: "all 0.15s ease" }}>
        <input type="file" accept=".csv,.xlsx,.xls" style={{ display: "none" }} onChange={(e) => e.target.files?.[0] && onFile(e.target.files[0])} />
        <Upload size={24} color={P.navy} style={{ marginBottom: 12 }} />
        <div style={{ fontSize: 15, fontWeight: 600, marginBottom: 5 }}>Drop a file here, or click to browse</div>
        <div className="mono" style={{ fontSize: 11.5, color: P.muted }}>.csv · .xlsx · .xls — line-list or case-report format</div>
      </label>
      {error && <div style={{ marginTop: 16, padding: "12px 16px", background: P.redSoft, border: `1px solid ${P.red}44`, borderRadius: 8, color: P.red, fontSize: 13.5 }}>{error}</div>}
      <div style={{ textAlign: "center", marginTop: 20 }}>
        <button onClick={onSample} className="mono" style={{ fontSize: 12, background: "none", border: "none", color: P.teal, cursor: "pointer", textDecoration: "underline", textUnderlineOffset: 3 }}>No file handy? Try a sample overdose surveillance dataset →</button>
      </div>
      <div style={{ display: "flex", gap: 24, marginTop: 52, flexWrap: "wrap", justifyContent: "center", borderTop: `1px solid ${P.border}`, paddingTop: 28 }}>
        {[
          { icon: <ShieldCheck size={16} />, label: "Semantic validation", desc: "Cross-field checks, format rules, fuzzy category matching" },
          { icon: <Table2 size={16} />, label: "Descriptives + Table 1", desc: "Mean, SD, median, IQR, group comparison" },
          { icon: <TrendingUp size={16} />, label: "Trend charts", desc: "Auto-selected event-date field, override available" },
        ].map((f) => (
          <div key={f.label} style={{ maxWidth: 210, textAlign: "left" }}>
            <div style={{ display: "flex", alignItems: "center", gap: 7, color: P.navy, marginBottom: 5 }}>{f.icon}<span style={{ fontSize: 13, fontWeight: 600, color: P.ink }}>{f.label}</span></div>
            <div style={{ fontSize: 12, color: P.muted, lineHeight: 1.5 }}>{f.desc}</div>
          </div>
        ))}
      </div>
    </div>
  );
}

function AnalyzingScreen({ fileName }) {
  const steps = ["Reading file", "Classifying variables by semantic type", "Running validation & cross-field checks", "Computing descriptive statistics"];
  const [active, setActive] = useState(0);
  useEffect(() => { const t = setInterval(() => setActive((a) => Math.min(a + 1, steps.length - 1)), 160); return () => clearInterval(t); }, []);
  return (
    <div style={{ maxWidth: 480, margin: "0 auto", padding: "130px 24px", textAlign: "center" }}>
      <div className="mono" style={{ fontSize: 12, color: P.muted, marginBottom: 22 }}>{fileName}</div>
      <div style={{ display: "flex", flexDirection: "column", gap: 11, alignItems: "flex-start", margin: "0 auto", width: "fit-content" }}>
        {steps.map((s, i) => (
          <div key={s} style={{ display: "flex", alignItems: "center", gap: 10, fontSize: 13.5, color: i <= active ? P.ink : P.mutedLight }}>
            {i < active ? <CheckCircle2 size={15} color={P.teal} /> : <div style={{ width: 7, height: 7, borderRadius: 999, background: i === active ? P.teal : P.border }} />}
            {s}
          </div>
        ))}
      </div>
    </div>
  );
}

function ResultsScreen({ rawRows, fileName, fileKind, sheetNames, activeSheet, onSwitchSheet, columns, colTypes, issues, qualityScore, descriptives, numericCols, categoricalCols, identifierCols, phoneCols, emailCols, zipCols, groupVar, setGroupVar, table1, trend, dateCols, trendCol, setTrendCol, onDownload }) {
  const errorCount = issues.filter((i) => i.level === "error").length;
  const warnCount = issues.filter((i) => i.level === "warn").length;
  const infoCount = issues.filter((i) => i.level === "info").length;

  return (
    <div style={{ maxWidth: 1060, margin: "0 auto", padding: "30px 28px 90px" }}>
      <div className="fade-in" style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-end", marginBottom: 22, flexWrap: "wrap", gap: 14 }}>
        <div>
          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
            {fileKind === "excel" ? <FileSpreadsheet size={13} color={P.muted} /> : <FileText size={13} color={P.muted} />}
            <span className="mono" style={{ fontSize: 11, color: P.muted }}>{fileName}</span>
          </div>
          <h2 className="serif" style={{ fontSize: 25, margin: "5px 0 0", fontWeight: 600 }}>{rawRows.length.toLocaleString()} rows · {columns.length} variables</h2>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          {sheetNames && sheetNames.length > 1 && (
            <div style={{ display: "flex", alignItems: "center", gap: 7 }}>
              <Layers size={14} color={P.muted} />
              <select value={activeSheet} onChange={(e) => onSwitchSheet(e.target.value)} className="mono" style={{ fontSize: 12, padding: "7px 9px", borderRadius: 6, border: `1px solid ${P.borderStrong}`, background: P.panel, color: P.ink }}>
                {sheetNames.map((s) => <option key={s} value={s}>{s}</option>)}
              </select>
            </div>
          )}
          <button onClick={onDownload} style={{ display: "flex", alignItems: "center", gap: 8, background: P.navy, color: "#fff", border: "none", borderRadius: 7, padding: "10px 16px", fontSize: 13, fontWeight: 600, cursor: "pointer" }}>
            <Download size={14} /> Export summary
          </button>
        </div>
      </div>

      {qualityScore && (
        <Panel title="Data quality scorecard" eyebrow="COMPUTED FROM THIS FILE">
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(140px, 1fr))", gap: 16, marginBottom: 12 }}>
            <ScoreTile label="Completeness" value={qualityScore.completeness} />
            <ScoreTile label="Validity" value={qualityScore.validity} />
            <ScoreTile label="Uniqueness" value={qualityScore.uniqueness} />
            <ScoreTile label="Consistency" value={qualityScore.consistency} />
          </div>
          <div style={{ display: "flex", gap: 8, fontSize: 12, color: P.muted, background: P.bg, borderRadius: 8, padding: "10px 12px", lineHeight: 1.5 }}>
            <Info size={14} style={{ flexShrink: 0, marginTop: 2 }} />
            <span>These four dimensions are computed directly from patterns in this file. Two dimensions sometimes reported elsewhere — <em>accuracy against ground truth</em> and <em>timeliness against a reporting-lag benchmark</em> — require an external reference this file alone can't provide, so they're intentionally left out rather than estimated.</span>
          </div>
        </Panel>
      )}

      <Panel title="Validation issues" eyebrow={`${issues.length} found`}>
        {issues.length === 0 ? (
          <div style={{ display: "flex", alignItems: "center", gap: 8, color: P.teal, fontSize: 14 }}><CheckCircle2 size={17} /> No obvious problems detected.</div>
        ) : (
          <>
            <div style={{ display: "flex", gap: 16, marginBottom: 15, fontSize: 12.5 }}>
              {errorCount > 0 && <Tag color={P.red} label={`${errorCount} needs review`} />}
              {warnCount > 0 && <Tag color={P.amber} label={`${warnCount} warning${warnCount === 1 ? "" : "s"}`} />}
              {infoCount > 0 && <Tag color={P.mutedLight} label={`${infoCount} note${infoCount === 1 ? "" : "s"}`} />}
            </div>
            <div style={{ display: "flex", flexDirection: "column", gap: 9 }}>
              {issues.map((iss, idx) => (
                <div key={idx} style={{ display: "flex", gap: 10, alignItems: "flex-start", fontSize: 13.5 }}>
                  <span style={{ color: iss.level === "error" ? P.red : iss.level === "warn" ? P.amber : P.mutedLight, marginTop: 2 }}>
                    {iss.level === "error" ? <AlertTriangle size={14} /> : <ChevronRight size={14} />}
                  </span>
                  <div><span style={{ fontWeight: 600 }}>{iss.label}</span><span style={{ color: P.muted }}> — {iss.detail}</span>
                    {iss.fixable && <span className="mono" style={{ color: P.teal, marginLeft: 6, fontSize: 11 }}>· excluded from analysis as missing</span>}
                  </div>
                </div>
              ))}
            </div>
          </>
        )}
      </Panel>

      <Panel title="Variables detected" eyebrow={`${columns.length} columns classified by semantic type`}>
        <div style={{ display: "flex", flexWrap: "wrap", gap: 8 }}>
          {columns.map((c) => (
            <div key={c} style={{ display: "flex", alignItems: "center", gap: 6, border: `1px solid ${P.border}`, borderRadius: 6, padding: "6px 10px", fontSize: 12.5, background: P.panel }}>
              <span style={{ fontWeight: 600 }}>{c}</span>
              <span className="mono" style={{ fontSize: 10, color: P.teal, background: P.tealSoft, padding: "1px 6px", borderRadius: 4 }}>
                {TYPE_LABEL[colTypes[c]?.type] || "unknown"}{colTypes[c]?.subtype ? ` · ${colTypes[c].subtype}` : ""}
              </span>
            </div>
          ))}
        </div>
        {(identifierCols.length > 0 || phoneCols.length > 0 || emailCols.length > 0 || zipCols.length > 0) && (
          <div style={{ marginTop: 12, fontSize: 12, color: P.mutedLight }}>
            Identifiers, contact fields, and postal codes are recognized as their own types and excluded from numeric statistics and trend detection — they're validated (format, duplicates) instead of averaged.
          </div>
        )}
      </Panel>

      <Panel title="Descriptive statistics" eyebrow="Numeric measurements only">
        {numericCols.length === 0 ? <EmptyNote text="No numeric measurement variables detected." /> : (
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
            <thead><tr style={{ borderBottom: `1.5px solid ${P.ink}` }}>{["Variable", "n", "Mean", "SD", "Median", "IQR", "Range"].map((h) => <th key={h} className="mono" style={{ textAlign: h === "Variable" ? "left" : "right", padding: "8px 10px", color: P.muted, fontWeight: 600, fontSize: 10.5 }}>{h.toUpperCase()}</th>)}</tr></thead>
            <tbody>
              {Object.entries(descriptives.numeric).map(([c, d]) => (
                <tr key={c} style={{ borderBottom: `1px solid ${P.border}` }}>
                  <td style={{ padding: "9px 10px", fontWeight: 600 }}>{c}</td>
                  <td className="mono" style={{ padding: "9px 10px", textAlign: "right" }}>{d.n}</td>
                  <td className="mono" style={{ padding: "9px 10px", textAlign: "right" }}>{d.mean.toFixed(1)}</td>
                  <td className="mono" style={{ padding: "9px 10px", textAlign: "right" }}>{d.sd.toFixed(1)}</td>
                  <td className="mono" style={{ padding: "9px 10px", textAlign: "right" }}>{d.median.toFixed(1)}</td>
                  <td className="mono" style={{ padding: "9px 10px", textAlign: "right" }}>{d.q1.toFixed(0)}–{d.q3.toFixed(0)}</td>
                  <td className="mono" style={{ padding: "9px 10px", textAlign: "right" }}>{d.min}–{d.max}</td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </Panel>

      <Panel title="Categorical frequencies" eyebrow="Counts and percentages">
        {categoricalCols.length === 0 ? <EmptyNote text="No categorical variables detected." /> : (
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(230px, 1fr))", gap: 22 }}>
            {categoricalCols.map((c) => (
              <div key={c}>
                <div style={{ fontWeight: 600, fontSize: 13.5, marginBottom: 9 }}>{c}</div>
                {descriptives.categorical[c].levels.slice(0, 6).map((l) => (
                  <div key={l.level} style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 6 }}>
                    <div style={{ fontSize: 12, width: 92, color: P.muted, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{String(l.level)}</div>
                    <div style={{ flex: 1, background: P.tealSoft, borderRadius: 4, height: 7, overflow: "hidden" }}><div style={{ width: `${l.pct}%`, background: P.teal, height: "100%" }} /></div>
                    <div className="mono" style={{ fontSize: 11, width: 66, textAlign: "right", color: P.ink }}>{l.n} ({l.pct}%)</div>
                  </div>
                ))}
              </div>
            ))}
          </div>
        )}
      </Panel>

      <Panel title="Table 1" eyebrow="Baseline characteristics by group">
        {categoricalCols.length === 0 ? <EmptyNote text="Add a categorical variable to compare groups." /> : (
          <>
            <div style={{ marginBottom: 16, display: "flex", alignItems: "center", gap: 10 }}>
              <span style={{ fontSize: 13, color: P.muted }}>Group by:</span>
              <select value={groupVar || ""} onChange={(e) => setGroupVar(e.target.value || null)} style={{ fontSize: 13, padding: "6px 10px", borderRadius: 6, border: `1px solid ${P.borderStrong}`, background: P.panel, color: P.ink }}>
                <option value="">Select a grouping variable…</option>
                {categoricalCols.map((c) => <option key={c} value={c}>{c}</option>)}
              </select>
            </div>
            {!table1 ? <EmptyNote text={groupVar ? "This variable has too many categories to group by (max 6)." : "Choose a variable above to generate Table 1."} /> : (
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
                <thead><tr style={{ borderBottom: `1.5px solid ${P.ink}` }}><th className="mono" style={{ textAlign: "left", padding: "8px 10px", fontSize: 10.5, color: P.muted }}>VARIABLE</th>{table1.groups.map((g) => <th key={g} className="mono" style={{ textAlign: "right", padding: "8px 10px", fontSize: 10.5, color: P.muted }}>{String(g).toUpperCase()}</th>)}</tr></thead>
                <tbody>
                  {table1.rows.map((r, idx) => (
                    <tr key={idx} style={{ borderBottom: `1px solid ${P.border}` }}>
                      <td style={{ padding: "8px 10px" }}>{r.variable}</td>
                      {table1.groups.map((g) => { const v = r.byGroup[g]; return <td key={g} className="mono" style={{ padding: "8px 10px", textAlign: "right" }}>{!v ? "—" : r.type === "numeric" ? `${v.mean.toFixed(1)} ± ${v.sd.toFixed(1)}` : `${v.n} (${v.pct}%)`}</td>; })}
                    </tr>
                  ))}
                </tbody>
              </table>
            )}
          </>
        )}
      </Panel>

      <Panel title="Trend over time" eyebrow={trend ? `Monthly counts from ${trend.column}` : "No suitable date field found"}>
        {dateCols.length > 1 && (
          <div style={{ marginBottom: 14, display: "flex", alignItems: "center", gap: 10 }}>
            <span style={{ fontSize: 13, color: P.muted }}>Date field:</span>
            <select value={trendCol || ""} onChange={(e) => setTrendCol(e.target.value || null)} style={{ fontSize: 13, padding: "6px 10px", borderRadius: 6, border: `1px solid ${P.borderStrong}`, background: P.panel, color: P.ink }}>
              {dateCols.map((c) => <option key={c} value={c}>{c}{colTypes[c]?.subtype === "dob" ? " (birth date — unusual choice for a trend)" : ""}</option>)}
            </select>
          </div>
        )}
        {!trend ? <EmptyNote text="Upload a dataset with a plausible event-date column to see trends over time." /> : (
          <div style={{ height: 260 }}>
            <ResponsiveContainer width="100%" height="100%">
              <AreaChart data={trend.data} margin={{ top: 8, right: 16, left: -12, bottom: 0 }}>
                <defs><linearGradient id="areaFill" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stopColor={P.navy} stopOpacity={0.18} /><stop offset="100%" stopColor={P.navy} stopOpacity={0} /></linearGradient></defs>
                <CartesianGrid strokeDasharray="2 4" stroke={P.border} />
                <XAxis dataKey="month" tick={{ fontSize: 11, fill: P.muted }} />
                <YAxis tick={{ fontSize: 11, fill: P.muted }} allowDecimals={false} />
                <Tooltip contentStyle={{ fontSize: 12.5, borderRadius: 8, border: `1px solid ${P.border}` }} />
                <Area type="monotone" dataKey="count" stroke={P.navy} strokeWidth={2} fill="url(#areaFill)" dot={{ r: 2.5, fill: P.navy }} />
              </AreaChart>
            </ResponsiveContainer>
          </div>
        )}
      </Panel>

      {categoricalCols.length > 0 && (
        <Panel title="Distribution snapshot" eyebrow={`Top categories in ${categoricalCols[0]}`}>
          <div style={{ height: 240 }}>
            <ResponsiveContainer width="100%" height="100%">
              <BarChart data={descriptives.categorical[categoricalCols[0]].levels.slice(0, 8)} margin={{ top: 8, right: 16, left: -12, bottom: 0 }}>
                <CartesianGrid strokeDasharray="2 4" stroke={P.border} />
                <XAxis dataKey="level" tick={{ fontSize: 11, fill: P.muted }} />
                <YAxis tick={{ fontSize: 11, fill: P.muted }} allowDecimals={false} />
                <Tooltip contentStyle={{ fontSize: 12.5, borderRadius: 8, border: `1px solid ${P.border}` }} />
                <Bar dataKey="n" radius={[4, 4, 0, 0]}>{descriptives.categorical[categoricalCols[0]].levels.slice(0, 8).map((_, i) => <Cell key={i} fill={CHART_COLORS[i % CHART_COLORS.length]} />)}</Bar>
              </BarChart>
            </ResponsiveContainer>
          </div>
        </Panel>
      )}
    </div>
  );
}

function ScoreTile({ label, value }) {
  const pct = Math.round(value * 100);
  const color = pct >= 90 ? P.teal : pct >= 75 ? P.amber : P.red;
  return (
    <div style={{ border: `1px solid ${P.border}`, borderRadius: 8, padding: "12px 14px" }}>
      <div className="mono" style={{ fontSize: 10, color: P.mutedLight, marginBottom: 4 }}>{label.toUpperCase()}</div>
      <div className="serif" style={{ fontSize: 22, fontWeight: 600, color }}>{pct}%</div>
    </div>
  );
}
function Panel({ title, eyebrow, children }) {
  return (
    <div className="fade-in" style={{ background: P.panel, border: `1px solid ${P.border}`, borderRadius: 10, padding: "22px 24px", marginBottom: 16 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline", marginBottom: 15, flexWrap: "wrap", gap: 6 }}>
        <h3 className="serif" style={{ fontSize: 17, margin: 0, fontWeight: 600 }}>{title}</h3>
        {eyebrow && <span className="mono" style={{ fontSize: 10.5, color: P.mutedLight }}>{eyebrow}</span>}
      </div>
      {children}
    </div>
  );
}
function Tag({ color, label }) { return <span style={{ display: "flex", alignItems: "center", gap: 5, color }}><span style={{ width: 6, height: 6, borderRadius: 999, background: color }} />{label}</span>; }
function EmptyNote({ text }) { return <div style={{ fontSize: 13, color: P.mutedLight, fontStyle: "italic" }}>{text}</div>; }