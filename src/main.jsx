import React, { useEffect, useMemo, useRef, useState } from "react";
import { createRoot } from "react-dom/client";
import * as XLSX from "xlsx";
import {
  Activity,
  BadgeIndianRupee,
  Check,
  Download,
  LineChart,
  LogOut,
  Plus,
  RotateCcw,
  Save,
  Trash2,
  Upload,
  X
} from "lucide-react";
import { supabase } from "./supabaseClient";
import "./styles.css";

const STORAGE_KEY = "gtt-darvas-tracker-v1";
const IGNORE_PCT = 0.0275;
const TARGET_PCT = 0.06;
const ALL_STOCKS = "__ALL_STOCKS__";

const excelInputHeaders = [
  "Symbol",
  "Week Start",
  "Week End",
  "Monday High",
  "Tuesday High",
  "Wednesday High",
  "Thursday High",
  "Friday High",
  "Shares",
  "Execution",
  "Execution Week",
  "Next Week Open",
  "Next Week High",
  "Target Hit"
];

const executionAliases = {
  "placed / pending": "pending",
  pending: "pending",
  placed: "pending",
  filled: "filled",
  "filled @ trigger": "filled",
  "gap up unable": "gap",
  gap: "gap",
  "unable to trigger": "unable",
  unable: "unable"
};

const monthLookup = {
  jan: 0,
  feb: 1,
  mar: 2,
  apr: 3,
  may: 4,
  jun: 5,
  jul: 6,
  aug: 7,
  sep: 8,
  oct: 9,
  nov: 10,
  dec: 11
};

const seedWeeks = [
  {
    id: crypto.randomUUID(),
    symbol: "EXAMPLE",
    weekStart: "2023-01-02",
    weekEnd: "2023-01-06",
    monHigh: 4025,
    tueHigh: 4072,
    wedHigh: 4098,
    thuHigh: 4056,
    friHigh: 4105,
    shares: 1,
    execution: "filled",
    executionWeek: "2023-01-09",
    nextWeekOpen: 4080,
    nextWeekHigh: 4140
  },
  {
    id: crypto.randomUUID(),
    symbol: "EXAMPLE",
    weekStart: "2023-01-09",
    weekEnd: "2023-01-13",
    monHigh: 4110,
    tueHigh: 4120,
    wedHigh: 4132,
    thuHigh: 4144,
    friHigh: 4140,
    shares: 1,
    execution: "pending",
    executionWeek: "2023-01-16",
    nextWeekOpen: "",
    nextWeekHigh: ""
  },
  {
    id: crypto.randomUUID(),
    symbol: "DEMO2",
    weekStart: "2023-01-02",
    weekEnd: "2023-01-06",
    monHigh: 1180,
    tueHigh: 1192,
    wedHigh: 1205,
    thuHigh: 1198,
    friHigh: 1210,
    shares: 2,
    execution: "filled",
    executionWeek: "2023-01-09",
    nextWeekOpen: 1200,
    nextWeekHigh: 1225
  },
  {
    id: crypto.randomUUID(),
    symbol: "DEMO2",
    weekStart: "2023-01-09",
    weekEnd: "2023-01-13",
    monHigh: 1215,
    tueHigh: 1222,
    wedHigh: 1230,
    thuHigh: 1228,
    friHigh: 1235,
    shares: 2,
    execution: "pending",
    executionWeek: "2023-01-16",
    nextWeekOpen: "",
    nextWeekHigh: ""
  }
];

function toNumber(value) {
  const number = Number(String(value ?? "").replaceAll(",", ""));
  return Number.isFinite(number) ? number : 0;
}

function money(value) {
  if (!Number.isFinite(value)) return "-";
  return new Intl.NumberFormat("en-IN", {
    style: "currency",
    currency: "INR",
    maximumFractionDigits: 2
  }).format(value);
}

function compactMoney(value) {
  if (!Number.isFinite(value)) return "-";
  return new Intl.NumberFormat("en-IN", {
    maximumFractionDigits: 0
  }).format(value);
}

function getWeeklyHigh(week) {
  return Math.max(
    toNumber(week.monHigh),
    toNumber(week.tueHigh),
    toNumber(week.wedHigh),
    toNumber(week.thuHigh),
    toNumber(week.friHigh)
  );
}

function estimateShares(price) {
  if (!price) return 1;
  return price <= 2700 ? 2 : 1;
}

function normalizeWeek(week) {
  const weeklyHigh = getWeeklyHigh(week);
  return {
    ...week,
    shares: toNumber(week.shares) || estimateShares(weeklyHigh)
  };
}

function readCell(row, names) {
  for (const name of names) {
    if (row[name] !== undefined && row[name] !== null && row[name] !== "") return row[name];
  }
  return "";
}

function normalizeDateCell(value) {
  if (!value) return "";
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value.toISOString().slice(0, 10);
  }
  if (typeof value === "number") {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (!parsed) return "";
    return `${parsed.y}-${String(parsed.m).padStart(2, "0")}-${String(parsed.d).padStart(2, "0")}`;
  }
  const text = String(value).trim();
  const isoLike = text.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
  if (isoLike) {
    return `${isoLike[1]}-${isoLike[2].padStart(2, "0")}-${isoLike[3].padStart(2, "0")}`;
  }
  const nseLike = text.match(/^(\d{1,2})[-\s/]([A-Za-z]{3,})[-\s/](\d{2,4})$/);
  if (nseLike) {
    const month = monthLookup[nseLike[2].slice(0, 3).toLowerCase()];
    const fullYear = Number(nseLike[3]) < 100 ? 2000 + Number(nseLike[3]) : Number(nseLike[3]);
    if (month !== undefined) {
      const date = new Date(Date.UTC(fullYear, month, Number(nseLike[1])));
      return date.toISOString().slice(0, 10);
    }
  }
  const parsed = new Date(text);
  if (!Number.isNaN(parsed.getTime())) return parsed.toISOString().slice(0, 10);
  return text.slice(0, 10);
}

function parseExecution(value) {
  const key = String(value || "pending").trim().toLowerCase();
  return executionAliases[key] || "pending";
}

function parseBoolean(value) {
  if (typeof value === "boolean") return value;
  const text = String(value || "").trim().toLowerCase();
  return ["yes", "y", "true", "1", "target hit", "hit"].includes(text);
}

function workbookRowsToWeeks(rows) {
  return rows
    .map((row) => ({
      id: crypto.randomUUID(),
      symbol: readCell(row, ["Symbol", "SYMBOL", "Stock", "Ticker"]),
      weekStart: normalizeDateCell(readCell(row, ["Week Start", "WeekStart", "Start Date", "From"])),
      weekEnd: normalizeDateCell(readCell(row, ["Week End", "WeekEnd", "End Date", "To"])),
      monHigh: readCell(row, ["Monday High", "Mon High", "MON HIGH", "Mon", "Monday"]),
      tueHigh: readCell(row, ["Tuesday High", "Tue High", "TUE HIGH", "Tue", "Tuesday"]),
      wedHigh: readCell(row, ["Wednesday High", "Wed High", "WED HIGH", "Wed", "Wednesday"]),
      thuHigh: readCell(row, ["Thursday High", "Thu High", "THU HIGH", "Thu", "Thursday"]),
      friHigh: readCell(row, ["Friday High", "Fri High", "FRI HIGH", "Fri", "Friday"]),
      shares: readCell(row, ["Shares", "Qty", "Quantity"]),
      execution: parseExecution(readCell(row, ["Execution", "GTT Result", "Status"])),
      executionWeek: normalizeDateCell(readCell(row, ["Execution Week", "Next Week", "Trigger Week"])),
      nextWeekOpen: readCell(row, ["Next Week Open", "Open", "Next Open"]),
      nextWeekHigh: readCell(row, ["Next Week High", "High", "Next High"]),
      exitHit: parseBoolean(readCell(row, ["Target Hit", "Exit Hit", "Exit", "Booked"]))
    }))
    .filter((week) =>
      week.symbol ||
      week.weekStart ||
      week.weekEnd ||
      ["monHigh", "tueHigh", "wedHigh", "thuHigh", "friHigh"].some((field) => toNumber(week[field]))
    );
}

function isoDateFromOffset(date, offsetDays) {
  const next = new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate() + offsetDays));
  return next.toISOString().slice(0, 10);
}

function nseDailyRowsToWeeks(rows, fallbackSymbol) {
  const groups = new Map();

  rows.forEach((row) => {
    const dateText = normalizeDateCell(readCell(row, ["DATE", "Date", "TIMESTAMP", "Timestamp"]));
    const high = readCell(row, ["HIGH", "High", "high"]);
    if (!dateText || !high) return;

    const date = new Date(`${dateText}T00:00:00Z`);
    if (Number.isNaN(date.getTime())) return;

    const day = date.getUTCDay();
    if (day === 0 || day === 6) return;

    const daysSinceMonday = day - 1;
    const weekStart = isoDateFromOffset(date, -daysSinceMonday);
    const weekEnd = isoDateFromOffset(date, 5 - day);
    const current = groups.get(weekStart) || {
      id: crypto.randomUUID(),
      symbol: readCell(row, ["SYMBOL", "Symbol", "Ticker", "SECURITY"]) || fallbackSymbol || "",
      weekStart,
      weekEnd,
      monHigh: "",
      tueHigh: "",
      wedHigh: "",
      thuHigh: "",
      friHigh: "",
      shares: "",
      execution: "pending",
      executionWeek: "",
      nextWeekOpen: "",
      nextWeekHigh: "",
      exitHit: false
    };

    const fieldByDay = {
      1: "monHigh",
      2: "tueHigh",
      3: "wedHigh",
      4: "thuHigh",
      5: "friHigh"
    };
    current[fieldByDay[day]] = high;
    groups.set(weekStart, current);
  });

  return [...groups.values()]
    .sort((a, b) => a.weekStart.localeCompare(b.weekStart))
    .map(normalizeWeek);
}

function rowsLookLikeNseDaily(rows) {
  return rows.some((row) => readCell(row, ["DATE", "Date", "TIMESTAMP", "Timestamp"])) &&
    rows.some((row) => readCell(row, ["HIGH", "High", "high"]));
}

function weekKey(week) {
  return `${String(week.symbol || "").trim().toUpperCase()}|${week.weekStart}|${week.weekEnd}`;
}

function sortWeeks(weeks) {
  return [...weeks].sort((a, b) => {
    const symbolCompare = String(a.symbol || "").localeCompare(String(b.symbol || ""));
    if (symbolCompare) return symbolCompare;
    return String(a.weekStart || "").localeCompare(String(b.weekStart || ""));
  });
}

// App state keeps weeks in camelCase; the Supabase table uses snake_case
// columns. These two functions translate one way and back so the rest of
// the app (evaluation engine, table rendering, Excel import/export) never
// has to know the database's column naming.
function weekToDbRow(week, userId) {
  return {
    id: week.id,
    user_id: userId,
    symbol: week.symbol || "",
    week_start: week.weekStart || null,
    week_end: week.weekEnd || null,
    mon_high: String(week.monHigh ?? ""),
    tue_high: String(week.tueHigh ?? ""),
    wed_high: String(week.wedHigh ?? ""),
    thu_high: String(week.thuHigh ?? ""),
    fri_high: String(week.friHigh ?? ""),
    shares: String(week.shares ?? ""),
    execution: week.execution || "pending",
    execution_week: week.executionWeek || null,
    next_week_open: String(week.nextWeekOpen ?? ""),
    next_week_high: String(week.nextWeekHigh ?? ""),
    exit_hit: Boolean(week.exitHit)
  };
}

function dbRowToWeek(row) {
  return {
    id: row.id,
    symbol: row.symbol || "",
    weekStart: row.week_start || "",
    weekEnd: row.week_end || "",
    monHigh: row.mon_high || "",
    tueHigh: row.tue_high || "",
    wedHigh: row.wed_high || "",
    thuHigh: row.thu_high || "",
    friHigh: row.fri_high || "",
    shares: row.shares || "",
    execution: row.execution || "pending",
    executionWeek: row.execution_week || "",
    nextWeekOpen: row.next_week_open || "",
    nextWeekHigh: row.next_week_high || "",
    exitHit: Boolean(row.exit_hit)
  };
}

function weeksToInputRows(weeks) {
  return weeks.map((week) => ({
    "Symbol": week.symbol,
    "Week Start": week.weekStart,
    "Week End": week.weekEnd,
    "Monday High": week.monHigh,
    "Tuesday High": week.tueHigh,
    "Wednesday High": week.wedHigh,
    "Thursday High": week.thuHigh,
    "Friday High": week.friHigh,
    "Shares": week.shares,
    "Execution": week.execution,
    "Execution Week": week.executionWeek,
    "Next Week Open": week.nextWeekOpen,
    "Next Week High": week.nextWeekHigh,
    "Target Hit": week.exitHit ? "Yes" : "No"
  }));
}

function evaluatedToRows(evaluated) {
  return evaluated.map((week) => ({
    "Symbol": week.symbol,
    "Week Start": week.weekStart,
    "Week End": week.weekEnd,
    "Weekly High / Trigger": week.trigger,
    "Ignorable Low": week.lower || "",
    "Ignorable High": week.upper || "",
    "In Ignorable Range": week.inIgnorableRange ? "Yes" : "No",
    "GTT Placed": week.gttPlaced ? "Yes" : "No",
    "GTT Status": week.status,
    "Shares Filled": week.filled ? week.shares : "",
    "Position Shares": week.positionShares,
    "Total Investment": week.positionInvestment,
    "Average Price": week.averagePrice || "",
    "Target 6 Percent": week.target || "",
    "Gross Profit": week.grossProfit || "",
    "Net Profit": week.netProfit || "",
    "Realized Profit Running Total": week.realizedProfit,
    "Note": week.exitNote || week.note
  }));
}

function symbolKey(symbol) {
  return String(symbol || "").trim().toUpperCase();
}

// Runs the Darvas/GTT cycle math for a single stock's weeks, in order.
// Position size, average price, target, and realized profit only ever
// accumulate within this one symbol's rows.
function evaluateSymbolWeeks(weeks, chargesPerExit) {
  let totalShares = 0;
  let totalInvestment = 0;
  let lastExecutedTrigger = null;
  let realizedProfit = 0;
  let cycle = 1;

  return weeks.map((rawWeek) => {
    const week = normalizeWeek(rawWeek);
    const trigger = getWeeklyHigh(week);
    const lower = lastExecutedTrigger ? lastExecutedTrigger * (1 - IGNORE_PCT) : null;
    const upper = lastExecutedTrigger ? lastExecutedTrigger * (1 + IGNORE_PCT) : null;
    const inIgnorableRange = Boolean(lastExecutedTrigger && trigger >= lower && trigger <= upper);
    const gttPlaced = !inIgnorableRange;
    const shares = toNumber(week.shares);
    const nextWeekOpen = toNumber(week.nextWeekOpen);
    const nextWeekHigh = toNumber(week.nextWeekHigh);

    let status = "No GTT";
    let note = "Weekly high is inside the 2.75% band.";
    let filled = false;
    let gapIssue = false;

    if (gttPlaced) {
      if (week.execution === "filled") {
        filled = true;
        status = "Filled";
        note = `Bought ${shares} @ ${compactMoney(trigger)}.`;
      } else if (week.execution === "gap") {
        gapIssue = true;
        status = "Gap up";
        note = `Due to gap up opening unable to buy @ ${compactMoney(trigger)}.`;
      } else if (week.execution === "unable") {
        gapIssue = true;
        status = "Unable";
        note = "Unable to trigger.";
      } else if (nextWeekOpen > trigger && nextWeekHigh >= nextWeekOpen) {
        gapIssue = true;
        status = "Gap risk";
        note = `Next week opened above trigger; review fill @ ${compactMoney(trigger)}.`;
      } else {
        status = "Placed";
        note = "GTT is waiting for next week's touch.";
      }
    }

    if (filled) {
      totalShares += shares;
      totalInvestment += trigger * shares;
      lastExecutedTrigger = trigger;
    }

    const averagePrice = totalShares ? totalInvestment / totalShares : null;
    const target = averagePrice ? averagePrice * (1 + TARGET_PCT) : null;
    let exitNote = "";
    let grossProfit = 0;
    let netProfit = 0;

    if (week.exitHit && totalShares && target) {
      grossProfit = (target - averagePrice) * totalShares;
      netProfit = Math.max(0, Math.floor(grossProfit - chargesPerExit));
      realizedProfit += netProfit;
      exitNote = `6 Percent Profit Book @${compactMoney(target)} Total Profit ${compactMoney(grossProfit)} after Brokerage and DP charges near ${compactMoney(netProfit)}`;
      totalShares = 0;
      totalInvestment = 0;
      lastExecutedTrigger = null;
      cycle += 1;
    }

    return {
      ...week,
      cycle,
      trigger,
      lower,
      upper,
      inIgnorableRange,
      gttPlaced,
      status,
      note,
      filled,
      gapIssue,
      positionShares: totalShares,
      positionInvestment: totalInvestment,
      averagePrice: totalShares ? totalInvestment / totalShares : null,
      target: totalShares ? (totalInvestment / totalShares) * (1 + TARGET_PCT) : null,
      exitNote,
      grossProfit,
      netProfit,
      realizedProfit
    };
  });
}

// Groups weeks by symbol, evaluates each stock's cycle independently, then
// returns the flattened result back in the original array order. This is
// what makes multi-stock tracking correct: stock B's first row never
// inherits stock A's open position, average price, or running profit.
function evaluateWeeks(weeks, chargesPerExit) {
  const groups = new Map();
  weeks.forEach((week, index) => {
    const key = symbolKey(week.symbol);
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push({ week, index });
  });

  const evaluatedByIndex = new Array(weeks.length);
  groups.forEach((entries) => {
    const symbolWeeks = entries.map((entry) => entry.week);
    const evaluatedSymbolWeeks = evaluateSymbolWeeks(symbolWeeks, chargesPerExit);
    entries.forEach((entry, position) => {
      evaluatedByIndex[entry.index] = evaluatedSymbolWeeks[position];
    });
  });

  return evaluatedByIndex;
}

function Tracker({ user, onSignOut }) {
  const [weeks, setWeeks] = useState([]);
  const [chargesPerExit, setChargesPerExit] = useState(65);
  const [loading, setLoading] = useState(true);
  const [loadError, setLoadError] = useState("");
  const [saving, setSaving] = useState(false);
  const [toast, setToast] = useState(null);
  const [selectedSymbol, setSelectedSymbol] = useState(ALL_STOCKS);
  const toastTimer = useRef(null);

  // Load this user's weeks + settings from Supabase once on mount / on
  // user change. Falls back to the bundled sample data only if the user
  // has never saved anything yet (so brand-new accounts aren't blank).
  useEffect(() => {
    let cancelled = false;
    async function loadData() {
      setLoading(true);
      setLoadError("");
      const [weeksResult, settingsResult] = await Promise.all([
        supabase
          .from("tracker_weeks")
          .select("*")
          .eq("user_id", user.id)
          .order("symbol", { ascending: true })
          .order("week_start", { ascending: true }),
        supabase.from("tracker_settings").select("*").eq("user_id", user.id).maybeSingle()
      ]);
      if (cancelled) return;

      if (weeksResult.error) {
        setLoadError(weeksResult.error.message);
        setLoading(false);
        return;
      }

      if (weeksResult.data && weeksResult.data.length > 0) {
        setWeeks(weeksResult.data.map(dbRowToWeek));
      } else {
        // First time this account has been used: seed with sample data
        // locally. It won't be saved until the user hits Save.
        setWeeks(seedWeeks);
      }

      if (settingsResult.data) {
        setChargesPerExit(Number(settingsResult.data.charges_per_exit));
      }

      setLoading(false);
    }
    loadData();
    return () => {
      cancelled = true;
    };
  }, [user.id]);

  // evaluateWeeks groups internally by symbol, so each stock's position,
  // average price, target, and realized profit are computed independently
  // even though everything still lives in one flat `weeks` array.
  const evaluated = useMemo(
    () => evaluateWeeks(weeks, toNumber(chargesPerExit)),
    [weeks, chargesPerExit]
  );

  const symbols = useMemo(() => {
    const seen = new Map();
    weeks.forEach((week) => {
      const key = symbolKey(week.symbol);
      if (key && !seen.has(key)) seen.set(key, week.symbol.trim());
    });
    return [...seen.values()].sort((a, b) => a.localeCompare(b));
  }, [weeks]);

  // Reset back to "All stocks" if the currently selected symbol no longer
  // exists (e.g. its last row was deleted or renamed).
  useEffect(() => {
    if (selectedSymbol !== ALL_STOCKS && !symbols.includes(selectedSymbol)) {
      setSelectedSymbol(ALL_STOCKS);
    }
  }, [symbols, selectedSymbol]);

  const visibleEvaluated = useMemo(() => {
    if (selectedSymbol === ALL_STOCKS) return evaluated;
    return evaluated.filter((week) => symbolKey(week.symbol) === symbolKey(selectedSymbol));
  }, [evaluated, selectedSymbol]);

  // Per-stock "last row" lookup, used to combine metrics correctly for the
  // All Stocks view (summing each stock's own running totals rather than
  // summing every row, which would double count).
  const lastRowPerSymbol = useMemo(() => {
    const map = new Map();
    evaluated.forEach((week) => {
      map.set(symbolKey(week.symbol), week);
    });
    return map;
  }, [evaluated]);

  const dashboard = useMemo(() => {
    if (selectedSymbol === ALL_STOCKS) {
      const lastRows = [...lastRowPerSymbol.values()];
      const realizedProfit = lastRows.reduce((sum, week) => sum + (week.realizedProfit || 0), 0);
      const activeRows = lastRows.filter((week) => week.positionShares > 0);
      const activeShares = activeRows.reduce((sum, week) => sum + week.positionShares, 0);
      const activeInvestment = activeRows.reduce((sum, week) => sum + week.positionInvestment, 0);
      const averagePrice = activeShares ? activeInvestment / activeShares : null;
      return {
        realizedProfit,
        positionShares: activeShares,
        averagePrice,
        // A single blended target across multiple open stocks isn't
        // meaningful, so show how many stocks currently have an open
        // position instead.
        target: null,
        openPositions: activeRows.length,
        placedCount: evaluated.filter((week) => week.gttPlaced).length,
        filledCount: evaluated.filter((week) => week.filled).length
      };
    }
    const last = lastRowPerSymbol.get(symbolKey(selectedSymbol));
    return {
      realizedProfit: last?.realizedProfit ?? 0,
      positionShares: last?.positionShares ?? 0,
      averagePrice: last?.averagePrice ?? null,
      target: last?.target ?? null,
      openPositions: last?.positionShares > 0 ? 1 : 0,
      placedCount: visibleEvaluated.filter((week) => week.gttPlaced).length,
      filledCount: visibleEvaluated.filter((week) => week.filled).length
    };
  }, [selectedSymbol, lastRowPerSymbol, evaluated, visibleEvaluated]);

  function updateWeek(id, field, value) {
    setWeeks((current) =>
      current.map((week) => (week.id === id ? { ...week, [field]: value } : week))
    );
  }

  function showToast(message, type = "success") {
    if (toastTimer.current) clearTimeout(toastTimer.current);
    setToast({ message, type });
    toastTimer.current = setTimeout(() => setToast(null), 3200);
  }

  function addWeek() {
    const stockWeeks = selectedSymbol === ALL_STOCKS
      ? weeks
      : weeks.filter((week) => symbolKey(week.symbol) === symbolKey(selectedSymbol));
    const previous = stockWeeks.at(-1);
    const symbol = selectedSymbol === ALL_STOCKS ? previous?.symbol || "" : selectedSymbol;
    setWeeks((current) => [
      ...current,
      {
        id: crypto.randomUUID(),
        symbol,
        weekStart: "",
        weekEnd: "",
        monHigh: "",
        tueHigh: "",
        wedHigh: "",
        thuHigh: "",
        friHigh: "",
        shares: "",
        execution: "pending",
        executionWeek: "",
        nextWeekOpen: "",
        nextWeekHigh: "",
        exitHit: false
      }
    ]);
    showToast("New weekend row added.");
  }

  async function saveData() {
    setSaving(true);
    try {
      const dbRows = weeks.map((week) => weekToDbRow(week, user.id));

      // Upsert everything currently in state.
      if (dbRows.length > 0) {
        const { error: upsertError } = await supabase.from("tracker_weeks").upsert(dbRows);
        if (upsertError) throw upsertError;
      }

      // Remove any rows that exist in the DB for this user but are no
      // longer present locally (covers deleted rows).
      const currentIds = weeks.map((week) => week.id);
      const { data: existingRows, error: fetchError } = await supabase
        .from("tracker_weeks")
        .select("id")
        .eq("user_id", user.id);
      if (fetchError) throw fetchError;

      const idsToDelete = (existingRows || [])
        .map((row) => row.id)
        .filter((id) => !currentIds.includes(id));
      if (idsToDelete.length > 0) {
        const { error: deleteError } = await supabase
          .from("tracker_weeks")
          .delete()
          .in("id", idsToDelete);
        if (deleteError) throw deleteError;
      }

      const { error: settingsError } = await supabase
        .from("tracker_settings")
        .upsert({ user_id: user.id, charges_per_exit: toNumber(chargesPerExit) });
      if (settingsError) throw settingsError;

      showToast("Tracker saved to your account.");
    } catch (error) {
      showToast(`Save failed: ${error.message}`, "error");
    } finally {
      setSaving(false);
    }
  }

  async function resetData() {
    setWeeks(seedWeeks);
    setChargesPerExit(65);
    showToast("Sample tracker data restored locally. Click Save to overwrite your account's saved data.", "info");
  }

  function exportData() {
    try {
      const workbook = XLSX.utils.book_new();
      const inputSheet = XLSX.utils.json_to_sheet(weeksToInputRows(weeks), {
        header: excelInputHeaders
      });
      const computedSheet = XLSX.utils.json_to_sheet(evaluatedToRows(evaluated));
      const settingsSheet = XLSX.utils.json_to_sheet([
        { Setting: "Brokerage + DP estimate per completed exit", Value: chargesPerExit },
        { Setting: "Ignorable range percent", Value: IGNORE_PCT },
        { Setting: "Target percent", Value: TARGET_PCT }
      ]);

      inputSheet["!cols"] = excelInputHeaders.map(() => ({ wch: 16 }));
      computedSheet["!cols"] = Array.from({ length: 18 }, () => ({ wch: 18 }));
      settingsSheet["!cols"] = [{ wch: 42 }, { wch: 14 }];

      XLSX.utils.book_append_sheet(workbook, inputSheet, "Weekly Data");
      XLSX.utils.book_append_sheet(workbook, computedSheet, "Computed Tracker");
      XLSX.utils.book_append_sheet(workbook, settingsSheet, "Settings");
      XLSX.writeFile(workbook, "gtt-darvas-tracker.xlsx");
      showToast("Excel workbook exported.");
    } catch (error) {
      showToast("Export failed. Please try again.", "error");
    }
  }

  function importData(event) {
    const file = event.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const workbook = XLSX.read(reader.result, { type: "array", cellDates: true });
        const sheetName = workbook.SheetNames.includes("Weekly Data")
          ? "Weekly Data"
          : workbook.SheetNames[0];
        const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
          defval: "",
          raw: false
        });
        const fallbackSymbol =
          selectedSymbol !== ALL_STOCKS ? selectedSymbol : weeks.find((week) => week.symbol)?.symbol || "";
        const importedWeeks = rowsLookLikeNseDaily(rows)
          ? nseDailyRowsToWeeks(rows, fallbackSymbol)
          : workbookRowsToWeeks(rows);

        if (workbook.Sheets.Settings) {
          const settings = XLSX.utils.sheet_to_json(workbook.Sheets.Settings, { defval: "" });
          const charges = settings.find((row) =>
            String(row.Setting || "").toLowerCase().includes("brokerage")
          )?.Value;
          if (charges !== undefined && charges !== "") setChargesPerExit(Number(charges));
        }

        if (importedWeeks.length) {
          setWeeks((current) => {
            const existingKeys = new Set(current.map(weekKey));
            const newWeeks = importedWeeks.filter((week) => !existingKeys.has(weekKey(week)));
            if (!newWeeks.length) {
              showToast(`No new weekly rows found in ${file.name}.`, "info");
              return current;
            }
            showToast(`Imported and appended ${newWeeks.length} weekly row${newWeeks.length === 1 ? "" : "s"}.`);
            return sortWeeks([...current, ...newWeeks]);
          });
        } else {
          showToast(`No usable weekly or NSE daily rows found in ${file.name}.`, "error");
        }
      } catch (error) {
        showToast("Import failed. Please check the Excel format.", "error");
      }
    };
    reader.onerror = () => showToast("Import failed. Could not read the file.", "error");
    reader.readAsArrayBuffer(file);
    event.target.value = "";
  }

  if (loading) {
    return (
      <main>
        <p className="loadingState">Loading your tracker...</p>
      </main>
    );
  }

  if (loadError) {
    return (
      <main>
        <p className="loadingState error">Could not load your data: {loadError}</p>
      </main>
    );
  }

  return (
    <main>
      {toast && (
        <div className={`toast ${toast.type}`} role="status" aria-live="polite">
          {toast.type === "error" ? <X size={18} /> : <Check size={18} />}
          <span>{toast.message}</span>
        </div>
      )}
      <header className="topbar">
        <div>
          <p className="eyebrow">Weekly Darvas Box GTT &middot; {user.email}</p>
          <h1>GTT Order Tracker</h1>
        </div>
        <div className="actions">
          <button type="button" onClick={saveData} disabled={saving} title="Save to your account">
            <Save size={18} /> {saving ? "Saving..." : "Save"}
          </button>
          <button type="button" onClick={exportData} title="Export JSON backup">
            <Download size={18} /> Export Excel
          </button>
          <label className="fileButton" title="Import Excel workbook">
            <Upload size={18} /> Import Excel
            <input type="file" accept=".xlsx,.xls,.csv,.ods" onChange={importData} />
          </label>
          <button type="button" className="ghost" onClick={resetData} title="Reset sample data">
            <RotateCcw size={18} />
          </button>
          <button type="button" className="ghost" onClick={onSignOut} title="Sign out">
            <LogOut size={18} />
          </button>
        </div>
      </header>

      <nav className="stockTabs" aria-label="Select stock">
        <button
          type="button"
          className={selectedSymbol === ALL_STOCKS ? "active" : ""}
          onClick={() => setSelectedSymbol(ALL_STOCKS)}
        >
          All stocks
        </button>
        {symbols.map((symbol) => (
          <button
            key={symbol}
            type="button"
            className={symbolKey(selectedSymbol) === symbolKey(symbol) ? "active" : ""}
            onClick={() => setSelectedSymbol(symbol)}
          >
            {symbol}
          </button>
        ))}
      </nav>

      <section className="summary">
        <Metric
          icon={<BadgeIndianRupee />}
          label="Realized profit"
          value={money(dashboard.realizedProfit)}
        />
        {selectedSymbol === ALL_STOCKS ? (
          <Metric icon={<Activity />} label="Open positions" value={dashboard.openPositions} />
        ) : (
          <Metric icon={<Activity />} label="Active shares" value={dashboard.positionShares} />
        )}
        <Metric
          icon={<LineChart />}
          label={selectedSymbol === ALL_STOCKS ? "Blended average" : "Average / target"}
          value={
            selectedSymbol === ALL_STOCKS
              ? money(dashboard.averagePrice)
              : `${money(dashboard.averagePrice)} / ${money(dashboard.target)}`
          }
        />
        <Metric icon={<Check />} label="Placed / filled" value={`${dashboard.placedCount} / ${dashboard.filledCount}`} />
      </section>

      <section className="controls">
        <label>
          Brokerage + DP estimate per completed exit
          <input
            type="number"
            value={chargesPerExit}
            onChange={(event) => setChargesPerExit(event.target.value)}
          />
        </label>
        <button type="button" onClick={addWeek}>
          <Plus size={18} />
          {selectedSymbol === ALL_STOCKS ? "Add weekend data" : `Add ${selectedSymbol} week`}
        </button>
      </section>
      <section className="tableWrap" aria-label="Weekly GTT tracker">
        <table>
          <thead>
            <tr>
              <th>Week</th>
              <th>Daily highs</th>
              <th>Trigger</th>
              <th>Ignorable band</th>
              <th>GTT result</th>
              <th>Position</th>
              <th>Exit</th>
              <th></th>
            </tr>
          </thead>
          <tbody>
            {visibleEvaluated.map((week) => (
              <tr key={week.id}>
                <td className="weekCell">
                  <input
                    value={week.symbol}
                    onChange={(event) => updateWeek(week.id, "symbol", event.target.value)}
                    placeholder="Symbol"
                  />
                  <div className="dateGrid">
                    <input
                      type="date"
                      value={week.weekStart}
                      onChange={(event) => updateWeek(week.id, "weekStart", event.target.value)}
                    />
                    <input
                      type="date"
                      value={week.weekEnd}
                      onChange={(event) => updateWeek(week.id, "weekEnd", event.target.value)}
                    />
                  </div>
                </td>
                <td>
                  <div className="highGrid">
                    {["monHigh", "tueHigh", "wedHigh", "thuHigh", "friHigh"].map((field, index) => (
                      <label key={field}>
                        {["M", "T", "W", "T", "F"][index]}
                        <input
                          type="number"
                          value={week[field]}
                          onChange={(event) => updateWeek(week.id, field, event.target.value)}
                        />
                      </label>
                    ))}
                  </div>
                </td>
                <td>
                  <strong>{money(week.trigger)}</strong>
                  <span>Next Week Buying Trigger Price</span>
                </td>
                <td>
                  {week.lower ? (
                    <>
                      <strong>{money(week.lower)} - {money(week.upper)}</strong>
                      <span className={week.inIgnorableRange ? "dangerText" : "goodText"}>
                        {week.inIgnorableRange ? "In ignorable range" : "Fresh breakout level"}
                      </span>
                    </>
                  ) : (
                    <span>No previous executed GTT</span>
                  )}
                </td>
                <td>
                  <div className={`pill ${week.status.toLowerCase().replaceAll(" ", "")}`}>
                    {week.gttPlaced ? <Check size={14} /> : <X size={14} />}
                    {week.status}
                  </div>
                  <div className="executionGrid">
                    <select
                      value={week.execution}
                      onChange={(event) => updateWeek(week.id, "execution", event.target.value)}
                    >
                      <option value="pending">Placed / pending</option>
                      <option value="filled">Filled @ trigger</option>
                      <option value="gap">Gap up unable</option>
                      <option value="unable">Unable to trigger</option>
                    </select>
                    <input
                      type="number"
                      value={week.shares}
                      onChange={(event) => updateWeek(week.id, "shares", event.target.value)}
                      placeholder="Shares"
                    />
                  </div>
                  <small>{week.note}</small>
                </td>
                <td>
                  <strong>{week.positionShares} shares</strong>
                  <span>Invested {money(week.positionInvestment)}</span>
                  <span>Avg {money(week.averagePrice)}</span>
                  <span>Target {money(week.target)}</span>
                </td>
                <td>
                  <label className="checkLine">
                    <input
                      type="checkbox"
                      checked={Boolean(week.exitHit)}
                      onChange={(event) => updateWeek(week.id, "exitHit", event.target.checked)}
                    />
                    Target hit
                  </label>
                  <small>{week.exitNote || "Position carries until target trades."}</small>
                </td>
                <td>
                  <button
                    type="button"
                    className="iconOnly danger"
                    title="Delete row"
                    onClick={() => {
                      setWeeks((current) => current.filter((item) => item.id !== week.id));
                      showToast("Weekly row deleted.", "info");
                    }}
                  >
                    <Trash2 size={18} />
                  </button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </section>
    </main>
  );
}

function Metric({ icon, label, value }) {
  return (
    <article className="metric">
      <div className="metricIcon">{React.cloneElement(icon, { size: 20 })}</div>
      <span>{label}</span>
      <strong>{value}</strong>
    </article>
  );
}

function LoginScreen() {
  const [mode, setMode] = useState("sign_in");
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [submitting, setSubmitting] = useState(false);
  const [message, setMessage] = useState(null);

  async function handleSubmit(event) {
    event.preventDefault();
    setSubmitting(true);
    setMessage(null);

    if (mode === "sign_in") {
      const { error } = await supabase.auth.signInWithPassword({ email, password });
      if (error) setMessage({ type: "error", text: error.message });
    } else {
      const { error } = await supabase.auth.signUp({ email, password });
      if (error) {
        setMessage({ type: "error", text: error.message });
      } else {
        setMessage({
          type: "success",
          text: "Account created. Check your email to confirm, then sign in."
        });
      }
    }
    setSubmitting(false);
  }

  return (
    <main className="authShell">
      <form className="authCard" onSubmit={handleSubmit}>
        <p className="eyebrow">Weekly Darvas Box GTT</p>
        <h1>{mode === "sign_in" ? "Sign in" : "Create account"}</h1>
        <label>
          Email
          <input
            type="email"
            required
            autoComplete="email"
            value={email}
            onChange={(event) => setEmail(event.target.value)}
          />
        </label>
        <label>
          Password
          <input
            type="password"
            required
            minLength={6}
            autoComplete={mode === "sign_in" ? "current-password" : "new-password"}
            value={password}
            onChange={(event) => setPassword(event.target.value)}
          />
        </label>
        {message && (
          <p className={message.type === "error" ? "dangerText" : "goodText"}>{message.text}</p>
        )}
        <button type="submit" disabled={submitting}>
          {submitting ? "Please wait..." : mode === "sign_in" ? "Sign in" : "Sign up"}
        </button>
        <button
          type="button"
          className="ghost linkButton"
          onClick={() => {
            setMode(mode === "sign_in" ? "sign_up" : "sign_in");
            setMessage(null);
          }}
        >
          {mode === "sign_in" ? "Need an account? Sign up" : "Already have an account? Sign in"}
        </button>
      </form>
    </main>
  );
}

function App() {
  const [session, setSession] = useState(null);
  const [checkingSession, setCheckingSession] = useState(true);

  useEffect(() => {
    supabase.auth.getSession().then(({ data }) => {
      setSession(data.session);
      setCheckingSession(false);
    });
    const { data: listener } = supabase.auth.onAuthStateChange((_event, nextSession) => {
      setSession(nextSession);
    });
    return () => listener.subscription.unsubscribe();
  }, []);

  if (checkingSession) {
    return (
      <main>
        <p className="loadingState">Loading...</p>
      </main>
    );
  }

  if (!session) {
    return <LoginScreen />;
  }

  return <Tracker user={session.user} onSignOut={() => supabase.auth.signOut()} />;
}

createRoot(document.getElementById("root")).render(<App />);
