import React, { useState, useMemo, useCallback, useRef } from "react";
import Papa from "papaparse";
import * as XLSX from "xlsx";
import {
  ComposedChart, Bar, Line, PieChart, Pie, Cell,
  XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer
} from "recharts";
import {
  Upload, RefreshCw, Download, AlertCircle, TrendingUp, TrendingDown,
  Scale, FileSpreadsheet, Search, X, Trash2
} from "lucide-react";

// ---------------------------------------------------------------------------
// Tokens
// ---------------------------------------------------------------------------
const INK = "#242A21";
const MUTED = "#7A7D70";
const PAPER = "#F7F5EE";
const CARD = "#FFFFFF";
const LINE = "#E4E1D3";
const INCOME = "#3F6B52";
const INCOME_SOFT = "#E7EFE7";
const EXPENSE = "#8A3B4E";
const EXPENSE_SOFT = "#F5E7EA";
const GOLD = "#AD8A2C";

// Decorative page background — an original bioluminescent night-jungle motif,
// evoking floating glowing spores and forest silhouettes. No copyrighted
// artwork, characters, or stills are used; this is an abstract atmosphere only.
const BG_DEEP = "#03141a";
const BG_MID = "#0a2b30";
const BG_EDGE = "#062024";
const GLOW_TEAL = "#5EEAD4";
const GLOW_CYAN = "#22D3EE";
const FOREST_LINE = "#0f4a42";
const ON_DARK_TEXT = "#EAF7F4";
const ON_DARK_MUTED = "#9FD9CE";

const SPORES = Array.from({ length: 22 }).map((_, i) => ({
  left: Math.round(Math.random() * 100),
  top: Math.round(Math.random() * 100),
  size: 2 + Math.random() * 4,
  duration: 14 + Math.random() * 12,
  delay: Math.random() * 10,
  glow: i % 3 === 0 ? GLOW_CYAN : GLOW_TEAL,
}));

const CAT_PALETTE = [
  "#3F6B52", "#8A3B4E", "#AD8A2C", "#3F6B8A", "#8A5A3F",
  "#5A6B3F", "#6B3F8A", "#3F8A7A", "#8A5F3F", "#3F5A8A",
];

const EXPENSE_CATEGORIES = [
  "Groceries", "Rent/Housing", "Utilities", "Transport", "Dining",
  "Entertainment", "Shopping", "Healthcare", "Insurance",
  "Subscriptions", "Education", "Other Expense",
];
const INCOME_CATEGORIES = [
  "Siva_Salary", "Thivyaa_Salary", "Freelance", "Investment", "Gift", "Refund", "Business", "Other Income",
];

const EXPENSE_KEYWORDS = {
  Groceries: ["grocery", "groceries", "supermarket", "walmart", "tesco", "carrefour", "lulu", "spinneys", "whole foods", "aldi", "kroger"],
  "Rent/Housing": ["rent", "mortgage", "landlord", "housing society"],
  Utilities: ["electric", "water bill", "utility", "utilities", "dewa", "internet bill", "broadband", "etisalat", "du telecom", "gas bill"],
  Transport: ["uber", "taxi", "careem", "fuel", "petrol", "gas station", "metro", "parking", "lyft", "toll", "rta"],
  Dining: ["restaurant", "cafe", "coffee", "starbucks", "mcdonald", "kfc", "deliveroo", "talabat", "zomato", "swiggy", "diner", "bakery"],
  Entertainment: ["netflix", "spotify", "cinema", "movie", "disney", "prime video", "hulu", "game", "concert", "ticket"],
  Shopping: ["amazon", "noon", "shopping", "mall", "ikea", "zara", "h&m", "shein"],
  Healthcare: ["pharmacy", "hospital", "clinic", "doctor", "medical", "dental", "optic"],
  Insurance: ["insurance", "premium"],
  Subscriptions: ["subscription", "membership", "icloud", "adobe", "microsoft 365"],
  Education: ["tuition", "school", "course", "udemy", "coursera", "university"],
};

const INCOME_KEYWORDS = {
  Siva_Salary: ["siva"],
  Thivyaa_Salary: ["thivyaa", "thivya"],
  Freelance: ["freelance", "contract payment", "invoice paid"],
  Investment: ["dividend", "interest", "capital gain", "stock sale"],
  Gift: ["gift"],
  Refund: ["refund", "reimbursement", "cashback"],
  Business: ["business income", "sales revenue", "client payment"],
};

const DATE_KEYS = ["date", "transaction date", "posting date", "value date", "txn date"];
const DESC_KEYS = ["description", "narration", "particulars", "details", "memo", "payee", "remarks"];
const AMOUNT_KEYS = ["amount", "value", "transaction amount", "amt"];
const DEBIT_KEYS = ["debit", "withdrawal", "debit amount", "dr"];
const CREDIT_KEYS = ["credit", "deposit", "credit amount", "cr"];
const TYPE_KEYS = ["type", "transaction type", "dr/cr", "cr/dr"];
const CATEGORY_KEYS = ["category", "tag", "label"];

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------
function normKey(k) {
  return String(k || "").trim().toLowerCase();
}

function findKey(headers, candidates) {
  const norm = headers.map((h) => ({ raw: h, n: normKey(h) }));
  for (const c of candidates) {
    const hit = norm.find((h) => h.n === c);
    if (hit) return hit.raw;
  }
  for (const c of candidates) {
    const hit = norm.find((h) => h.n.includes(c));
    if (hit) return hit.raw;
  }
  return null;
}

function parseAmount(raw) {
  if (typeof raw === "number") return raw;
  if (!raw) return 0;
  let s = String(raw).trim();
  let neg = false;
  if (/^\(.*\)$/.test(s)) {
    neg = true;
    s = s.slice(1, -1);
  }
  if (/^-/.test(s)) neg = true;
  s = s.replace(/[^0-9.]/g, "");
  let val = parseFloat(s) || 0;
  return neg ? -Math.abs(val) : val;
}

function parseDate(raw) {
  if (!raw) return null;
  if (raw instanceof Date && !isNaN(raw)) return raw;
  if (typeof raw === "number") {
    // Excel serial date
    const d = XLSX.SSF.parse_date_code(raw);
    if (d) return new Date(d.y, d.m - 1, d.d);
  }
  const s = String(raw).trim();
  let d = new Date(s);
  if (!isNaN(d.getTime())) return d;
  const parts = s.split(/[\/\-.]/);
  if (parts.length === 3) {
    let [a, b, c] = parts.map((p) => parseInt(p, 10));
    if (c < 100) c += 2000;
    d = new Date(c, b - 1, a);
    if (!isNaN(d.getTime())) return d;
    d = new Date(c, a - 1, b);
    if (!isNaN(d.getTime())) return d;
  }
  return null;
}

function autoCategorize(desc, type) {
  const d = String(desc || "").toLowerCase();
  const dict = type === "income" ? INCOME_KEYWORDS : EXPENSE_KEYWORDS;
  for (const [cat, keywords] of Object.entries(dict)) {
    if (keywords.some((k) => d.includes(k))) return cat;
  }
  return type === "income" ? "Other Income" : "Other Expense";
}

function dayKey(date) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, "0");
  const d = String(date.getDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
}
function dayLabel(key) {
  const [y, m, d] = key.split("-").map(Number);
  return new Date(y, m - 1, d).toLocaleDateString(undefined, { month: "short", day: "numeric" });
}
function monthKey(date) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
}
function monthLabel(key) {
  const [y, m] = key.split("-");
  return new Date(Number(y), Number(m) - 1, 1).toLocaleDateString(undefined, { month: "short", year: "numeric" });
}
function yearKey(date) {
  return String(date.getFullYear());
}
function yearLabel(key) {
  return key;
}

const PERIODS = {
  daily: { key: dayKey, label: dayLabel, take: 14, title: "Daily" },
  monthly: { key: monthKey, label: monthLabel, take: 12, title: "Monthly" },
  yearly: { key: yearKey, label: yearLabel, take: 20, title: "Yearly" },
};

const CURRENCIES = {
  USD: "$", EUR: "€", GBP: "£", INR: "₹", AED: "AED ", JPY: "¥",
};

function fmtMoney(n, symbol) {
  const sign = n < 0 ? "-" : "";
  const abs = Math.abs(n);
  return `${sign}${symbol}${abs.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
}

let idCounter = 1;
function nextId() { return idCounter++; }

// Sample data generator for "try it out"
function buildSampleData() {
  const today = new Date();
  const rows = [];
  const sampleExpenses = [
    ["Whole Foods Market", 86.4], ["Uber trip", 18.5], ["Starbucks Coffee", 6.75],
    ["Netflix Subscription", 15.99], ["DEWA Electricity Bill", 120.0], ["Rent Payment", 1450.0],
    ["Amazon Purchase", 64.2], ["Talabat Delivery", 32.1], ["Pharmacy - Medicines", 22.0],
    ["Gym Membership", 45.0], ["Spotify Premium", 10.99], ["Fuel Station", 55.0],
    ["Zara Shopping", 98.5], ["Health Insurance Premium", 210.0], ["Restaurant Dinner", 74.3],
  ];
  const sampleIncome = [
    ["Siva Monthly Salary", 4200], ["Thivyaa Monthly Salary", 3100], ["Freelance Project Payment", 650],
    ["Dividend Payout", 120], ["Cashback Refund", 25], ["Bonus Payment", 500],
  ];
  for (let m = 5; m >= 0; m--) {
    const base = new Date(today.getFullYear(), today.getMonth() - m, 1);
    sampleIncome.forEach(([desc, amt], i) => {
      const d = new Date(base.getFullYear(), base.getMonth(), 1 + i * 3);
      if (d <= today) rows.push({ date: d, description: desc, amount: amt, type: "income" });
    });
    sampleExpenses.forEach(([desc, amt], i) => {
      const d = new Date(base.getFullYear(), base.getMonth(), 2 + ((i * 2) % 27));
      if (d <= today) {
        const jitter = amt * (0.85 + Math.random() * 0.3);
        rows.push({ date: d, description: desc, amount: Number(jitter.toFixed(2)), type: "expense" });
      }
    });
  }
  return rows.map((r) => ({
    id: nextId(),
    date: r.date,
    description: r.description,
    amount: Math.abs(r.amount),
    type: r.type,
    category: autoCategorize(r.description, r.type),
  }));
}

// ---------------------------------------------------------------------------
// Component
// ---------------------------------------------------------------------------
export default function BudgetPlanner() {
  const [transactions, setTransactions] = useState([]);
  const [currency, setCurrency] = useState("AED");
  const [period, setPeriod] = useState("monthly");
  const [dailyFrom, setDailyFrom] = useState("");
  const [dailyTo, setDailyTo] = useState("");
  const [search, setSearch] = useState("");
  const [typeFilter, setTypeFilter] = useState("all");
  const [error, setError] = useState("");
  const fileRef = useRef(null);
  const symbol = CURRENCIES[currency];

  const handleFile = useCallback((file) => {
    setError("");
    const ext = file.name.split(".").pop().toLowerCase();

    const processRows = (rows) => {
      if (!rows.length) {
        setError("That file doesn't seem to have any rows we could read.");
        return;
      }
      const headers = Object.keys(rows[0]);
      const dateH = findKey(headers, DATE_KEYS);
      const descH = findKey(headers, DESC_KEYS);
      const amountH = findKey(headers, AMOUNT_KEYS);
      const debitH = findKey(headers, DEBIT_KEYS);
      const creditH = findKey(headers, CREDIT_KEYS);
      const typeH = findKey(headers, TYPE_KEYS);
      const categoryH = findKey(headers, CATEGORY_KEYS);

      if (!dateH || (!amountH && !debitH && !creditH)) {
        setError("Couldn't find date and amount columns. Expected headers like Date, Description, Amount (or Debit/Credit).");
        return;
      }

      const parsed = rows.map((row) => {
        const date = parseDate(row[dateH]);
        const description = descH ? String(row[descH] || "").trim() : "Transaction";
        let amount = 0;
        let type = null;

        if (amountH) {
          amount = parseAmount(row[amountH]);
          type = amount >= 0 ? "income" : "expense";
        } else {
          const debit = debitH ? parseAmount(row[debitH]) : 0;
          const credit = creditH ? parseAmount(row[creditH]) : 0;
          if (Math.abs(debit) > 0) { amount = -Math.abs(debit); type = "expense"; }
          else { amount = Math.abs(credit); type = "income"; }
        }

        if (typeH) {
          const t = normKey(row[typeH]);
          if (["income", "credit", "cr", "deposit"].some((x) => t.includes(x))) type = "income";
          else if (["expense", "debit", "dr", "withdrawal"].some((x) => t.includes(x))) type = "expense";
        }

        let category = categoryH ? String(row[categoryH] || "").trim() : "";
        if (!category) category = autoCategorize(description, type);

        return {
          id: nextId(),
          date,
          description: description || "Transaction",
          amount: Math.abs(amount),
          type,
          category,
        };
      }).filter((r) => r.date);

      if (!parsed.length) {
        setError("We couldn't parse any valid dates from this file.");
        return;
      }
      setTransactions(parsed);
    };

    if (ext === "csv") {
      const reader = new FileReader();
      reader.onload = (e) => {
        const result = Papa.parse(e.target.result, { header: true, skipEmptyLines: true });
        processRows(result.data);
      };
      reader.readAsText(file);
    } else if (ext === "xlsx" || ext === "xls") {
      const reader = new FileReader();
      reader.onload = (e) => {
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, { type: "array" });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
        processRows(rows);
      };
      reader.readAsArrayBuffer(file);
    } else {
      setError("Please upload a .csv, .xlsx, or .xls file.");
    }
  }, []);

  const onDrop = (e) => {
    e.preventDefault();
    if (e.dataTransfer.files?.[0]) handleFile(e.dataTransfer.files[0]);
  };

  const updateTransaction = (id, patch) => {
    setTransactions((prev) => prev.map((t) => (t.id === id ? { ...t, ...patch } : t)));
  };

  const deleteTransaction = (id) => {
    setTransactions((prev) => prev.filter((t) => t.id !== id));
  };

  const reset = () => {
    setTransactions([]);
    setError("");
    setSearch("");
    setTypeFilter("all");
  };

  const loadSample = () => setTransactions(buildSampleData());

  // ---- Derived data ------------------------------------------------------
  const totals = useMemo(() => {
    let income = 0, expense = 0;
    transactions.forEach((t) => {
      if (t.type === "income") income += t.amount;
      else expense += t.amount;
    });
    return { income, expense, net: income - expense };
  }, [transactions]);

  const uncategorizedCount = useMemo(
    () => transactions.filter((t) => !t.category || t.category === "Other Expense" || t.category === "Other Income").length,
    [transactions]
  );

  const periodCfg = PERIODS[period];
  const useDailyRange = period === "daily" && dailyFrom && dailyTo;

  const dateExtent = useMemo(() => {
    if (!transactions.length) return null;
    let min = transactions[0].date, max = transactions[0].date;
    transactions.forEach((t) => {
      if (t.date < min) min = t.date;
      if (t.date > max) max = t.date;
    });
    return { min: dayKey(min), max: dayKey(max) };
  }, [transactions]);

  const trendData = useMemo(() => {
    const buckets = {};
    let scoped = transactions;
    if (useDailyRange) {
      const start = new Date(dailyFrom);
      start.setHours(0, 0, 0, 0);
      const end = new Date(dailyTo);
      end.setHours(23, 59, 59, 999);
      scoped = transactions.filter((t) => t.date >= start && t.date <= end);
    }
    scoped.forEach((t) => {
      const k = periodCfg.key(t.date);
      if (!buckets[k]) buckets[k] = { key: k, income: 0, expense: 0 };
      if (t.type === "income") buckets[k].income += t.amount;
      else buckets[k].expense += t.amount;
    });
    const arr = Object.values(buckets).sort((a, b) => (a.key > b.key ? 1 : -1));
    const sliced = useDailyRange ? arr : arr.slice(-periodCfg.take);
    return sliced.map((b) => ({
      ...b,
      label: periodCfg.label(b.key),
      net: b.income - b.expense,
    }));
  }, [transactions, periodCfg, useDailyRange, dailyFrom, dailyTo]);

  const visibleKeys = useMemo(() => new Set(trendData.map((d) => d.key)), [trendData]);

  const categoryBreakdown = useMemo(() => {
    const inScope = transactions.filter((t) => visibleKeys.has(periodCfg.key(t.date)));
    const expenseMap = {};
    const incomeMap = {};
    inScope.forEach((t) => {
      const map = t.type === "expense" ? expenseMap : incomeMap;
      map[t.category] = (map[t.category] || 0) + t.amount;
    });
    const toArr = (map) =>
      Object.entries(map)
        .map(([name, value]) => ({ name, value }))
        .sort((a, b) => b.value - a.value);
    return { expense: toArr(expenseMap), income: toArr(incomeMap) };
  }, [transactions, visibleKeys, periodCfg]);

  const filteredTransactions = useMemo(() => {
    return transactions
      .filter((t) => (typeFilter === "all" ? true : t.type === typeFilter))
      .filter((t) => t.description.toLowerCase().includes(search.toLowerCase()) || t.category.toLowerCase().includes(search.toLowerCase()))
      .sort((a, b) => b.date - a.date);
  }, [transactions, typeFilter, search]);

  const exportCsv = () => {
    const rows = transactions.map((t) => ({
      Date: t.date.toISOString().slice(0, 10),
      Description: t.description,
      Amount: t.amount,
      Type: t.type,
      Category: t.category,
    }));
    const csv = Papa.unparse(rows);
    const blob = new Blob([csv], { type: "text/csv" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "categorized-budget.csv";
    a.click();
    URL.revokeObjectURL(url);
  };

  // ---- Styles --------------------------------------------------------------
  const styles = {
    root: {
      fontFamily: "'Inter', -apple-system, sans-serif",
      background: `radial-gradient(ellipse 1100px 700px at 15% -10%, ${BG_MID} 0%, transparent 60%), radial-gradient(ellipse 900px 650px at 105% 5%, #0d3b3f 0%, transparent 55%), linear-gradient(180deg, ${BG_DEEP} 0%, ${BG_EDGE} 100%)`,
      color: INK,
      minHeight: "100vh",
      padding: "0",
      position: "relative",
      overflowX: "hidden",
    },
    inner: { maxWidth: 1120, margin: "0 auto", padding: "40px 28px 80px", position: "relative", zIndex: 1 },
    header: {
      display: "flex", justifyContent: "space-between", alignItems: "flex-end",
      flexWrap: "wrap", gap: 20, marginBottom: 36,
      borderBottom: `1px solid rgba(255,255,255,0.14)`, paddingBottom: 24,
    },
    eyebrow: {
      fontFamily: "'IBM Plex Mono', monospace", fontSize: 11, letterSpacing: "0.14em",
      textTransform: "uppercase", color: ON_DARK_MUTED, marginBottom: 6,
    },
    h1: {
      fontFamily: "'Fraunces', Georgia, serif", fontSize: 40, fontWeight: 600,
      margin: 0, letterSpacing: "-0.01em", color: ON_DARK_TEXT,
      textShadow: `0 0 22px ${GLOW_TEAL}40`,
    },
    controlsRow: { display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap" },
    select: {
      fontFamily: "'IBM Plex Mono', monospace", fontSize: 13, padding: "9px 12px",
      borderRadius: 6, border: `1px solid ${LINE}`, background: CARD, color: INK, cursor: "pointer",
    },
    btn: {
      fontFamily: "'Inter', sans-serif", fontSize: 13.5, fontWeight: 500, padding: "9px 16px",
      borderRadius: 6, border: `1px solid ${GLOW_TEAL}55`, background: INK, color: PAPER,
      cursor: "pointer", display: "inline-flex", alignItems: "center", gap: 7,
    },
    btnGhost: {
      fontFamily: "'Inter', sans-serif", fontSize: 13.5, fontWeight: 500, padding: "9px 16px",
      borderRadius: 6, border: `1px solid ${LINE}`, background: "transparent", color: INK,
      cursor: "pointer", display: "inline-flex", alignItems: "center", gap: 7,
    },
    btnGhostDark: {
      fontFamily: "'Inter', sans-serif", fontSize: 13.5, fontWeight: 500, padding: "9px 16px",
      borderRadius: 6, border: `1px solid rgba(255,255,255,0.28)`, background: "rgba(255,255,255,0.06)",
      color: ON_DARK_TEXT, cursor: "pointer", display: "inline-flex", alignItems: "center", gap: 7,
      backdropFilter: "blur(6px)",
    },
  };

  const isEmpty = transactions.length === 0;

  return (
    <div style={styles.root}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Fraunces:wght@400;500;600;700&family=Inter:wght@400;500;600&family=IBM+Plex+Mono:wght@400;500&display=swap');
        * { box-sizing: border-box; }
        .bp-tab { font-family:'Inter',sans-serif; font-size:13.5px; font-weight:500; padding:8px 18px; border-radius:6px; border:1px solid ${LINE}; background:${CARD}; color:${MUTED}; cursor:pointer; }
        .bp-tab.active { background:${INK}; color:${PAPER}; border-color:${INK}; }
        .bp-row:hover { background:#FAF9F4; }
        .bp-drop { transition: border-color .15s ease, background .15s ease; }
        .bp-drop:hover { border-color:${GOLD}; background:#FBF7EC; }
        select.bp-inline { font-family:'Inter',sans-serif; font-size:12.5px; padding:5px 8px; border-radius:5px; border:1px solid ${LINE}; background:#fff; color:${INK}; }
        input.bp-search { font-family:'Inter',sans-serif; font-size:13.5px; padding:9px 12px 9px 34px; border-radius:6px; border:1px solid ${LINE}; background:#fff; width:240px; }
        @keyframes bpFloat {
          0%, 100% { transform: translateY(0) translateX(0); opacity: 0.45; }
          50% { transform: translateY(-36px) translateX(8px); opacity: 1; }
        }
      `}</style>

      <div style={{ position: "fixed", inset: 0, overflow: "hidden", pointerEvents: "none", zIndex: 0 }} aria-hidden="true">
        {SPORES.map((s, i) => (
          <div
            key={i}
            style={{
              position: "absolute", left: `${s.left}%`, top: `${s.top}%`,
              width: s.size, height: s.size, borderRadius: "50%",
              background: s.glow, boxShadow: `0 0 ${s.size * 3}px ${s.size}px ${s.glow}55`,
              animation: `bpFloat ${s.duration}s ease-in-out ${s.delay}s infinite`,
            }}
          />
        ))}
        <svg style={{ position: "absolute", bottom: -30, left: -50, width: 320, opacity: 0.22 }} viewBox="0 0 220 320" fill="none">
          <path d="M110 320 C100 240 70 200 25 165" stroke={FOREST_LINE} strokeWidth="3" strokeLinecap="round" />
          <path d="M110 320 C105 220 125 165 170 100" stroke={FOREST_LINE} strokeWidth="3" strokeLinecap="round" />
          <path d="M110 320 C110 260 110 210 110 160" stroke={FOREST_LINE} strokeWidth="3" strokeLinecap="round" />
          <circle cx="25" cy="165" r="4" fill={GLOW_TEAL} opacity="0.7" />
          <circle cx="170" cy="100" r="4" fill={GLOW_CYAN} opacity="0.7" />
          <circle cx="110" cy="160" r="3" fill={GLOW_TEAL} opacity="0.6" />
        </svg>
        <svg style={{ position: "absolute", bottom: -30, right: -50, width: 320, opacity: 0.2 }} viewBox="0 0 220 320" fill="none">
          <path d="M110 320 C120 240 150 200 195 165" stroke={FOREST_LINE} strokeWidth="3" strokeLinecap="round" />
          <path d="M110 320 C115 220 95 165 50 100" stroke={FOREST_LINE} strokeWidth="3" strokeLinecap="round" />
          <path d="M110 320 C110 260 110 210 110 160" stroke={FOREST_LINE} strokeWidth="3" strokeLinecap="round" />
          <circle cx="195" cy="165" r="4" fill={GLOW_CYAN} opacity="0.7" />
          <circle cx="50" cy="100" r="4" fill={GLOW_TEAL} opacity="0.7" />
        </svg>
      </div>

      <div style={styles.inner}>
        {/* Header */}
        <div style={styles.header}>
          <div>
            <div style={styles.eyebrow}>Personal finance · runs locally in your browser</div>
            <h1 style={styles.h1}>Ledger.</h1>
          </div>
          <div style={styles.controlsRow}>
            <select style={styles.select} value={currency} onChange={(e) => setCurrency(e.target.value)}>
              {Object.keys(CURRENCIES).map((c) => <option key={c} value={c}>{c}</option>)}
            </select>
            <input ref={fileRef} type="file" accept=".csv,.xlsx,.xls" style={{ display: "none" }}
              onChange={(e) => e.target.files?.[0] && handleFile(e.target.files[0])} />
            <button style={styles.btn} onClick={() => fileRef.current?.click()}>
              <Upload size={14} /> Upload statement
            </button>
            {!isEmpty && (
              <>
                <button style={styles.btnGhostDark} onClick={exportCsv}><Download size={14} /> Export</button>
                <button style={styles.btnGhostDark} onClick={reset}><RefreshCw size={14} /> Reset</button>
              </>
            )}
          </div>
        </div>

        {error && (
          <div style={{
            display: "flex", gap: 10, alignItems: "flex-start", background: EXPENSE_SOFT,
            border: `1px solid ${EXPENSE}22`, color: EXPENSE, padding: "12px 16px", borderRadius: 8, marginBottom: 24, fontSize: 13.5,
          }}>
            <AlertCircle size={16} style={{ flexShrink: 0, marginTop: 1 }} />
            <div>{error} <button onClick={() => setError("")} style={{ background: "none", border: "none", color: EXPENSE, textDecoration: "underline", cursor: "pointer", fontSize: 13.5, padding: 0, marginLeft: 6 }}>Dismiss</button></div>
          </div>
        )}

        {isEmpty ? (
          <div
            className="bp-drop"
            onDrop={onDrop}
            onDragOver={(e) => e.preventDefault()}
            style={{
              border: `1.5px dashed ${LINE}`, borderRadius: 14, padding: "80px 32px",
              textAlign: "center", background: CARD,
            }}
          >
            <FileSpreadsheet size={34} color={GOLD} style={{ marginBottom: 16 }} />
            <div style={{ fontFamily: "'Fraunces', serif", fontSize: 24, marginBottom: 8 }}>Drop a bank statement here</div>
            <div style={{ color: MUTED, fontSize: 14, marginBottom: 24, maxWidth: 440, marginLeft: "auto", marginRight: "auto" }}>
              CSV or Excel, with columns for date, description and amount (or separate debit/credit columns).
              Everything is parsed and stored in your browser tab only — nothing leaves your machine.
            </div>
            <div style={{ display: "flex", gap: 10, justifyContent: "center" }}>
              <button style={styles.btn} onClick={() => fileRef.current?.click()}><Upload size={14} /> Choose file</button>
              <button style={styles.btnGhost} onClick={loadSample}>Try sample data</button>
            </div>
          </div>
        ) : (
          <>
            {uncategorizedCount > 0 && (
              <div style={{
                display: "flex", gap: 10, alignItems: "center", background: "#FBF7EC",
                border: `1px solid ${GOLD}33`, color: "#6B551F", padding: "11px 16px", borderRadius: 8, marginBottom: 24, fontSize: 13.5,
              }}>
                <AlertCircle size={15} />
                {uncategorizedCount} transaction{uncategorizedCount > 1 ? "s" : ""} landed in a generic "Other" category — review them in the table below.
              </div>
            )}

            {/* Summary cards */}
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 16, marginBottom: 32 }}>
              <SummaryCard label="Total income" value={fmtMoney(totals.income, symbol)} color={INCOME} icon={<TrendingUp size={16} />} />
              <SummaryCard label="Total expense" value={fmtMoney(totals.expense, symbol)} color={EXPENSE} icon={<TrendingDown size={16} />} />
              <NetCard value={fmtMoney(totals.net, symbol)} positive={totals.net >= 0} />
            </div>

            {/* Period tabs */}
            <div style={{ display: "flex", gap: 8, marginBottom: 20, alignItems: "center", flexWrap: "wrap" }}>
              {Object.keys(PERIODS).map((p) => (
                <button key={p} className={`bp-tab ${period === p ? "active" : ""}`} onClick={() => setPeriod(p)}>
                  {PERIODS[p].title}
                </button>
              ))}
              {period === "daily" && (
                <div style={{ display: "flex", alignItems: "center", gap: 8, marginLeft: 8 }}>
                  <span style={{ fontSize: 12.5, color: ON_DARK_MUTED, fontFamily: "'Inter',sans-serif" }}>From</span>
                  <input
                    type="date"
                    className="bp-inline"
                    value={dailyFrom}
                    min={dateExtent?.min}
                    max={dailyTo || dateExtent?.max}
                    onChange={(e) => setDailyFrom(e.target.value)}
                    style={{ padding: "7px 10px" }}
                  />
                  <span style={{ fontSize: 12.5, color: ON_DARK_MUTED, fontFamily: "'Inter',sans-serif" }}>To</span>
                  <input
                    type="date"
                    className="bp-inline"
                    value={dailyTo}
                    min={dailyFrom || dateExtent?.min}
                    max={dateExtent?.max}
                    onChange={(e) => setDailyTo(e.target.value)}
                    style={{ padding: "7px 10px" }}
                  />
                  {(dailyFrom || dailyTo) && (
                    <button
                      style={{ ...styles.btnGhostDark, padding: "7px 12px", fontSize: 12.5 }}
                      onClick={() => { setDailyFrom(""); setDailyTo(""); }}
                    >
                      <X size={12} /> Clear
                    </button>
                  )}
                </div>
              )}
            </div>
            {period === "daily" && !useDailyRange && (
              <div style={{ fontSize: 12.5, color: ON_DARK_MUTED, marginTop: -12, marginBottom: 16 }}>
                Showing the last {PERIODS.daily.take} days by default — pick a date range above to see a specific window.
              </div>
            )}

            {/* Trend chart */}
            <div style={{ background: CARD, border: `1px solid ${LINE}`, borderRadius: 12, padding: "20px 20px 8px", marginBottom: 20 }}>
              <div style={{ fontSize: 13, color: MUTED, marginBottom: 12, fontFamily: "'IBM Plex Mono', monospace", letterSpacing: "0.03em" }}>
                INCOME VS EXPENSE — {periodCfg.title.toUpperCase()}
              </div>
              <ResponsiveContainer width="100%" height={280}>
                <ComposedChart data={trendData} margin={{ top: 4, right: 8, left: 0, bottom: 4 }}>
                  <CartesianGrid stroke={LINE} vertical={false} />
                  <XAxis dataKey="label" tick={{ fontSize: 11, fill: MUTED }} axisLine={{ stroke: LINE }} tickLine={false} />
                  <YAxis tick={{ fontSize: 11, fill: MUTED }} axisLine={false} tickLine={false}
                    tickFormatter={(v) => `${symbol}${v >= 1000 ? (v / 1000).toFixed(0) + "k" : v}`} />
                  <Tooltip
                    contentStyle={{ borderRadius: 8, border: `1px solid ${LINE}`, fontSize: 12.5, fontFamily: "'Inter',sans-serif" }}
                    formatter={(v, name) => [fmtMoney(v, symbol), name === "income" ? "Income" : name === "expense" ? "Expense" : "Net"]}
                  />
                  <Legend wrapperStyle={{ fontSize: 12.5 }} formatter={(v) => (v === "income" ? "Income" : v === "expense" ? "Expense" : "Net")} />
                  <Bar dataKey="income" fill={INCOME} radius={[4, 4, 0, 0]} maxBarSize={34} />
                  <Bar dataKey="expense" fill={EXPENSE} radius={[4, 4, 0, 0]} maxBarSize={34} />
                  <Line type="monotone" dataKey="net" stroke={GOLD} strokeWidth={2} dot={{ r: 3, fill: GOLD }} />
                </ComposedChart>
              </ResponsiveContainer>
            </div>

            {/* Pie charts */}
            <div style={{ display: "grid", gridTemplateColumns: categoryBreakdown.income.length ? "1fr 1fr" : "1fr", gap: 16, marginBottom: 32 }}>
              <PieCard title="Expense by category" data={categoryBreakdown.expense} symbol={symbol} />
              {categoryBreakdown.income.length > 0 && (
                <PieCard title="Income by source" data={categoryBreakdown.income} symbol={symbol} />
              )}
            </div>

            {/* Transactions table */}
            <div style={{ background: CARD, border: `1px solid ${LINE}`, borderRadius: 12, overflow: "hidden" }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "16px 20px", borderBottom: `1px solid ${LINE}`, flexWrap: "wrap", gap: 12 }}>
                <div style={{ fontFamily: "'Fraunces', serif", fontSize: 18 }}>Transactions</div>
                <div style={{ display: "flex", gap: 10, alignItems: "center" }}>
                  <div style={{ position: "relative" }}>
                    <Search size={14} color={MUTED} style={{ position: "absolute", left: 11, top: 11 }} />
                    <input className="bp-search" placeholder="Search description or category" value={search} onChange={(e) => setSearch(e.target.value)} />
                  </div>
                  <select className="bp-inline" value={typeFilter} onChange={(e) => setTypeFilter(e.target.value)}>
                    <option value="all">All types</option>
                    <option value="income">Income only</option>
                    <option value="expense">Expense only</option>
                  </select>
                </div>
              </div>
              <div style={{ maxHeight: 460, overflowY: "auto" }}>
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13.5 }}>
                  <thead>
                    <tr style={{ textAlign: "left", color: MUTED, fontSize: 11.5, textTransform: "uppercase", letterSpacing: "0.05em", fontFamily: "'IBM Plex Mono', monospace" }}>
                      <th style={{ padding: "10px 20px" }}>Date</th>
                      <th style={{ padding: "10px 12px" }}>Description</th>
                      <th style={{ padding: "10px 12px" }}>Type</th>
                      <th style={{ padding: "10px 12px" }}>Category</th>
                      <th style={{ padding: "10px 20px", textAlign: "right" }}>Amount</th>
                      <th style={{ padding: "10px 20px 10px 8px", textAlign: "center", width: 40 }}></th>
                    </tr>
                  </thead>
                  <tbody>
                    {filteredTransactions.map((t) => (
                      <tr key={t.id} className="bp-row" style={{ borderTop: `1px solid ${LINE}` }}>
                        <td style={{ padding: "10px 20px", color: MUTED, whiteSpace: "nowrap", fontFamily: "'IBM Plex Mono', monospace", fontSize: 12.5 }}>
                          {t.date.toLocaleDateString(undefined, { day: "2-digit", month: "short", year: "numeric" })}
                        </td>
                        <td style={{ padding: "10px 12px", maxWidth: 260 }}>{t.description}</td>
                        <td style={{ padding: "10px 12px" }}>
                          <button
                            onClick={() => updateTransaction(t.id, { type: t.type === "income" ? "expense" : "income", category: autoCategorize(t.description, t.type === "income" ? "expense" : "income") })}
                            style={{
                              fontSize: 11.5, fontWeight: 500, padding: "3px 10px", borderRadius: 20, border: "none", cursor: "pointer",
                              background: t.type === "income" ? INCOME_SOFT : EXPENSE_SOFT,
                              color: t.type === "income" ? INCOME : EXPENSE,
                            }}
                            title="Click to flip income/expense"
                          >
                            {t.type === "income" ? "Income" : "Expense"}
                          </button>
                        </td>
                        <td style={{ padding: "10px 12px" }}>
                          <select
                            className="bp-inline"
                            value={t.category}
                            onChange={(e) => updateTransaction(t.id, { category: e.target.value })}
                          >
                            {(t.type === "income" ? INCOME_CATEGORIES : EXPENSE_CATEGORIES).map((c) => (
                              <option key={c} value={c}>{c}</option>
                            ))}
                            {!((t.type === "income" ? INCOME_CATEGORIES : EXPENSE_CATEGORIES).includes(t.category)) && (
                              <option value={t.category}>{t.category}</option>
                            )}
                          </select>
                        </td>
                        <td style={{ padding: "10px 20px", textAlign: "right", fontFamily: "'IBM Plex Mono', monospace", fontWeight: 500, color: t.type === "income" ? INCOME : EXPENSE, whiteSpace: "nowrap" }}>
                          {t.type === "income" ? "+" : "−"}{fmtMoney(t.amount, symbol)}
                        </td>
                        <td style={{ padding: "10px 8px", textAlign: "center" }}>
                          <button
                            onClick={() => deleteTransaction(t.id)}
                            title="Delete this transaction"
                            style={{
                              background: "none", border: "none", cursor: "pointer", color: MUTED,
                              padding: 4, borderRadius: 6, display: "inline-flex",
                            }}
                            onMouseEnter={(e) => { e.currentTarget.style.color = EXPENSE; }}
                            onMouseLeave={(e) => { e.currentTarget.style.color = MUTED; }}
                          >
                            <Trash2 size={14} />
                          </button>
                        </td>
                      </tr>
                    ))}
                    {filteredTransactions.length === 0 && (
                      <tr><td colSpan={6} style={{ padding: "32px 20px", textAlign: "center", color: MUTED }}>No transactions match your filters.</td></tr>
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          </>
        )}
      </div>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Small components
// ---------------------------------------------------------------------------
function SummaryCard({ label, value, color, icon }) {
  return (
    <div style={{ background: CARD, border: `1px solid ${LINE}`, borderRadius: 12, padding: "18px 20px" }}>
      <div style={{ display: "flex", alignItems: "center", gap: 8, color: MUTED, fontSize: 12.5, marginBottom: 10, fontFamily: "'Inter',sans-serif" }}>
        <span style={{ color }}>{icon}</span>{label}
      </div>
      <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 24, fontWeight: 500, color }}>{value}</div>
    </div>
  );
}

function NetCard({ value, positive }) {
  return (
    <div style={{
      background: CARD, border: `1px solid ${LINE}`, borderRadius: 12, padding: "18px 20px",
      display: "flex", alignItems: "center", justifyContent: "space-between",
    }}>
      <div>
        <div style={{ display: "flex", alignItems: "center", gap: 8, color: MUTED, fontSize: 12.5, marginBottom: 10 }}>
          <Scale size={16} color={GOLD} />Net balance
        </div>
        <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 24, fontWeight: 500, color: positive ? INCOME : EXPENSE }}>{value}</div>
      </div>
      <div style={{
        width: 54, height: 54, borderRadius: "50%", border: `2px double ${GOLD}`,
        display: "flex", alignItems: "center", justifyContent: "center", transform: "rotate(-8deg)",
        fontFamily: "'IBM Plex Mono', monospace", fontSize: 9.5, letterSpacing: "0.05em", color: GOLD, textAlign: "center", flexShrink: 0,
      }}>
        {positive ? "SURPLUS" : "DEFICIT"}
      </div>
    </div>
  );
}

function PieCard({ title, data, symbol }) {
  const total = data.reduce((s, d) => s + d.value, 0);
  return (
    <div style={{ background: CARD, border: `1px solid ${LINE}`, borderRadius: 12, padding: 20 }}>
      <div style={{ fontSize: 13, color: MUTED, marginBottom: 8, fontFamily: "'IBM Plex Mono', monospace", letterSpacing: "0.03em" }}>
        {title.toUpperCase()}
      </div>
      {data.length === 0 ? (
        <div style={{ color: MUTED, fontSize: 13, padding: "40px 0", textAlign: "center" }}>No data for this period.</div>
      ) : (
        <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
          <ResponsiveContainer width="55%" height={220}>
            <PieChart>
              <Pie data={data} dataKey="value" nameKey="name" innerRadius={55} outerRadius={90} paddingAngle={2}>
                {data.map((_, i) => <Cell key={i} fill={CAT_PALETTE[i % CAT_PALETTE.length]} />)}
              </Pie>
              <Tooltip formatter={(v) => fmtMoney(v, symbol)} contentStyle={{ borderRadius: 8, border: `1px solid ${LINE}`, fontSize: 12.5 }} />
            </PieChart>
          </ResponsiveContainer>
          <div style={{ flex: 1, display: "flex", flexDirection: "column", gap: 6, maxHeight: 220, overflowY: "auto" }}>
            {data.map((d, i) => (
              <div key={d.name} style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 12.5 }}>
                <span style={{ width: 9, height: 9, borderRadius: 2, background: CAT_PALETTE[i % CAT_PALETTE.length], flexShrink: 0 }} />
                <span style={{ flex: 1, color: INK }}>{d.name}</span>
                <span style={{ color: MUTED, fontFamily: "'IBM Plex Mono', monospace", fontSize: 11.5 }}>
                  {total ? Math.round((d.value / total) * 100) : 0}%
                </span>
              </div>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}
