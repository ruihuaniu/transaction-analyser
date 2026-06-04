import React, { useState, useMemo, useCallback, useRef } from "react";
import {
  useReactTable, getCoreRowModel, getSortedRowModel,
  getFilteredRowModel, getPaginationRowModel, flexRender, createColumnHelper,
} from "@tanstack/react-table";
import { useQueryClient, QueryClient, QueryClientProvider } from "@tanstack/react-query";
import * as XLSX from "xlsx";
import Papa from "papaparse";
import {
  AreaChart, Area, BarChart, Bar, PieChart, Pie, Cell,
  XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, Legend,
} from "recharts";

const queryClient = new QueryClient();

// ─────────────────────────────────────────────────────────────────────────────
// COLUMN FINDER
// Strategy: exact match first (case-insensitive, trimmed), then "starts with",
// then "includes". This prevents "Debit Amount" accidentally matching a
// "Credit Amount" column when both are present.
// ─────────────────────────────────────────────────────────────────────────────
function findColumn(keys, candidates) {
  const norm = s => s.toLowerCase().trim();
  // 1. exact match
  for (const c of candidates) {
    const k = keys.find(k => norm(k) === norm(c));
    if (k) return k;
  }
  // 2. starts-with match
  for (const c of candidates) {
    const k = keys.find(k => norm(k).startsWith(norm(c)));
    if (k) return k;
  }
  // 3. includes match (last resort)
  for (const c of candidates) {
    const k = keys.find(k => norm(k).includes(norm(c)));
    if (k) return k;
  }
  return null;
}

// ─────────────────────────────────────────────────────────────────────────────
// AMOUNT PARSER
// Bank CSVs: "Debit Amount" and "Credit Amount" columns hold positive values.
// A blank cell means $0 for that direction.
// Some banks export negative numbers in the debit column for reversals — we
// take abs() so they don't cancel out the sum.
// ─────────────────────────────────────────────────────────────────────────────
function parseAmount(val) {
  if (val === undefined || val === null || val === "") return 0;
  const n = parseFloat(String(val).replace(/[$,\s]/g, ""));
  return isNaN(n) ? 0 : Math.abs(n);
}

// ─────────────────────────────────────────────────────────────────────────────
// NORMALIZE ROWS
// Called once when a file is loaded. Extracts the four expected columns by
// name (flexible matching) and computes derived fields.
// ─────────────────────────────────────────────────────────────────────────────
function normalizeRows(rawRows) {
  if (!rawRows.length) return [];
  const keys = Object.keys(rawRows[0]);

  // Find each column key once (not per row) — deterministic
  const dateKey      = findColumn(keys, ["date"]);
  const narrativeKey = findColumn(keys, ["narrative", "description", "details", "merchant", "narr"]);
  // IMPORTANT: look for the more-specific "debit amount" BEFORE plain "debit"
  // and "credit amount" BEFORE plain "credit" to avoid cross-matching.
  const debitKey     = findColumn(keys, ["debit amount", "debit amt", "debit", "dr amount", "dr"]);
  const creditKey    = findColumn(keys, ["credit amount", "credit amt", "credit", "cr amount", "cr"]);

  // Sanity: if debitKey === creditKey something went wrong — fall back to null
  const safeDebitKey  = debitKey !== creditKey ? debitKey  : null;
  const safeCreditKey = debitKey !== creditKey ? creditKey : null;

  return rawRows
    .map((row, idx) => {
      const dateRaw  = dateKey      ? row[dateKey]      : "";
      const narrative = narrativeKey ? String(row[narrativeKey] ?? "") : "";
      const debit    = parseAmount(safeDebitKey  ? row[safeDebitKey]  : "");
      const credit   = parseAmount(safeCreditKey ? row[safeCreditKey] : "");

      const dateObj  = parseDateStr(dateRaw);
      return {
        id: idx,
        date:     dateObj,
        dateStr:  dateObj ? dateObj.toLocaleDateString("en-AU") : String(dateRaw ?? ""),
        dateSort: dateObj ? dateObj.getTime() : 0,
        narrative,
        debit,   // ← "Debit Amount" column value
        credit,  // ← "Credit Amount" column value
        category: categorize(narrative),
      };
    })
    .filter(r => r.narrative || r.debit > 0 || r.credit > 0);
}

function parseDateStr(val) {
  if (!val) return null;
  if (typeof val === "number") {
    const d = XLSX.SSF.parse_date_code(val);
    if (d) return new Date(d.y, d.m - 1, d.d);
  }
  const d = new Date(val);
  return isNaN(d.getTime()) ? null : d;
}

function fmt(n) {
  return new Intl.NumberFormat("en-AU", {
    style: "currency", currency: "AUD", minimumFractionDigits: 2,
  }).format(n ?? 0);
}

// ─────────────────────────────────────────────────────────────────────────────
// CATEGORY RULES — order matters: more specific rules first
// ─────────────────────────────────────────────────────────────────────────────
const CATEGORY_RULES = [
  // Income
  { keywords: ["salary","payroll","wage","income","pay slip","direct credit","employer"], label: "Income" },
  // Housing
  { keywords: ["rent","mortgage","landlord","lease","strata","body corporate","real estate","property management"], label: "Housing" },
  // Hardware & Home
  { keywords: ["bunnings","mitre 10","masters home","hardware","timber","plumbing","electrical supplies","paint","ikea","freedom furniture","harvey norman","the good guys","jb hi-fi","bing lee","officeworks","pottery barn","bed bath","kmart","target","big w","spotlight","lincraft","howards storage"], label: "Hardware & Home" },
  // Groceries — before dining so "supermarket" wins over restaurants
  { keywords: ["woolworths","woolworth","coles","aldi","iga","harris farm","foodland","spar","drakes","costco","grocery","supermarket","fruit shop","butcher","bakery","deli","fresh market"], label: "Groceries" },
  // Dining
  { keywords: ["restaurant","cafe","coffee","hungry jacks","mcdonald","kfc","subway","domino","pizza","burger","sushi","ramen","thai","chinese","indian","italian","brunch","bistro","bar & grill","eatery","takeaway","takeout","uber eats","doordash","menulog","deliveroo","dining"], label: "Dining" },
  // Transport
  { keywords: ["uber","lyft","ola","didi","taxi","transport nsw","translink","myki","opal","train","bus","tram","metro","ferry","fuel","petrol","shell","bp ","caltex","ampol","parking","car park","toll","linkt","e-toll","roam express"], label: "Transport" },
  // Health & Fitness
  { keywords: ["gym","fitness","crossfit","yoga","pilates","f45","anytime fitness","planet fitness","snap fitness","doctor","gp ","medical centre","hospital","pharmacy","chemist","priceline","terry white","healthdirect","dentist","optical","physio","pathology"], label: "Health & Fitness" },
  // Subscriptions
  { keywords: ["netflix","stan","binge","paramount","disney","apple tv","prime video","spotify","apple music","youtube premium","audible","kindle","adobe","microsoft 365","google one","icloud","dropbox","canva subscription","hbo"], label: "Subscriptions" },
  // Utilities
  { keywords: ["electricity","energy australia","agl ","origin energy","simply energy","powershop","water ","sydney water","yarra valley water","gas ","ausnet","jemena","internet","nbn","telstra","optus","vodafone","tpg","aussie broadband","aussie bb","mobile plan","phone bill"], label: "Utilities" },
  // Travel
  { keywords: ["qantas","virgin australia","jetstar","rex airline","air new zealand","singapore air","emirates","flight","airfare","airbnb","booking.com","hotels.com","expedia","wotif","agoda","hotel","motel","resort","hostel","holiday","vacation","travel insurance","visa fee"], label: "Travel" },
  // Education
  { keywords: ["university","tafe","school fee","tuition","udemy","coursera","skillshare","linkedin learning","masterclass","textbook","stationery"], label: "Education" },
  // Entertainment
  { keywords: ["cinema","event cinema","hoyts","village cinema","reading cinema","concert","ticketek","ticketmaster","museum","gallery","zoo","theme park","entertainment"], label: "Entertainment" },
  // Shopping (catch-all online/fashion after hardware & home already handled)
  { keywords: ["amazon","ebay","etsy","aliexpress","shein","asos","zara","h&m","uniqlo","cotton on","factorie","rivers","lowes","tarocash","myer","david jones","the iconic","net-a-porter","shop","store","clothing","fashion","jewellery"], label: "Shopping" },
  // Investments & Finance
  { keywords: ["commsec","nabtrade","selfwealth","stake","raiz","spaceship","vanguard","brokerage","share purchase","dividends received","super","superannuation","insurance","life insurance","car insurance","home insurance","income protection"], label: "Insurance & Finance" },
  // Banking & Fees
  { keywords: ["bank fee","account fee","monthly fee","transaction fee","dishonour","overdrawn","atm fee","interest charge","late payment","annual fee","bpay","eft","wire transfer"], label: "Banking & Fees" },
  // Government & Tax
  { keywords: ["ato ","tax office","centrelink","medicare","services australia","council rates","land tax","stamp duty","fine ","toll infringement","government","dept of"], label: "Government & Tax" },
  // Charity
  { keywords: ["donate","donation","charity","foundation","red cross","cancer council","oxfam","beyond blue","lifeline","salvos"], label: "Charity" },
];

function categorize(narrative) {
  if (!narrative) return "Other";
  const lower = narrative.toLowerCase();
  for (const rule of CATEGORY_RULES) {
    if (rule.keywords.some(k => lower.includes(k))) return rule.label;
  }
  return "Other";
}

// ─────────────────────────────────────────────────────────────────────────────
// PALETTE & HELPERS
// ─────────────────────────────────────────────────────────────────────────────
const PALETTE = [
  "#10b981","#3b82f6","#f43f5e","#f59e0b","#8b5cf6","#ec4899",
  "#06b6d4","#84cc16","#f97316","#d946ef","#14b8a6","#6366f1",
  "#a78bfa","#fb923c","#34d399","#60a5fa","#f472b6","#a3e635",
];

const columnHelper = createColumnHelper();

// ─────────────────────────────────────────────────────────────────────────────
// THEME
// ─────────────────────────────────────────────────────────────────────────────
function makeTheme(dark) {
  return dark ? {
    bg: "#0d0d1a",
    bgCard: "rgba(255,255,255,0.03)",
    border: "rgba(255,255,255,0.08)",
    borderStrong: "rgba(255,255,255,0.13)",
    text: "#e8e8e8",
    textMuted: "#888",
    textFaint: "#333",
    headerBg: "rgba(13,13,26,0.9)",
    tableHead: "rgba(255,255,255,0.04)",
    tableRow: "rgba(255,255,255,0.02)",
    tableRowHover: "rgba(110,231,183,0.07)",
    inputBg: "rgba(255,255,255,0.05)",
    btnBg: "rgba(255,255,255,0.06)",
    btnColor: "#ccc",
    btnDisabled: "#333",
    scrollThumb: "rgba(255,255,255,0.12)",
    chartGrid: "rgba(255,255,255,0.04)",
    chartTick: "#555",
    tooltipBg: "#1a1a2e",
    accent: "#6EE7B7",
    accentDim: "rgba(110,231,183,0.12)",
    accentBorder: "rgba(110,231,183,0.25)",
    debit: "#f87171",
    credit: "#34d399",
    glow: "rgba(110,231,183,0.06)",
  } : {
    bg: "#f4f5f7",
    bgCard: "#ffffff",
    border: "rgba(0,0,0,0.08)",
    borderStrong: "rgba(0,0,0,0.14)",
    text: "#111827",
    textMuted: "#6b7280",
    textFaint: "#d1d5db",
    headerBg: "rgba(244,245,247,0.93)",
    tableHead: "#f9fafb",
    tableRow: "#fafafa",
    tableRowHover: "rgba(16,185,129,0.05)",
    inputBg: "#ffffff",
    btnBg: "#f3f4f6",
    btnColor: "#374151",
    btnDisabled: "#d1d5db",
    scrollThumb: "rgba(0,0,0,0.15)",
    chartGrid: "rgba(0,0,0,0.05)",
    chartTick: "#9ca3af",
    tooltipBg: "#ffffff",
    accent: "#059669",
    accentDim: "rgba(5,150,105,0.1)",
    accentBorder: "rgba(5,150,105,0.3)",
    debit: "#ef4444",
    credit: "#059669",
    glow: "rgba(5,150,105,0.04)",
  };
}

// ─────────────────────────────────────────────────────────────────────────────
// STAT CARD
// ─────────────────────────────────────────────────────────────────────────────
function StatCard({ label, value, sub, accent, th }) {
  return (
    <div style={{
      background: th.bgCard, border: `1px solid ${th.border}`,
      borderRadius: 16, padding: "20px 24px",
      display: "flex", flexDirection: "column", gap: 4,
      flex: "1 1 180px", minWidth: 160,
      position: "relative", overflow: "hidden",
      boxShadow: "0 1px 4px rgba(0,0,0,0.06)",
    }}>
      <div style={{ position: "absolute", top: 0, left: 0, right: 0, height: 3, background: accent, borderRadius: "16px 16px 0 0" }} />
      <span style={{ fontSize: 11, fontWeight: 600, letterSpacing: "0.12em", color: th.textMuted, textTransform: "uppercase" }}>{label}</span>
      <span style={{ fontSize: 22, fontWeight: 700, color: th.text, fontFamily: "'DM Mono', monospace", marginTop: 4 }}>{value}</span>
      {sub && <span style={{ fontSize: 12, color: th.textMuted }}>{sub}</span>}
    </div>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// SEARCH INPUT
// ─────────────────────────────────────────────────────────────────────────────
function SearchInput({ value, onChange, placeholder, th }) {
  return (
    <div style={{ position: "relative", display: "inline-flex", alignItems: "center" }}>
      <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke={th.textMuted} strokeWidth="2.5"
        style={{ position: "absolute", left: 10, pointerEvents: "none" }}>
        <circle cx="11" cy="11" r="8" /><path d="m21 21-4.35-4.35" />
      </svg>
      <input value={value} onChange={e => onChange(e.target.value)}
        placeholder={placeholder || "Search…"}
        style={{
          background: th.inputBg, border: `1px solid ${th.border}`,
          borderRadius: 8, padding: "7px 12px 7px 30px",
          color: th.text, fontSize: 13, outline: "none", width: 220, fontFamily: "inherit",
        }} />
    </div>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// DATA TABLE (TanStack Table)
// ─────────────────────────────────────────────────────────────────────────────
function DataTable({ data, columns, th }) {
  const [sorting, setSorting] = useState([]);
  const [globalFilter, setGlobalFilter] = useState("");
  const [pagination, setPagination] = useState({ pageIndex: 0, pageSize: 15 });

  const table = useReactTable({
    data, columns,
    state: { sorting, globalFilter, pagination },
    onSortingChange: setSorting,
    onGlobalFilterChange: setGlobalFilter,
    onPaginationChange: setPagination,
    getCoreRowModel: getCoreRowModel(),
    getSortedRowModel: getSortedRowModel(),
    getFilteredRowModel: getFilteredRowModel(),
    getPaginationRowModel: getPaginationRowModel(),
  });

  const rows = table.getRowModel().rows;

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8 }}>
        <SearchInput
          value={globalFilter}
          onChange={v => { setGlobalFilter(v); setPagination(p => ({ ...p, pageIndex: 0 })); }}
          placeholder="Search all columns…"
          th={th}
        />
        <span style={{ fontSize: 12, color: th.textMuted }}>
          {table.getFilteredRowModel().rows.length.toLocaleString()} rows
        </span>
      </div>
      <div style={{ overflowX: "auto", borderRadius: 12, border: `1px solid ${th.border}` }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13, minWidth: 500 }}>
          <thead>
            {table.getHeaderGroups().map(hg => (
              <tr key={hg.id} style={{ borderBottom: `1px solid ${th.borderStrong}`, background: th.tableHead }}>
                {hg.headers.map(header => (
                  <th key={header.id} onClick={header.column.getToggleSortingHandler()}
                    style={{
                      padding: "10px 14px", textAlign: "left", fontWeight: 600,
                      fontSize: 11, letterSpacing: "0.08em", textTransform: "uppercase",
                      color: th.textMuted, cursor: header.column.getCanSort() ? "pointer" : "default",
                      userSelect: "none", whiteSpace: "nowrap",
                    }}>
                    <div style={{ display: "flex", alignItems: "center", gap: 4 }}>
                      {flexRender(header.column.columnDef.header, header.getContext())}
                      {header.column.getIsSorted() === "asc" && " ↑"}
                      {header.column.getIsSorted() === "desc" && " ↓"}
                      {!header.column.getIsSorted() && header.column.getCanSort() && <span style={{ opacity: 0.3 }}> ⇅</span>}
                    </div>
                  </th>
                ))}
              </tr>
            ))}
          </thead>
          <tbody>
            {rows.length === 0
              ? <tr><td colSpan={columns.length} style={{ padding: "40px", textAlign: "center", color: th.textMuted }}>No results</td></tr>
              : rows.map((row, i) => (
                <tr key={row.id}
                  style={{ borderBottom: `1px solid ${th.border}`, background: i % 2 === 0 ? "transparent" : th.tableRow, transition: "background 0.1s" }}
                  onMouseEnter={e => { e.currentTarget.style.background = th.tableRowHover; }}
                  onMouseLeave={e => { e.currentTarget.style.background = i % 2 === 0 ? "transparent" : th.tableRow; }}>
                  {row.getVisibleCells().map(cell => (
                    <td key={cell.id} style={{ padding: "9px 14px", color: th.text }}>
                      {flexRender(cell.column.columnDef.cell, cell.getContext())}
                    </td>
                  ))}
                </tr>
              ))
            }
          </tbody>
        </table>
      </div>
      <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap", justifyContent: "space-between" }}>
        <div style={{ display: "flex", gap: 4 }}>
          {[
            { label: "«", fn: () => table.setPageIndex(0), disabled: !table.getCanPreviousPage() },
            { label: "‹", fn: () => table.previousPage(), disabled: !table.getCanPreviousPage() },
            { label: "›", fn: () => table.nextPage(), disabled: !table.getCanNextPage() },
            { label: "»", fn: () => table.setPageIndex(table.getPageCount() - 1), disabled: !table.getCanNextPage() },
          ].map((btn, idx) => (
            <button key={idx} onClick={btn.fn} disabled={btn.disabled}
              style={{
                background: th.btnBg, border: `1px solid ${th.border}`,
                borderRadius: 6, padding: "4px 10px", color: btn.disabled ? th.btnDisabled : th.btnColor,
                cursor: btn.disabled ? "default" : "pointer", fontSize: 14, fontFamily: "inherit",
              }}>{btn.label}</button>
          ))}
        </div>
        <span style={{ fontSize: 12, color: th.textMuted }}>
          Page {table.getState().pagination.pageIndex + 1} / {Math.max(1, table.getPageCount())}
        </span>
        <select value={table.getState().pagination.pageSize} onChange={e => table.setPageSize(Number(e.target.value))}
          style={{ background: th.inputBg, border: `1px solid ${th.border}`, borderRadius: 6, padding: "4px 8px", color: th.text, fontSize: 12, fontFamily: "inherit" }}>
          {[10, 15, 25, 50, 100].map(s => <option key={s} value={s}>{s} / page</option>)}
        </select>
      </div>
    </div>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// CHART TOOLTIP
// ─────────────────────────────────────────────────────────────────────────────
function ChartTooltip({ active, payload, label, th }) {
  if (!active || !payload?.length) return null;
  return (
    <div style={{ background: th.tooltipBg, border: `1px solid ${th.border}`, borderRadius: 10, padding: "10px 14px", fontSize: 12, color: th.text, boxShadow: "0 4px 16px rgba(0,0,0,0.12)" }}>
      <div style={{ fontWeight: 700, marginBottom: 4, color: th.accent }}>{label}</div>
      {payload.map((p, i) => (
        <div key={i} style={{ display: "flex", gap: 12, justifyContent: "space-between" }}>
          <span style={{ color: p.color }}>{p.name}</span>
          <span style={{ fontFamily: "'DM Mono', monospace" }}>{fmt(p.value)}</span>
        </div>
      ))}
    </div>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// UPLOAD SCREEN
// ─────────────────────────────────────────────────────────────────────────────
function UploadScreen({ onData, dark, toggleDark }) {
  const [dragging, setDragging] = useState(false);
  const [error, setError] = useState(null);
  const [loading, setLoading] = useState(false);
  const inputRef = useRef();
  const th = makeTheme(dark);

  const processFile = useCallback((file) => {
    if (!file) return;
    setError(null); setLoading(true);
    const ext = file.name.split(".").pop().toLowerCase();

    if (ext === "csv") {
      Papa.parse(file, {
        header: true, skipEmptyLines: true,
        complete: (result) => {
          setLoading(false);
          if (!result.data.length) { setError("CSV appears empty."); return; }
          onData(normalizeRows(result.data));
        },
        error: () => { setLoading(false); setError("Failed to parse CSV."); },
      });
    } else if (["xlsx", "xls"].includes(ext)) {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const wb = XLSX.read(e.target.result, { type: "binary", cellDates: false });
          const ws = wb.Sheets[wb.SheetNames[0]];
          const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });
          setLoading(false);
          if (!rows.length) { setError("Excel sheet appears empty."); return; }
          onData(normalizeRows(rows));
        } catch { setLoading(false); setError("Failed to parse Excel file."); }
      };
      reader.readAsBinaryString(file);
    } else {
      setLoading(false);
      setError("Please upload a .csv, .xlsx, or .xls file.");
    }
  }, [onData]);

  return (
    <div style={{ minHeight: "100vh", background: th.bg, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", fontFamily: "'DM Sans', sans-serif", padding: 24, transition: "background 0.3s" }}>
      <div style={{ position: "fixed", top: 0, left: 0, right: 0, height: 400, background: `radial-gradient(ellipse at 50% -10%, ${th.glow} 0%, transparent 70%)`, pointerEvents: "none" }} />
      <button onClick={toggleDark} title="Toggle light/dark" style={{ position: "fixed", top: 20, right: 24, background: th.bgCard, border: `1px solid ${th.border}`, borderRadius: 10, padding: "7px 13px", cursor: "pointer", fontSize: 16, fontFamily: "inherit" }}>
        {dark ? "☀️" : "🌙"}
      </button>
      <div style={{ textAlign: "center", marginBottom: 48 }}>
        <div style={{ display: "inline-flex", alignItems: "center", gap: 8, background: th.accentDim, border: `1px solid ${th.accentBorder}`, borderRadius: 100, padding: "4px 14px", marginBottom: 20, fontSize: 12, fontWeight: 600, letterSpacing: "0.1em", color: th.accent, textTransform: "uppercase" }}>
          ◆ Transaction Analyser
        </div>
        <h1 style={{ fontSize: "clamp(32px,5vw,52px)", fontWeight: 800, color: th.text, margin: 0, lineHeight: 1.1, letterSpacing: "-0.02em" }}>
          Understand your<br /><span style={{ color: th.accent }}>financial story</span>
        </h1>
        <p style={{ color: th.textMuted, marginTop: 16, fontSize: 15, maxWidth: 400, lineHeight: 1.6 }}>
          Upload a bank statement to instantly visualise spending, categorise transactions, and spot patterns.
        </p>
      </div>

      <div
        onDragOver={e => { e.preventDefault(); setDragging(true); }}
        onDragLeave={() => setDragging(false)}
        onDrop={e => { e.preventDefault(); setDragging(false); processFile(e.dataTransfer.files[0]); }}
        onClick={() => inputRef.current?.click()}
        style={{
          width: "100%", maxWidth: 500,
          border: `2px dashed ${dragging ? th.accent : th.borderStrong}`,
          borderRadius: 20, padding: "48px 32px", textAlign: "center",
          cursor: "pointer", background: dragging ? th.accentDim : th.bgCard,
          transition: "all 0.2s", boxShadow: "0 2px 12px rgba(0,0,0,0.06)",
        }}>
        <input ref={inputRef} type="file" accept=".csv,.xlsx,.xls" style={{ display: "none" }} onChange={e => processFile(e.target.files[0])} />
        {loading
          ? <div style={{ color: th.accent, fontSize: 15 }}>
              <div style={{ fontSize: 32, marginBottom: 12, display: "inline-block", animation: "spin 1s linear infinite" }}>⟳</div>
              <br />Processing file…
            </div>
          : <>
              <div style={{ fontSize: 40, marginBottom: 16 }}>📂</div>
              <div style={{ color: th.text, fontWeight: 600, fontSize: 16, marginBottom: 6 }}>Drop your file here</div>
              <div style={{ color: th.textMuted, fontSize: 13 }}>or click to browse — CSV, XLSX, XLS</div>
            </>
        }
      </div>
      {error && (
        <div style={{ marginTop: 16, color: th.debit, fontSize: 13, background: `${th.debit}18`, border: `1px solid ${th.debit}40`, borderRadius: 8, padding: "8px 16px" }}>{error}</div>
      )}
      <div style={{ marginTop: 28, fontSize: 12, color: th.textMuted, textAlign: "center" }}>
        Expected columns: <span style={{ color: th.accent }}>Date · Narrative · Debit Amount · Credit Amount</span>
      </div>
      <style>{`@keyframes spin { to { transform: rotate(360deg); } }`}</style>
    </div>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// DASHBOARD
// ─────────────────────────────────────────────────────────────────────────────
function Dashboard({ transactions, onReset, dark, toggleDark }) {
  const [activeCategory, setActiveCategory] = useState(null);
  const [activeTab, setActiveTab] = useState("overview");
  const th = makeTheme(dark);

  // ── STATS ─────────────────────────────────────────────────────────────────
  // totalSpent   = sum of every row's debit  field (= "Debit Amount" column)
  // totalReceived= sum of every row's credit field (= "Credit Amount" column)
  const stats = useMemo(() => {
    let totalSpent = 0;
    let totalReceived = 0;
    const catMap = {};
    const monthMap = {};

    for (const t of transactions) {
      totalSpent    += t.debit;   // Debit Amount
      totalReceived += t.credit;  // Credit Amount

      // Category accumulation
      if (!catMap[t.category]) catMap[t.category] = { label: t.category, debit: 0, credit: 0, count: 0 };
      catMap[t.category].debit  += t.debit;
      catMap[t.category].credit += t.credit;
      catMap[t.category].count  += 1;

      // Monthly accumulation
      if (t.date) {
        const key   = `${t.date.getFullYear()}-${String(t.date.getMonth() + 1).padStart(2, "0")}`;
        const label = t.date.toLocaleDateString("en-AU", { month: "short", year: "2-digit" });
        if (!monthMap[key]) monthMap[key] = { month: key, label, debit: 0, credit: 0 };
        monthMap[key].debit  += t.debit;
        monthMap[key].credit += t.credit;
      }
    }

    const net        = totalReceived - totalSpent;
    const categories = Object.values(catMap).sort((a, b) => b.debit - a.debit);
    const monthly    = Object.values(monthMap).sort((a, b) => a.month.localeCompare(b.month)).slice(-12);

    return { totalSpent, totalReceived, net, categories, monthly };
  }, [transactions]);

  const filteredTx = useMemo(() =>
    activeCategory ? transactions.filter(t => t.category === activeCategory) : transactions,
    [transactions, activeCategory]
  );

  // ── COLUMNS ───────────────────────────────────────────────────────────────
  const txColumns = useMemo(() => [
    columnHelper.accessor("dateStr", {
      header: "Date",
      sortingFn: (a, b) => a.original.dateSort - b.original.dateSort,
      cell: info => <span style={{ color: th.textMuted, fontFamily: "'DM Mono', monospace", fontSize: 12 }}>{info.getValue()}</span>,
    }),
    columnHelper.accessor("narrative", {
      header: "Narrative",
      cell: info => <span style={{ maxWidth: 280, display: "block", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{info.getValue()}</span>,
    }),
    columnHelper.accessor("category", {
      header: "Category",
      cell: info => {
        const idx = stats.categories.findIndex(c => c.label === info.getValue());
        const col = PALETTE[idx % PALETTE.length];
        return <span style={{ background: `${col}22`, color: col, border: `1px solid ${col}44`, borderRadius: 100, padding: "2px 10px", fontSize: 11, fontWeight: 600, whiteSpace: "nowrap" }}>{info.getValue()}</span>;
      },
    }),
    columnHelper.accessor("debit", {
      header: "Debit Amount",
      cell: info => info.getValue() > 0
        ? <span style={{ color: th.debit, fontFamily: "'DM Mono', monospace" }}>{fmt(info.getValue())}</span>
        : <span style={{ color: th.textFaint }}>—</span>,
    }),
    columnHelper.accessor("credit", {
      header: "Credit Amount",
      cell: info => info.getValue() > 0
        ? <span style={{ color: th.credit, fontFamily: "'DM Mono', monospace" }}>{fmt(info.getValue())}</span>
        : <span style={{ color: th.textFaint }}>—</span>,
    }),
  ], [stats.categories, th]);

  const catColumns = useMemo(() => [
    columnHelper.accessor("label", {
      header: "Category",
      cell: info => {
        const idx = stats.categories.findIndex(c => c.label === info.getValue());
        return <span style={{ color: PALETTE[idx % PALETTE.length], fontWeight: 600 }}>{info.getValue()}</span>;
      },
    }),
    columnHelper.accessor("count", {
      header: "Transactions",
      cell: info => <span style={{ fontFamily: "'DM Mono', monospace" }}>{info.getValue()}</span>,
    }),
    columnHelper.accessor("debit", {
      header: "Total Spent",
      cell: info => <span style={{ color: th.debit, fontFamily: "'DM Mono', monospace" }}>{fmt(info.getValue())}</span>,
    }),
    columnHelper.accessor("credit", {
      header: "Total Received",
      cell: info => info.getValue() > 0
        ? <span style={{ color: th.credit, fontFamily: "'DM Mono', monospace" }}>{fmt(info.getValue())}</span>
        : <span style={{ color: th.textFaint }}>—</span>,
    }),
    columnHelper.display({
      id: "view", header: "",
      cell: ({ row }) => (
        <button onClick={e => { e.stopPropagation(); setActiveCategory(row.original.label); setActiveTab("transactions"); }}
          style={{ background: th.accentDim, color: th.accent, border: `1px solid ${th.accentBorder}`, borderRadius: 6, padding: "3px 10px", fontSize: 11, cursor: "pointer", fontFamily: "inherit", fontWeight: 600 }}>
          View →
        </button>
      ),
    }),
  ], [stats.categories, th]);

  const ctProps = { th };

  return (
    <div style={{ minHeight: "100vh", background: th.bg, fontFamily: "'DM Sans', sans-serif", color: th.text, transition: "background 0.3s, color 0.3s" }}>
      <div style={{ position: "fixed", top: 0, left: 0, right: 0, height: 300, background: `radial-gradient(ellipse at 50% -10%, ${th.glow} 0%, transparent 70%)`, pointerEvents: "none" }} />

      {/* ── Header ── */}
      <header style={{
        display: "flex", alignItems: "center", justifyContent: "space-between",
        padding: "16px 28px", borderBottom: `1px solid ${th.border}`,
        position: "sticky", top: 0, zIndex: 100,
        background: th.headerBg, backdropFilter: "blur(16px)",
      }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <span style={{ background: th.accent, color: dark ? "#0d0d1a" : "#fff", width: 28, height: 28, borderRadius: 8, display: "flex", alignItems: "center", justifyContent: "center", fontWeight: 900, fontSize: 14 }}>$</span>
          <span style={{ fontWeight: 700, fontSize: 16, letterSpacing: "-0.01em" }}>TxAnalyser</span>
        </div>
        <div style={{ display: "flex", gap: 4 }}>
          {["overview", "transactions"].map(tab => (
            <button key={tab} onClick={() => setActiveTab(tab)}
              style={{
                background: activeTab === tab ? th.accentDim : "transparent",
                color: activeTab === tab ? th.accent : th.textMuted,
                border: activeTab === tab ? `1px solid ${th.accentBorder}` : "1px solid transparent",
                borderRadius: 8, padding: "6px 16px", fontSize: 13, fontWeight: 600,
                cursor: "pointer", fontFamily: "inherit", transition: "all 0.15s", textTransform: "capitalize",
              }}>{tab}</button>
          ))}
        </div>
        <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
          <button onClick={toggleDark} title="Toggle theme" style={{ background: th.btnBg, border: `1px solid ${th.border}`, borderRadius: 8, padding: "6px 10px", cursor: "pointer", fontSize: 15, fontFamily: "inherit" }}>
            {dark ? "☀️" : "🌙"}
          </button>
          <button onClick={onReset} style={{ background: th.btnBg, border: `1px solid ${th.border}`, borderRadius: 8, padding: "6px 14px", color: th.textMuted, fontSize: 12, cursor: "pointer", fontFamily: "inherit" }}>
            ↑ New file
          </button>
        </div>
      </header>

      <main style={{ maxWidth: 1320, margin: "0 auto", padding: "24px 20px" }}>

        {/* ── Stat cards ── */}
        <div style={{ display: "flex", gap: 12, flexWrap: "wrap", marginBottom: 24 }}>
          <StatCard
            label="Total Spent"
            value={fmt(stats.totalSpent)}
            sub={`sum of ${transactions.filter(t => t.debit > 0).length} debit rows`}
            accent={th.debit} th={th}
          />
          <StatCard
            label="Total Received"
            value={fmt(stats.totalReceived)}
            sub={`sum of ${transactions.filter(t => t.credit > 0).length} credit rows`}
            accent={th.credit} th={th}
          />
          <StatCard
            label="Net Position"
            value={fmt(stats.net)}
            sub={stats.net >= 0 ? "net surplus" : "net deficit"}
            accent={stats.net >= 0 ? th.credit : th.debit} th={th}
          />
          <StatCard
            label="Categories"
            value={stats.categories.length}
            sub={`across ${transactions.length} transactions`}
            accent="#8b5cf6" th={th}
          />
        </div>

        {/* ── OVERVIEW TAB ── */}
        {activeTab === "overview" && (
          <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>

            {/* Row 1: area chart (wide) + donut (narrow) */}
            <div style={{ display: "grid", gridTemplateColumns: "2fr 1fr", gap: 16, minWidth: 0 }}>
              <div style={{ background: th.bgCard, border: `1px solid ${th.border}`, borderRadius: 16, padding: "20px 20px 12px", boxShadow: "0 1px 4px rgba(0,0,0,0.05)", minWidth: 0 }}>
                <div style={{ fontWeight: 700, fontSize: 15, color: th.text }}>Monthly Cash Flow</div>
                <div style={{ fontSize: 12, color: th.textMuted, marginTop: 2, marginBottom: 12 }}>Debit vs credit by month</div>
                <ResponsiveContainer width="100%" height={240}>
                  <AreaChart data={stats.monthly} margin={{ top: 10, right: 8, left: 0, bottom: 0 }}>
                    <defs>
                      <linearGradient id="gD" x1="0" y1="0" x2="0" y2="1">
                        <stop offset="5%" stopColor={th.debit} stopOpacity={0.25} />
                        <stop offset="95%" stopColor={th.debit} stopOpacity={0} />
                      </linearGradient>
                      <linearGradient id="gC" x1="0" y1="0" x2="0" y2="1">
                        <stop offset="5%" stopColor={th.credit} stopOpacity={0.25} />
                        <stop offset="95%" stopColor={th.credit} stopOpacity={0} />
                      </linearGradient>
                    </defs>
                    <CartesianGrid strokeDasharray="3 3" stroke={th.chartGrid} />
                    <XAxis dataKey="label" tick={{ fill: th.chartTick, fontSize: 11 }} tickLine={false} axisLine={false} />
                    <YAxis tick={{ fill: th.chartTick, fontSize: 11 }} tickLine={false} axisLine={false} tickFormatter={v => `$${(v / 1000).toFixed(0)}k`} width={48} />
                    <Tooltip content={<ChartTooltip {...ctProps} />} />
                    <Legend wrapperStyle={{ fontSize: 11, color: th.textMuted, paddingTop: 8 }} />
                    <Area type="monotone" dataKey="debit"  name="Spent"    stroke={th.debit}  fill="url(#gD)" strokeWidth={2} dot={false} />
                    <Area type="monotone" dataKey="credit" name="Received" stroke={th.credit} fill="url(#gC)" strokeWidth={2} dot={false} />
                  </AreaChart>
                </ResponsiveContainer>
              </div>

              <div style={{ background: th.bgCard, border: `1px solid ${th.border}`, borderRadius: 16, padding: "20px 16px 12px", boxShadow: "0 1px 4px rgba(0,0,0,0.05)", minWidth: 0, display: "flex", flexDirection: "column" }}>
                <div style={{ fontWeight: 700, fontSize: 15, color: th.text }}>Spending Mix</div>
                <div style={{ fontSize: 12, color: th.textMuted, marginTop: 2, marginBottom: 8 }}>Click a slice to drill in</div>
                <ResponsiveContainer width="100%" height={200}>
                  <PieChart>
                    <Pie data={stats.categories.filter(c => c.debit > 0)} dataKey="debit" nameKey="label"
                      cx="50%" cy="50%" outerRadius={80} innerRadius={46} paddingAngle={2}
                      onClick={d => { setActiveCategory(d.label); setActiveTab("transactions"); }} style={{ cursor: "pointer" }}>
                      {stats.categories.filter(c => c.debit > 0).map((_, i) => (
                        <Cell key={i} fill={PALETTE[i % PALETTE.length]} opacity={0.9} />
                      ))}
                    </Pie>
                    <Tooltip formatter={v => fmt(v)} contentStyle={{ background: th.tooltipBg, border: `1px solid ${th.border}`, borderRadius: 8, fontSize: 12, color: th.text }} />
                  </PieChart>
                </ResponsiveContainer>
                {/* Inline legend */}
                <div style={{ display: "flex", flexDirection: "column", gap: 5, marginTop: 4 }}>
                  {stats.categories.filter(c => c.debit > 0).slice(0, 6).map((cat, i) => (
                    <div key={cat.label} onClick={() => { setActiveCategory(cat.label); setActiveTab("transactions"); }}
                      style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 11, cursor: "pointer" }}>
                      <span style={{ width: 8, height: 8, borderRadius: 2, background: PALETTE[i % PALETTE.length], flexShrink: 0 }} />
                      <span style={{ color: th.textMuted, flex: 1, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{cat.label}</span>
                      <span style={{ color: th.text, fontFamily: "'DM Mono', monospace" }}>{fmt(cat.debit)}</span>
                    </div>
                  ))}
                  {stats.categories.filter(c => c.debit > 0).length > 6 && (
                    <div style={{ fontSize: 11, color: th.textMuted, paddingLeft: 14 }}>
                      +{stats.categories.filter(c => c.debit > 0).length - 6} more categories
                    </div>
                  )}
                </div>
              </div>
            </div>

            {/* Row 2: horizontal bar — full width */}
            <div style={{ background: th.bgCard, border: `1px solid ${th.border}`, borderRadius: 16, padding: "20px 20px 12px", boxShadow: "0 1px 4px rgba(0,0,0,0.05)" }}>
              <div style={{ fontWeight: 700, fontSize: 15, color: th.text }}>Top Spending Categories</div>
              <div style={{ fontSize: 12, color: th.textMuted, marginTop: 2, marginBottom: 12 }}>Total debit amount per category</div>
              <ResponsiveContainer width="100%" height={Math.max(180, Math.min(stats.categories.filter(c => c.debit > 0).length, 10) * 36)}>
                <BarChart data={stats.categories.filter(c => c.debit > 0).slice(0, 10)} layout="vertical" margin={{ top: 4, right: 60, left: 8, bottom: 4 }}>
                  <CartesianGrid strokeDasharray="3 3" stroke={th.chartGrid} horizontal={false} />
                  <XAxis type="number" tick={{ fill: th.chartTick, fontSize: 11 }} tickLine={false} axisLine={false} tickFormatter={v => `$${(v / 1000).toFixed(0)}k`} />
                  <YAxis type="category" dataKey="label" tick={{ fill: th.text, fontSize: 12 }} tickLine={false} axisLine={false} width={120} />
                  <Tooltip content={<ChartTooltip {...ctProps} />} />
                  <Bar dataKey="debit" name="Spent" radius={[0, 6, 6, 0]}>
                    {stats.categories.filter(c => c.debit > 0).slice(0, 10).map((_, i) => (
                      <Cell key={i} fill={PALETTE[i % PALETTE.length]} />
                    ))}
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            </div>

            {/* Category summary table */}
            <div style={{ background: th.bgCard, border: `1px solid ${th.border}`, borderRadius: 16, padding: 20, boxShadow: "0 1px 4px rgba(0,0,0,0.05)" }}>
              <div style={{ fontWeight: 700, fontSize: 15, color: th.text, marginBottom: 16 }}>Category Summary</div>
              <DataTable data={stats.categories} columns={catColumns} th={th} />
            </div>
          </div>
        )}

        {/* ── TRANSACTIONS TAB ── */}
        {activeTab === "transactions" && (
          <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
            <div style={{ display: "flex", gap: 6, flexWrap: "wrap", alignItems: "center" }}>
              <span style={{ fontSize: 12, color: th.textMuted, marginRight: 2 }}>Filter:</span>
              <button onClick={() => setActiveCategory(null)}
                style={{ background: !activeCategory ? th.accentDim : th.btnBg, color: !activeCategory ? th.accent : th.textMuted, border: !activeCategory ? `1px solid ${th.accentBorder}` : `1px solid ${th.border}`, borderRadius: 100, padding: "4px 14px", fontSize: 12, fontWeight: 600, cursor: "pointer", fontFamily: "inherit" }}>
                All
              </button>
              {stats.categories.map((cat, i) => (
                <button key={cat.label} onClick={() => setActiveCategory(cat.label === activeCategory ? null : cat.label)}
                  style={{
                    background: activeCategory === cat.label ? `${PALETTE[i % PALETTE.length]}22` : th.btnBg,
                    color: activeCategory === cat.label ? PALETTE[i % PALETTE.length] : th.textMuted,
                    border: activeCategory === cat.label ? `1px solid ${PALETTE[i % PALETTE.length]}55` : `1px solid ${th.border}`,
                    borderRadius: 100, padding: "4px 14px", fontSize: 12, fontWeight: 600,
                    cursor: "pointer", fontFamily: "inherit",
                  }}>{cat.label}
                </button>
              ))}
            </div>

            {activeCategory && (
              <div style={{ background: th.accentDim, border: `1px solid ${th.accentBorder}`, borderRadius: 12, padding: "12px 16px", display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8 }}>
                <span style={{ fontWeight: 700, color: th.accent }}>{activeCategory}</span>
                <span style={{ color: th.textMuted, fontSize: 13 }}>
                  {filteredTx.length} transactions · Spent {fmt(filteredTx.reduce((s, t) => s + t.debit, 0))} · Received {fmt(filteredTx.reduce((s, t) => s + t.credit, 0))}
                </span>
              </div>
            )}

            <div style={{ background: th.bgCard, border: `1px solid ${th.border}`, borderRadius: 16, padding: 20, boxShadow: "0 1px 4px rgba(0,0,0,0.05)" }}>
              <DataTable data={filteredTx} columns={txColumns} th={th} />
            </div>
          </div>
        )}
      </main>

      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=DM+Sans:ital,opsz,wght@0,9..40,400;0,9..40,600;0,9..40,700;0,9..40,800&family=DM+Mono:wght@400;500&display=swap');
        * { box-sizing: border-box; }
        ::-webkit-scrollbar { width: 6px; height: 6px; }
        ::-webkit-scrollbar-track { background: transparent; }
        ::-webkit-scrollbar-thumb { background: ${th.scrollThumb}; border-radius: 3px; }
        input, select, button { font-family: 'DM Sans', sans-serif; }
      `}</style>
    </div>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// ROOT
// ─────────────────────────────────────────────────────────────────────────────
function App() {
  const qc = useQueryClient();
  const [transactions, setTransactions] = useState(null);
  const [dark, setDark] = useState(true);

  const handleData = useCallback((data) => {
    qc.setQueryData(["transactions"], data);
    setTransactions(data);
  }, [qc]);

  const handleReset = useCallback(() => {
    qc.removeQueries({ queryKey: ["transactions"] });
    setTransactions(null);
  }, [qc]);

  if (!transactions) return <UploadScreen onData={handleData} dark={dark} toggleDark={() => setDark(d => !d)} />;
  return <Dashboard transactions={transactions} onReset={handleReset} dark={dark} toggleDark={() => setDark(d => !d)} />;
}

export default function Root() {
  return (
    <QueryClientProvider client={queryClient}>
      <App />
    </QueryClientProvider>
  );
}
