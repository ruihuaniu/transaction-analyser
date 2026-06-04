import React, {
  useState,
  useMemo,
  useCallback,
  useEffect,
} from "react";
import {
  useReactTable,
  getCoreRowModel,
  getSortedRowModel,
  getFilteredRowModel,
  getPaginationRowModel,
  flexRender,
  createColumnHelper,
} from "@tanstack/react-table";
import {
  useQueryClient,
  QueryClient,
  QueryClientProvider,
} from "@tanstack/react-query";
import * as XLSX from "xlsx";
import Papa from "papaparse";
import {
  AreaChart,
  Area,
  BarChart,
  Bar,
  LineChart,
  Line,
  PieChart,
  Pie,
  Cell,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  ResponsiveContainer,
  Legend,
  ReferenceLine,
} from "recharts";

const queryClient = new QueryClient();

// ─────────────────────────────────────────────────────────────────────────────
// COLUMN FINDER
// Strategy: exact match first (case-insensitive, trimmed), then "starts with",
// then "includes". This prevents "Debit Amount" accidentally matching a
// "Credit Amount" column when both are present.
// ─────────────────────────────────────────────────────────────────────────────
function findColumn(keys, candidates) {
  const norm = (s) => s.toLowerCase().trim();
  // 1. exact match
  for (const c of candidates) {
    const k = keys.find((k) => norm(k) === norm(c));
    if (k) return k;
  }
  // 2. starts-with match
  for (const c of candidates) {
    const k = keys.find((k) => norm(k).startsWith(norm(c)));
    if (k) return k;
  }
  // 3. includes match (last resort)
  for (const c of candidates) {
    const k = keys.find((k) => norm(k).includes(norm(c)));
    if (k) return k;
  }
  return null;
}

// ─────────────────────────────────────────────────────────────────────────────
// AMOUNT PARSER
// Bank CSVs: "Debit Amount" and "Credit Amount" columns normally hold
// positive values. A blank cell means $0 for that direction.
// Some banks export negative numbers in the debit column for reversals —
// preserve the sign so totals equal the raw column sums.
// ─────────────────────────────────────────────────────────────────────────────
function parseAmount(val) {
  if (val === undefined || val === null || val === "") return 0;
  const n = parseFloat(String(val).replace(/[$,\s]/g, ""));
  return isNaN(n) ? 0 : n;
}

function normalizeMerchantLabel(value) {
  const label = String(value ?? "").trim();
  return label ? label.replace(/\s+/g, " ") : "Unknown merchant";
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
  const dateKey = findColumn(keys, ["date"]);
  const narrativeKey = findColumn(keys, [
    "narrative",
    "description",
    "details",
    "merchant",
    "narr",
  ]);
  // IMPORTANT: look for the more-specific "debit amount" BEFORE plain "debit"
  // and "credit amount" BEFORE plain "credit" to avoid cross-matching.
  const debitKey = findColumn(keys, [
    "debit amount",
    "debit amt",
    "debit",
    "dr amount",
    "dr",
  ]);
  const creditKey = findColumn(keys, [
    "credit amount",
    "credit amt",
    "credit",
    "cr amount",
    "cr",
  ]);

  // Sanity: if debitKey === creditKey something went wrong — fall back to null
  const safeDebitKey = debitKey !== creditKey ? debitKey : null;
  const safeCreditKey = debitKey !== creditKey ? creditKey : null;

  return rawRows
    .map((row, idx) => {
      const dateRaw = dateKey ? row[dateKey] : "";
      const narrative = narrativeKey ? String(row[narrativeKey] ?? "") : "";
      const debit = parseAmount(safeDebitKey ? row[safeDebitKey] : "");
      const credit = parseAmount(safeCreditKey ? row[safeCreditKey] : "");

      const dateObj = parseDateStr(dateRaw);
      return {
        id: idx,
        date: dateObj,
        dateStr: dateObj
          ? dateObj.toLocaleDateString("en-AU")
          : String(dateRaw ?? ""),
        dateSort: dateObj ? dateObj.getTime() : 0,
        narrative,
        debit, // ← "Debit Amount" column value
        credit, // ← "Credit Amount" column value
        category: categorize(narrative),
      };
    })
    .filter((r) => r.narrative || r.debit !== 0 || r.credit !== 0);
}

function parseDateStr(val) {
  if (!val) return null;
  if (typeof val === "number") {
    const d = XLSX.SSF.parse_date_code(val);
    if (d) return new Date(d.y, d.m - 1, d.d);
  }
  const str = String(val).trim();
  // Support YYYYMMDD strings that appear in some CSV exports.
  const ymdMatch = /^([0-9]{4})([0-9]{2})([0-9]{2})$/u.exec(str);
  if (ymdMatch) {
    return new Date(
      Number(ymdMatch[1]),
      Number(ymdMatch[2]) - 1,
      Number(ymdMatch[3]),
    );
  }

  // Support AU-style day/month/year dates like 01/04/2025.
  const dmyMatch = /^([0-9]{1,2})[\/\-]([0-9]{1,2})[\/\-]([0-9]{2,4})$/u.exec(
    str,
  );
  if (dmyMatch) {
    const day = Number(dmyMatch[1]);
    const month = Number(dmyMatch[2]);
    let year = Number(dmyMatch[3]);
    if (year < 100) year += 2000;
    return new Date(year, month - 1, day);
  }

  const d = new Date(str);
  return isNaN(d.getTime()) ? null : d;
}

function fmt(n) {
  return new Intl.NumberFormat("en-AU", {
    style: "currency",
    currency: "AUD",
    minimumFractionDigits: 2,
  }).format(n ?? 0);
}

const STORAGE_KEY = "txanalyser.transactions";
const THEME_KEY = "txanalyser.dark";

function parseStoredTransactions(value) {
  if (!value) return null;
  try {
    const stored = JSON.parse(value);
    if (!Array.isArray(stored)) return null;
    return stored.map((item) => ({
      ...item,
      date: item?.date ? new Date(item.date) : null,
    }));
  } catch {
    return null;
  }
}

function getStoredTransactions() {
  if (typeof window === "undefined") return null;
  return parseStoredTransactions(localStorage.getItem(STORAGE_KEY));
}

function saveTransactions(transactions) {
  if (typeof window === "undefined") return;
  if (!transactions || !transactions.length) {
    localStorage.removeItem(STORAGE_KEY);
    return;
  }
  localStorage.setItem(
    STORAGE_KEY,
    JSON.stringify(
      transactions.map((t) => ({
        ...t,
        date: t.date ? t.date.toISOString() : null,
      })),
    ),
  );
}

function getStoredDarkMode() {
  if (typeof window === "undefined") return false;
  return localStorage.getItem(THEME_KEY) === "true";
}

function saveDarkMode(dark) {
  if (typeof window === "undefined") return;
  localStorage.setItem(THEME_KEY, dark ? "true" : "false");
}

// ─────────────────────────────────────────────────────────────────────────────
// CATEGORY RULES — order matters: more specific rules first
// ─────────────────────────────────────────────────────────────────────────────
const CATEGORY_RULES = [
  // Income
  {
    keywords: [
      "salary",
      "payroll",
      "wage",
      "income",
      "pay slip",
      "direct credit",
      "employer",
    ],
    label: "Income",
  },
  // Housing
  {
    keywords: [
      "rent",
      "mortgage",
      "landlord",
      "lease",
      "strata",
      "body corporate",
      "real estate",
      "property management",
    ],
    label: "Housing",
  },
  // Hardware & Home
  {
    keywords: [
      "bunnings",
      "mitre 10",
      "masters home",
      "hardware",
      "timber",
      "plumbing",
      "electrical supplies",
      "paint",
      "ikea",
      "freedom furniture",
      "harvey norman",
      "the good guys",
      "jb hi-fi",
      "bing lee",
      "officeworks",
      "pottery barn",
      "bed bath",
      "kmart",
      "target",
      "big w",
      "spotlight",
      "lincraft",
      "howards storage",
    ],
    label: "Hardware & Home",
  },
  // Groceries — before dining so "supermarket" wins over restaurants
  {
    keywords: [
      "woolworths",
      "woolworth",
      "coles",
      "aldi",
      "iga",
      "harris farm",
      "foodland",
      "spar",
      "drakes",
      "costco",
      "grocery",
      "supermarket",
      "fruit shop",
      "butcher",
      "bakery",
      "deli",
      "fresh market",
    ],
    label: "Groceries",
  },
  // Dining & takeaway
  {
    keywords: [
      "restaurant",
      "cafe",
      "coffee",
      "hungry jacks",
      "mcdonald",
      "kfc",
      "subway",
      "domino",
      "pizza",
      "burger",
      "sushi",
      "ramen",
      "thai",
      "chinese",
      "indian",
      "italian",
      "brunch",
      "bistro",
      "bar & grill",
      "eatery",
      "takeaway",
      "takeout",
      "uber eats",
      "doordash",
      "menulog",
      "deliveroo",
      "dining",
      "food truck",
      "noodle",
      "kebab",
      "shawarma",
      "taco",
      "tea",
      "bakery",
      "cuis",
      "hunan",
    ],
    label: "Dining & takeaway",
  },
  // Transport
  {
    keywords: [
      "uber",
      "lyft",
      "ola",
      "didi",
      "taxi",
      "transport nsw",
      "translink",
      "myki",
      "opal",
      "train",
      "bus",
      "tram",
      "metro",
      "ferry",
      "fuel",
      "petrol",
      "shell",
      "bp ",
      "caltex",
      "ampol",
      "parking",
      "car park",
      "toll",
      "linkt",
      "e-toll",
      "roam express",
    ],
    label: "Transport",
  },
  // Health & Fitness
  {
    keywords: [
      "gym",
      "fitness",
      "crossfit",
      "yoga",
      "pilates",
      "f45",
      "anytime fitness",
      "planet fitness",
      "snap fitness",
      "doctor",
      "gp ",
      "medical centre",
      "hospital",
      "pharmacy",
      "chemist",
      "priceline",
      "terry white",
      "healthdirect",
      "dentist",
      "optical",
      "physio",
      "pathology",
    ],
    label: "Health & Fitness",
  },
  // Subscriptions
  {
    keywords: [
      "netflix",
      "stan",
      "binge",
      "paramount",
      "disney",
      "apple tv",
      "prime video",
      "spotify",
      "apple music",
      "youtube premium",
      "audible",
      "kindle",
      "adobe",
      "microsoft 365",
      "google one",
      "icloud",
      "dropbox",
      "canva subscription",
      "hbo",
    ],
    label: "Subscriptions",
  },
  // Utilities
  {
    keywords: [
      "electricity",
      "energy australia",
      "agl ",
      "origin energy",
      "simply energy",
      "powershop",
      "water ",
      "sydney water",
      "yarra valley water",
      "gas ",
      "ausnet",
      "jemena",
      "internet",
      "nbn",
      "telstra",
      "optus",
      "vodafone",
      "tpg",
      "aussie broadband",
      "aussie bb",
      "mobile plan",
      "phone bill",
    ],
    label: "Utilities",
  },
  // Travel
  {
    keywords: [
      "qantas",
      "virgin australia",
      "jetstar",
      "rex airline",
      "air new zealand",
      "singapore air",
      "emirates",
      "flight",
      "airfare",
      "airbnb",
      "booking.com",
      "hotels.com",
      "expedia",
      "wotif",
      "agoda",
      "hotel",
      "motel",
      "resort",
      "hostel",
      "holiday",
      "vacation",
      "travel insurance",
      "visa fee",
    ],
    label: "Travel",
  },
  // Education
  {
    keywords: [
      "university",
      "tafe",
      "school fee",
      "tuition",
      "udemy",
      "coursera",
      "skillshare",
      "linkedin learning",
      "masterclass",
      "textbook",
      "stationery",
    ],
    label: "Education",
  },
  // Entertainment
  {
    keywords: [
      "cinema",
      "event cinema",
      "hoyts",
      "village cinema",
      "reading cinema",
      "concert",
      "ticketek",
      "ticketmaster",
      "museum",
      "gallery",
      "zoo",
      "theme park",
      "entertainment",
    ],
    label: "Entertainment",
  },
  // Shopping (catch-all online/fashion after hardware & home already handled)
  {
    keywords: [
      "amazon",
      "ebay",
      "etsy",
      "aliexpress",
      "shein",
      "asos",
      "zara",
      "h&m",
      "uniqlo",
      "cotton on",
      "factorie",
      "rivers",
      "lowes",
      "tarocash",
      "myer",
      "david jones",
      "the iconic",
      "net-a-porter",
      "shop",
      "store",
      "clothing",
      "fashion",
      "jewellery",
    ],
    label: "Shopping",
  },
  // Investments & Finance
  {
    keywords: [
      "commsec",
      "nabtrade",
      "selfwealth",
      "stake",
      "raiz",
      "spaceship",
      "vanguard",
      "brokerage",
      "share purchase",
      "dividends received",
      "super",
      "superannuation",
      "insurance",
      "life insurance",
      "car insurance",
      "home insurance",
      "income protection",
    ],
    label: "Insurance & Finance",
  },
  // Banking & Fees
  {
    keywords: [
      "bank fee",
      "account fee",
      "monthly fee",
      "transaction fee",
      "dishonour",
      "overdrawn",
      "atm fee",
      "interest charge",
      "late payment",
      "annual fee",
      "bpay",
      "eft",
      "wire transfer",
    ],
    label: "Banking & Fees",
  },
  // Government & Tax
  {
    keywords: [
      "ato ",
      "tax office",
      "centrelink",
      "medicare",
      "services australia",
      "council rates",
      "land tax",
      "stamp duty",
      "fine ",
      "toll infringement",
      "government",
      "dept of",
    ],
    label: "Government & Tax",
  },
  // Charity
  {
    keywords: [
      "donate",
      "donation",
      "charity",
      "foundation",
      "red cross",
      "cancer council",
      "oxfam",
      "beyond blue",
      "lifeline",
      "salvos",
    ],
    label: "Charity",
  },
];

function categorize(narrative) {
  if (!narrative) return "Other";
  const lower = narrative.toLowerCase();
  for (const rule of CATEGORY_RULES) {
    if (rule.keywords.some((k) => lower.includes(k))) return rule.label;
  }
  return "Other";
}

// ─────────────────────────────────────────────────────────────────────────────
// PALETTE & HELPERS
// ─────────────────────────────────────────────────────────────────────────────
const PALETTE = [
  "#10b981",
  "#3b82f6",
  "#f43f5e",
  "#f59e0b",
  "#8b5cf6",
  "#ec4899",
  "#06b6d4",
  "#84cc16",
  "#f97316",
  "#d946ef",
  "#14b8a6",
  "#6366f1",
  "#a78bfa",
  "#fb923c",
  "#34d399",
  "#60a5fa",
  "#f472b6",
  "#a3e635",
];

const columnHelper = createColumnHelper();

// ─────────────────────────────────────────────────────────────────────────────
// THEME
// ─────────────────────────────────────────────────────────────────────────────
function makeTheme(dark) {
  return dark
    ? {
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
      }
    : {
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
    <div
      style={{
        background: th.bgCard,
        border: `1px solid ${th.border}`,
        borderRadius: 16,
        padding: "20px 24px",
        display: "flex",
        flexDirection: "column",
        gap: 4,
        flex: "1 1 180px",
        minWidth: 160,
        position: "relative",
        overflow: "hidden",
        boxShadow: "0 1px 4px rgba(0,0,0,0.06)",
      }}
    >
      <div
        style={{
          position: "absolute",
          top: 0,
          left: 0,
          right: 0,
          height: 3,
          background: accent,
          borderRadius: "16px 16px 0 0",
        }}
      />
      <span
        style={{
          fontSize: 11,
          fontWeight: 600,
          letterSpacing: "0.12em",
          color: th.textMuted,
          textTransform: "uppercase",
        }}
      >
        {label}
      </span>
      <span
        style={{
          fontSize: 22,
          fontWeight: 700,
          color: th.text,
          fontFamily: "'DM Mono', monospace",
          marginTop: 4,
        }}
      >
        {value}
      </span>
      {sub && <span style={{ fontSize: 12, color: th.textMuted }}>{sub}</span>}
    </div>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// SEARCH INPUT
// ─────────────────────────────────────────────────────────────────────────────
function SearchInput({ value, onChange, placeholder, th }) {
  return (
    <div
      style={{
        position: "relative",
        display: "inline-flex",
        alignItems: "center",
      }}
    >
      <svg
        width='14'
        height='14'
        viewBox='0 0 24 24'
        fill='none'
        stroke={th.textMuted}
        strokeWidth='2.5'
        style={{ position: "absolute", left: 10, pointerEvents: "none" }}
      >
        <circle cx='11' cy='11' r='8' />
        <path d='m21 21-4.35-4.35' />
      </svg>
      <input
        value={value}
        onChange={(e) => onChange(e.target.value)}
        placeholder={placeholder || "Search…"}
        style={{
          background: th.inputBg,
          border: `1px solid ${th.border}`,
          borderRadius: 8,
          padding: "7px 12px 7px 30px",
          color: th.text,
          fontSize: 13,
          outline: "none",
          width: 220,
          fontFamily: "inherit",
        }}
      />
    </div>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// DATA TABLE (TanStack Table)
// ─────────────────────────────────────────────────────────────────────────────
function DataTable({ data, columns, th, rowProps }) {
  const [sorting, setSorting] = useState([]);
  const [globalFilter, setGlobalFilter] = useState("");
  const [pagination, setPagination] = useState({ pageIndex: 0, pageSize: 15 });

  const table = useReactTable({
    data,
    columns,
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
      <div
        style={{
          display: "flex",
          justifyContent: "space-between",
          alignItems: "center",
          flexWrap: "wrap",
          gap: 8,
        }}
      >
        <SearchInput
          value={globalFilter}
          onChange={(v) => {
            setGlobalFilter(v);
            setPagination((p) => ({ ...p, pageIndex: 0 }));
          }}
          placeholder='Search all columns…'
          th={th}
        />
        <span style={{ fontSize: 12, color: th.textMuted }}>
          {table.getFilteredRowModel().rows.length.toLocaleString()} rows
        </span>
      </div>
      <div
        style={{
          overflowX: "auto",
          borderRadius: 12,
          border: `1px solid ${th.border}`,
        }}
      >
        <table
          style={{
            width: "100%",
            borderCollapse: "collapse",
            fontSize: 13,
            minWidth: 500,
          }}
        >
          <thead>
            {table.getHeaderGroups().map((hg) => (
              <tr
                key={hg.id}
                style={{
                  borderBottom: `1px solid ${th.borderStrong}`,
                  background: th.tableHead,
                }}
              >
                {hg.headers.map((header) => (
                  <th
                    key={header.id}
                    onClick={header.column.getToggleSortingHandler()}
                    style={{
                      padding: "10px 14px",
                      textAlign: "left",
                      fontWeight: 600,
                      fontSize: 11,
                      letterSpacing: "0.08em",
                      textTransform: "uppercase",
                      color: th.textMuted,
                      cursor: header.column.getCanSort()
                        ? "pointer"
                        : "default",
                      userSelect: "none",
                      whiteSpace: "nowrap",
                    }}
                  >
                    <div
                      style={{ display: "flex", alignItems: "center", gap: 4 }}
                    >
                      {flexRender(
                        header.column.columnDef.header,
                        header.getContext(),
                      )}
                      {header.column.getIsSorted() === "asc" && " ↑"}
                      {header.column.getIsSorted() === "desc" && " ↓"}
                      {!header.column.getIsSorted() &&
                        header.column.getCanSort() && (
                          <span style={{ opacity: 0.3 }}> ⇅</span>
                        )}
                    </div>
                  </th>
                ))}
              </tr>
            ))}
          </thead>
          <tbody>
            {rows.length === 0 ? (
              <tr>
                <td
                  colSpan={columns.length}
                  style={{
                    padding: "40px",
                    textAlign: "center",
                    color: th.textMuted,
                  }}
                >
                  No results
                </td>
              </tr>
            ) : (
              rows.map((row, i) => (
                <tr
                  key={row.id}
                  {...rowProps?.(row)}
                  style={{
                    borderBottom: `1px solid ${th.border}`,
                    background: i % 2 === 0 ? "transparent" : th.tableRow,
                    transition: "background 0.1s",
                  }}
                  onMouseEnter={(e) => {
                    e.currentTarget.style.background = th.tableRowHover;
                  }}
                  onMouseLeave={(e) => {
                    e.currentTarget.style.background =
                      i % 2 === 0 ? "transparent" : th.tableRow;
                  }}
                >
                  {row.getVisibleCells().map((cell) => (
                    <td
                      key={cell.id}
                      style={{ padding: "9px 14px", color: th.text }}
                    >
                      {flexRender(
                        cell.column.columnDef.cell,
                        cell.getContext(),
                      )}
                    </td>
                  ))}
                </tr>
              ))
            )}
          </tbody>
        </table>
      </div>
      <div
        style={{
          display: "flex",
          alignItems: "center",
          gap: 8,
          flexWrap: "wrap",
          justifyContent: "space-between",
        }}
      >
        <div style={{ display: "flex", gap: 4 }}>
          {[
            {
              label: "«",
              fn: () => table.setPageIndex(0),
              disabled: !table.getCanPreviousPage(),
            },
            {
              label: "‹",
              fn: () => table.previousPage(),
              disabled: !table.getCanPreviousPage(),
            },
            {
              label: "›",
              fn: () => table.nextPage(),
              disabled: !table.getCanNextPage(),
            },
            {
              label: "»",
              fn: () => table.setPageIndex(table.getPageCount() - 1),
              disabled: !table.getCanNextPage(),
            },
          ].map((btn, idx) => (
            <button
              key={idx}
              onClick={btn.fn}
              disabled={btn.disabled}
              style={{
                background: th.btnBg,
                border: `1px solid ${th.border}`,
                borderRadius: 6,
                padding: "4px 10px",
                color: btn.disabled ? th.btnDisabled : th.btnColor,
                cursor: btn.disabled ? "default" : "pointer",
                fontSize: 14,
                fontFamily: "inherit",
              }}
            >
              {btn.label}
            </button>
          ))}
        </div>
        <span style={{ fontSize: 12, color: th.textMuted }}>
          Page {table.getState().pagination.pageIndex + 1} /{" "}
          {Math.max(1, table.getPageCount())}
        </span>
        <select
          value={table.getState().pagination.pageSize}
          onChange={(e) => table.setPageSize(Number(e.target.value))}
          style={{
            background: th.inputBg,
            border: `1px solid ${th.border}`,
            borderRadius: 6,
            padding: "4px 8px",
            color: th.text,
            fontSize: 12,
            fontFamily: "inherit",
          }}
        >
          {[10, 15, 25, 50, 100].map((s) => (
            <option key={s} value={s}>
              {s} / page
            </option>
          ))}
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
    <div
      style={{
        background: th.tooltipBg,
        border: `1px solid ${th.border}`,
        borderRadius: 10,
        padding: "10px 14px",
        fontSize: 12,
        color: th.text,
        boxShadow: "0 4px 16px rgba(0,0,0,0.12)",
      }}
    >
      <div style={{ fontWeight: 700, marginBottom: 4, color: th.accent }}>
        {label}
      </div>
      {payload.map((p, i) => (
        <div
          key={i}
          style={{ display: "flex", gap: 12, justifyContent: "space-between" }}
        >
          <span style={{ color: p.color }}>{p.name}</span>
          <span style={{ fontFamily: "'DM Mono', monospace" }}>
            {fmt(p.value)}
          </span>
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
  const th = makeTheme(dark);

  const processFile = useCallback(
    (file) => {
      if (!file) return;
      setError(null);
      setLoading(true);
      const ext = file.name.split(".").pop().toLowerCase();

      if (ext === "csv") {
        Papa.parse(file, {
          header: true,
          skipEmptyLines: true,
          complete: (result) => {
            setLoading(false);
            if (!result.data.length) {
              setError("CSV appears empty.");
              return;
            }
            onData(normalizeRows(result.data));
          },
          error: () => {
            setLoading(false);
            setError("Failed to parse CSV.");
          },
        });
      } else if (["xlsx", "xls"].includes(ext)) {
        const reader = new FileReader();
        reader.onload = (e) => {
          try {
            const wb = XLSX.read(e.target.result, {
              type: "binary",
              cellDates: false,
            });
            const ws = wb.Sheets[wb.SheetNames[0]];
            const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });
            setLoading(false);
            if (!rows.length) {
              setError("Excel sheet appears empty.");
              return;
            }
            onData(normalizeRows(rows));
          } catch {
            setLoading(false);
            setError("Failed to parse Excel file.");
          }
        };
        reader.readAsBinaryString(file);
      } else {
        setLoading(false);
        setError("Please upload a .csv, .xlsx, or .xls file.");
      }
    },
    [onData],
  );

  return (
    <div
      style={{
        minHeight: "100vh",
        background: th.bg,
        display: "flex",
        flexDirection: "column",
        alignItems: "center",
        justifyContent: "center",
        fontFamily: "'DM Sans', sans-serif",
        padding: 24,
        transition: "background 0.3s",
      }}
    >
      <div
        style={{
          position: "fixed",
          top: 0,
          left: 0,
          right: 0,
          height: 400,
          background: `radial-gradient(ellipse at 50% -10%, ${th.glow} 0%, transparent 70%)`,
          pointerEvents: "none",
        }}
      />
      <button
        onClick={toggleDark}
        title='Toggle light/dark'
        style={{
          position: "fixed",
          top: 20,
          right: 24,
          background: th.bgCard,
          border: `1px solid ${th.border}`,
          borderRadius: 10,
          padding: "7px 13px",
          cursor: "pointer",
          fontSize: 16,
          fontFamily: "inherit",
        }}
      >
        {dark ? "☀️" : "🌙"}
      </button>
      <div style={{ textAlign: "center", marginBottom: 48 }}>
        <div
          style={{
            display: "inline-flex",
            alignItems: "center",
            gap: 8,
            background: th.accentDim,
            border: `1px solid ${th.accentBorder}`,
            borderRadius: 100,
            padding: "4px 14px",
            marginBottom: 20,
            fontSize: 12,
            fontWeight: 600,
            letterSpacing: "0.1em",
            color: th.accent,
            textTransform: "uppercase",
          }}
        >
          ◆ Transaction Analyser
        </div>
        <h1
          style={{
            fontSize: "clamp(32px,5vw,52px)",
            fontWeight: 800,
            color: th.text,
            margin: 0,
            lineHeight: 1.1,
            letterSpacing: "-0.02em",
          }}
        >
          Understand your
          <br />
          <span style={{ color: th.accent }}>financial story</span>
        </h1>
        <p
          style={{
            color: th.textMuted,
            marginTop: 16,
            fontSize: 15,
            maxWidth: 400,
            lineHeight: 1.6,
          }}
        >
          Upload a bank statement to instantly visualise spending, categorise
          transactions, and spot patterns.
        </p>
      </div>

      <label
        onDragOver={(e) => {
          e.preventDefault();
          setDragging(true);
        }}
        onDragLeave={() => setDragging(false)}
        onDrop={(e) => {
          e.preventDefault();
          setDragging(false);
          processFile(e.dataTransfer.files[0]);
        }}
        style={{
          position: "relative",
          width: "100%",
          maxWidth: 500,
          border: `2px dashed ${dragging ? th.accent : th.borderStrong}`,
          borderRadius: 20,
          padding: "48px 32px",
          textAlign: "center",
          cursor: "pointer",
          background: dragging ? th.accentDim : th.bgCard,
          transition: "all 0.2s",
          boxShadow: "0 2px 12px rgba(0,0,0,0.06)",
          display: "block",
        }}
      >
        <input
          type='file'
          accept='.csv,text/csv,application/csv,.xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,.xls,application/vnd.ms-excel'
          style={{
            position: "absolute",
            inset: 0,
            width: "100%",
            height: "100%",
            opacity: 0,
            cursor: "pointer",
            pointerEvents: "auto",
            zIndex: 1,
          }}
          onChange={(e) => {
            const file = e.target.files?.[0];
            if (file) processFile(file);
            e.target.value = "";
          }}
        />
        {loading ? (
          <div style={{ color: th.accent, fontSize: 15 }}>
            <div
              style={{
                fontSize: 32,
                marginBottom: 12,
                display: "inline-block",
                animation: "spin 1s linear infinite",
              }}
            >
              ⟳
            </div>
            <br />
            Processing file…
          </div>
        ) : (
          <>
            <div style={{ fontSize: 40, marginBottom: 16 }}>📂</div>
            <div
              style={{
                color: th.text,
                fontWeight: 600,
                fontSize: 16,
                marginBottom: 6,
              }}
            >
              Drop your file here
            </div>
            <div style={{ color: th.textMuted, fontSize: 13 }}>
              or click to browse — CSV, XLSX, XLS
            </div>
          </>
        )}
      </label>
      {error && (
        <div
          style={{
            marginTop: 16,
            color: th.debit,
            fontSize: 13,
            background: `${th.debit}18`,
            border: `1px solid ${th.debit}40`,
            borderRadius: 8,
            padding: "8px 16px",
          }}
        >
          {error}
        </div>
      )}
      <div
        style={{
          marginTop: 28,
          fontSize: 12,
          color: th.textMuted,
          textAlign: "center",
        }}
      >
        Expected columns:{" "}
        <span style={{ color: th.accent }}>
          Date · Narrative · Debit Amount · Credit Amount
        </span>
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
  const [returnCategory, setReturnCategory] = useState(null);
  const th = makeTheme(dark);

  const getCategoryRowId = (label) =>
    `cat-summary-${label
      .replace(/[^a-z0-9]+/gi, "-")
      .replace(/(^-|-$)/g, "")
      .toLowerCase()}`;

  // ── STATS ─────────────────────────────────────────────────────────────────
  // totalSpent   = sum of every row's debit  field (= "Debit Amount" column)
  // totalReceived= sum of every row's credit field (= "Credit Amount" column)
  const stats = useMemo(() => {
    let totalSpent = 0;
    let totalReceived = 0;
    const catMap = {};
    const monthMap = {};

    const merchantMap = {};

    for (const t of transactions) {
      totalSpent += t.debit; // Debit Amount
      totalReceived += t.credit; // Credit Amount

      const merchant = normalizeMerchantLabel(t.narrative);
      if (!merchantMap[merchant])
        merchantMap[merchant] = {
          merchant,
          spent: 0,
          credit: 0,
          visits: 0,
          txCount: 0,
        };
      merchantMap[merchant].spent += t.debit;
      merchantMap[merchant].credit += t.credit;
      merchantMap[merchant].visits += 1;
      merchantMap[merchant].txCount += 1;

      // Category accumulation
      if (!catMap[t.category])
        catMap[t.category] = {
          label: t.category,
          debit: 0,
          credit: 0,
          count: 0,
        };
      catMap[t.category].debit += t.debit;
      catMap[t.category].credit += t.credit;
      catMap[t.category].count += 1;

      // Monthly accumulation
      if (t.date) {
        const key = `${t.date.getFullYear()}-${String(t.date.getMonth() + 1).padStart(2, "0")}`;
        const label = t.date.toLocaleDateString("en-AU", {
          month: "short",
          year: "2-digit",
        });
        if (!monthMap[key])
          monthMap[key] = { month: key, label, debit: 0, credit: 0 };
        monthMap[key].debit += t.debit;
        monthMap[key].credit += t.credit;
      }
    }

    const net = totalReceived - totalSpent;
    const categories = Object.values(catMap).sort((a, b) => b.debit - a.debit);
    const monthly = Object.values(monthMap).sort((a, b) =>
      a.month.localeCompare(b.month),
    );

    const merchants = Object.values(merchantMap)
      .map((m) => ({
        ...m,
        avg: m.visits ? m.spent / m.visits : 0,
      }))
      .sort((a, b) => b.spent - a.spent);
    const topMerchants = merchants.slice(0, 10);

    // ── Rolling 3-month average for trend line ────────────────────────────
    const monthlyWithAvg = monthly.map((m, i, arr) => {
      const window = arr.slice(Math.max(0, i - 2), i + 1);
      const avg = window.reduce((s, x) => s + x.debit, 0) / window.length;
      return { ...m, avg: Math.round(avg) };
    });

    // ── Key observations ──────────────────────────────────────────────────
    const observations = [];
    const spendCats = categories.filter(
      (c) => c.debit > 0 && c.label !== "Income",
    );
    const totalSpendCats = spendCats.reduce((s, c) => s + c.debit, 0);

    // 1. Top spending category
    if (spendCats[0]) {
      const pct =
        totalSpendCats > 0
          ? ((spendCats[0].debit / totalSpendCats) * 100).toFixed(0)
          : 0;
      observations.push({
        icon: "🏆",
        type: "info",
        title: `Biggest spend: ${spendCats[0].label}`,
        body: `${spendCats[0].label} accounts for ${pct}% of total spending at ${fmt(spendCats[0].debit)} — your single largest outgoing category.`,
      });
    }

    // 2. Monthly trend — is spending going up or down?
    if (monthly.length >= 3) {
      const recent = monthly.slice(-3).reduce((s, m) => s + m.debit, 0) / 3;
      const earlier =
        monthly
          .slice(0, Math.min(3, monthly.length - 3))
          .reduce((s, m) => s + m.debit, 0) / Math.min(3, monthly.length - 3);
      if (earlier > 0) {
        const changePct = (((recent - earlier) / earlier) * 100).toFixed(0);
        const rising = recent > earlier * 1.05;
        const falling = recent < earlier * 0.95;
        if (rising)
          observations.push({
            icon: "📈",
            type: "warning",
            title: "Spending is trending up",
            body: `Your average monthly spend has risen ~${changePct}% over the last 3 months (${fmt(recent)}/mo) vs the earlier period (${fmt(earlier)}/mo). Consider reviewing discretionary categories.`,
          });
        else if (falling)
          observations.push({
            icon: "📉",
            type: "success",
            title: "Spending is trending down",
            body: `Great progress! Your average monthly spend has dropped ~${Math.abs(changePct)}% to ${fmt(recent)}/mo vs ${fmt(earlier)}/mo earlier.`,
          });
        else
          observations.push({
            icon: "📊",
            type: "info",
            title: "Spending is relatively stable",
            body: `Monthly spending has stayed consistent at around ${fmt(recent)}/mo across the period analysed.`,
          });
      }
    }

    // 3. Highest single month
    if (monthly.length > 0) {
      const peak = [...monthly].sort((a, b) => b.debit - a.debit)[0];
      observations.push({
        icon: "📅",
        type: "info",
        title: `Peak month: ${peak.label}`,
        body: `${peak.label} had the highest withdrawals at ${fmt(peak.debit)}. Check if that month had one-off expenses like travel, bills, or large purchases.`,
      });
    }

    // 4. Savings rate
    if (totalReceived > 0) {
      const savingsRate = (
        ((totalReceived - totalSpent) / totalReceived) *
        100
      ).toFixed(0);
      const isPos = totalReceived >= totalSpent;
      observations.push({
        icon: isPos ? "💰" : "⚠️",
        type: isPos ? "success" : "warning",
        title: isPos
          ? `Savings rate: ${savingsRate}%`
          : "Spending exceeds income",
        body: isPos
          ? `You spent ${fmt(totalSpent)} against ${fmt(totalReceived)} received — a net surplus of ${fmt(totalReceived - totalSpent)}.`
          : `You spent ${fmt(Math.abs(totalSpent - totalReceived))} more than you received. Review recurring and discretionary costs.`,
      });
    }

    // 5. Dining out vs groceries ratio
    const dining = categories.find(
      (c) => c.label === "Dining & takeaway" || c.label === "Dining",
    );
    const groceries = categories.find((c) => c.label === "Groceries");
    if (dining && groceries && groceries.debit > 0) {
      const ratio = (dining.debit / groceries.debit).toFixed(1);
      if (ratio > 0.8)
        observations.push({
          icon: "🍽️",
          type: "warning",
          title: `Dining out is ${ratio}× your grocery spend`,
          body: `You spent ${fmt(dining.debit)} dining out vs ${fmt(groceries.debit)} on groceries. Cooking more at home could free up significant budget.`,
        });
    }

    // 6. Subscriptions flag
    const subs = categories.find((c) => c.label === "Subscriptions");
    if (subs && subs.debit > 0) {
      const perMonth =
        monthly.length > 0 ? subs.debit / monthly.length : subs.debit;
      observations.push({
        icon: "📱",
        type: "info",
        title: `Subscriptions: ${fmt(subs.debit)} total`,
        body: `Roughly ${fmt(perMonth)}/month on streaming and software subscriptions (${subs.count} transactions). Review for any unused services.`,
      });
    }

    return {
      totalSpent,
      totalReceived,
      net,
      categories,
      monthly: monthlyWithAvg,
      merchants,
      topMerchants,
      observations,
    };
  }, [transactions]);

  const filteredTx = useMemo(
    () =>
      activeCategory
        ? transactions.filter((t) => t.category === activeCategory)
        : transactions,
    [transactions, activeCategory],
  );

  // ── COLUMNS ───────────────────────────────────────────────────────────────
  const txColumns = useMemo(
    () => [
      columnHelper.accessor("dateStr", {
        header: "Date",
        sortingFn: (a, b) => a.original.dateSort - b.original.dateSort,
        cell: (info) => (
          <span
            style={{
              color: th.textMuted,
              fontFamily: "'DM Mono', monospace",
              fontSize: 12,
            }}
          >
            {info.getValue()}
          </span>
        ),
      }),
      columnHelper.accessor("narrative", {
        header: "Narrative",
        cell: (info) => (
          <span
            style={{
              maxWidth: 280,
              display: "block",
              overflow: "hidden",
              textOverflow: "ellipsis",
              whiteSpace: "nowrap",
            }}
          >
            {info.getValue()}
          </span>
        ),
      }),
      columnHelper.accessor("category", {
        header: "Category",
        cell: (info) => {
          const idx = stats.categories.findIndex(
            (c) => c.label === info.getValue(),
          );
          const col = PALETTE[idx % PALETTE.length];
          return (
            <span
              style={{
                background: `${col}22`,
                color: col,
                border: `1px solid ${col}44`,
                borderRadius: 100,
                padding: "2px 10px",
                fontSize: 11,
                fontWeight: 600,
                whiteSpace: "nowrap",
              }}
            >
              {info.getValue()}
            </span>
          );
        },
      }),
      columnHelper.accessor("debit", {
        header: "Debit Amount",
        cell: (info) =>
          info.getValue() > 0 ? (
            <span
              style={{ color: th.debit, fontFamily: "'DM Mono', monospace" }}
            >
              {fmt(info.getValue())}
            </span>
          ) : (
            <span style={{ color: th.textFaint }}>—</span>
          ),
      }),
      columnHelper.accessor("credit", {
        header: "Credit Amount",
        cell: (info) =>
          info.getValue() > 0 ? (
            <span
              style={{ color: th.credit, fontFamily: "'DM Mono', monospace" }}
            >
              {fmt(info.getValue())}
            </span>
          ) : (
            <span style={{ color: th.textFaint }}>—</span>
          ),
      }),
    ],
    [stats.categories, th],
  );

  const merchantColumns = useMemo(
    () => [
      columnHelper.accessor("merchant", {
        header: "Merchant",
        cell: (info) => (
          <span
            style={{
              maxWidth: 260,
              display: "block",
              overflow: "hidden",
              textOverflow: "ellipsis",
              whiteSpace: "nowrap",
            }}
          >
            {info.getValue()}
          </span>
        ),
      }),
      columnHelper.accessor("spent", {
        header: "Spend",
        cell: (info) => (
          <span style={{ color: th.debit, fontFamily: "'DM Mono', monospace" }}>
            {fmt(info.getValue())}
          </span>
        ),
      }),
      columnHelper.accessor("visits", {
        header: "Visits",
        cell: (info) => (
          <span style={{ fontFamily: "'DM Mono', monospace" }}>
            {info.getValue().toLocaleString()}
          </span>
        ),
      }),
      columnHelper.accessor("avg", {
        header: "Avg / Visit",
        cell: (info) => (
          <span style={{ fontFamily: "'DM Mono', monospace" }}>
            {fmt(info.getValue())}
          </span>
        ),
      }),
    ],
    [th],
  );

  const catColumns = useMemo(
    () => [
      columnHelper.accessor("label", {
        header: "Category",
        cell: (info) => {
          const idx = stats.categories.findIndex(
            (c) => c.label === info.getValue(),
          );
          return (
            <span
              style={{ color: PALETTE[idx % PALETTE.length], fontWeight: 600 }}
            >
              {info.getValue()}
            </span>
          );
        },
      }),
      columnHelper.accessor("count", {
        header: "Transactions",
        cell: (info) => (
          <span style={{ fontFamily: "'DM Mono', monospace" }}>
            {info.getValue()}
          </span>
        ),
      }),
      columnHelper.accessor("debit", {
        header: "Total Spent",
        cell: (info) => (
          <span style={{ color: th.debit, fontFamily: "'DM Mono', monospace" }}>
            {fmt(info.getValue())}
          </span>
        ),
      }),
      columnHelper.accessor("credit", {
        header: "Total Received",
        cell: (info) =>
          info.getValue() > 0 ? (
            <span
              style={{ color: th.credit, fontFamily: "'DM Mono', monospace" }}
            >
              {fmt(info.getValue())}
            </span>
          ) : (
            <span style={{ color: th.textFaint }}>—</span>
          ),
      }),
      columnHelper.display({
        id: "view",
        header: "",
        cell: ({ row }) => (
          <button
            onClick={(e) => {
              e.stopPropagation();
              setReturnCategory(row.original.label);
              setActiveCategory(row.original.label);
              setActiveTab("transactions");
            }}
            style={{
              background: th.accentDim,
              color: th.accent,
              border: `1px solid ${th.accentBorder}`,
              borderRadius: 6,
              padding: "3px 10px",
              fontSize: 11,
              cursor: "pointer",
              fontFamily: "inherit",
              fontWeight: 600,
            }}
          >
            View →
          </button>
        ),
      }),
    ],
    [stats.categories, th],
  );

  useEffect(() => {
    if (activeTab === "overview" && returnCategory) {
      const el = document.getElementById(getCategoryRowId(returnCategory));
      if (el) {
        el.scrollIntoView({ behavior: "smooth", block: "center" });
      }
    }
  }, [activeTab, returnCategory]);

  const ctProps = { th };

  return (
    <div
      style={{
        minHeight: "100vh",
        background: th.bg,
        fontFamily: "'DM Sans', sans-serif",
        color: th.text,
        transition: "background 0.3s, color 0.3s",
      }}
    >
      <div
        style={{
          position: "fixed",
          top: 0,
          left: 0,
          right: 0,
          height: 300,
          background: `radial-gradient(ellipse at 50% -10%, ${th.glow} 0%, transparent 70%)`,
          pointerEvents: "none",
        }}
      />

      {/* ── Header ── */}
      <header
        className='dashboard-header'
        style={{
          alignItems: "center",
          justifyContent: "space-between",
          padding: "16px 28px",
          borderBottom: `1px solid ${th.border}`,
          position: "sticky",
          top: 0,
          zIndex: 100,
          background: th.headerBg,
          backdropFilter: "blur(16px)",
        }}
      >
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <span
            style={{
              background: th.accent,
              color: dark ? "#0d0d1a" : "#fff",
              width: 28,
              height: 28,
              borderRadius: 8,
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              fontWeight: 900,
              fontSize: 14,
            }}
          >
            $
          </span>
          <span
            style={{ fontWeight: 700, fontSize: 16, letterSpacing: "-0.01em" }}
          >
            Transaction Analyser
          </span>
        </div>
        <div className='tab-buttons' style={{ display: "flex", gap: 4 }}>
          {["overview", "transactions"].map((tab) => (
            <button
              key={tab}
              onClick={() => setActiveTab(tab)}
              style={{
                background: activeTab === tab ? th.accentDim : "transparent",
                color: activeTab === tab ? th.accent : th.textMuted,
                border:
                  activeTab === tab
                    ? `1px solid ${th.accentBorder}`
                    : "1px solid transparent",
                borderRadius: 8,
                padding: "6px 16px",
                fontSize: 13,
                fontWeight: 600,
                cursor: "pointer",
                fontFamily: "inherit",
                transition: "all 0.15s",
                textTransform: "capitalize",
              }}
            >
              {tab}
            </button>
          ))}
        </div>
        <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
          <button
            onClick={toggleDark}
            title='Toggle theme'
            style={{
              background: th.btnBg,
              border: `1px solid ${th.border}`,
              borderRadius: 8,
              padding: "6px 10px",
              cursor: "pointer",
              fontSize: 15,
              fontFamily: "inherit",
            }}
          >
            {dark ? "☀️" : "🌙"}
          </button>
          <button
            onClick={onReset}
            style={{
              background: th.btnBg,
              border: `1px solid ${th.border}`,
              borderRadius: 8,
              padding: "6px 14px",
              color: th.textMuted,
              fontSize: 12,
              cursor: "pointer",
              fontFamily: "inherit",
            }}
          >
            ↑ New file
          </button>
        </div>
      </header>

      <main style={{ maxWidth: 1320, margin: "0 auto", padding: "24px 20px" }}>
        {/* ── Stat cards ── */}
        <div
          style={{
            display: "flex",
            gap: 12,
            flexWrap: "wrap",
            marginBottom: 24,
          }}
        >
          <StatCard
            label='Total Spent'
            value={fmt(stats.totalSpent)}
            sub={`sum of ${transactions.filter((t) => t.debit > 0).length} debit rows`}
            accent={th.debit}
            th={th}
          />
          <StatCard
            label='Total Received'
            value={fmt(stats.totalReceived)}
            sub={`sum of ${transactions.filter((t) => t.credit > 0).length} credit rows`}
            accent={th.credit}
            th={th}
          />
          <StatCard
            label='Net Position'
            value={fmt(stats.net)}
            sub={stats.net >= 0 ? "net surplus" : "net deficit"}
            accent={stats.net >= 0 ? th.credit : th.debit}
            th={th}
          />
          <StatCard
            label='Categories'
            value={stats.categories.length}
            sub={`across ${transactions.length} transactions`}
            accent='#8b5cf6'
            th={th}
          />
        </div>

        {/* ── OVERVIEW TAB ── */}
        {activeTab === "overview" && (
          <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
            {/* Row 1: area chart (wide) + donut (narrow) */}
            <div
              className='dashboard-grid dashboard-grid--wide'
              style={{
                minWidth: 0,
              }}
            >
              <div
                style={{
                  background: th.bgCard,
                  border: `1px solid ${th.border}`,
                  borderRadius: 16,
                  padding: "20px 20px 12px",
                  boxShadow: "0 1px 4px rgba(0,0,0,0.05)",
                  minWidth: 0,
                }}
              >
                <div style={{ fontWeight: 700, fontSize: 15, color: th.text }}>
                  Monthly Cash Flow
                </div>
                <div
                  style={{
                    fontSize: 12,
                    color: th.textMuted,
                    marginTop: 2,
                    marginBottom: 12,
                  }}
                >
                  Debit vs credit by month
                </div>
                <ResponsiveContainer width='100%' height={240}>
                  <AreaChart
                    data={stats.monthly}
                    margin={{ top: 10, right: 8, left: 0, bottom: 0 }}
                  >
                    <defs>
                      <linearGradient id='gD' x1='0' y1='0' x2='0' y2='1'>
                        <stop
                          offset='5%'
                          stopColor={th.debit}
                          stopOpacity={0.25}
                        />
                        <stop
                          offset='95%'
                          stopColor={th.debit}
                          stopOpacity={0}
                        />
                      </linearGradient>
                      <linearGradient id='gC' x1='0' y1='0' x2='0' y2='1'>
                        <stop
                          offset='5%'
                          stopColor={th.credit}
                          stopOpacity={0.25}
                        />
                        <stop
                          offset='95%'
                          stopColor={th.credit}
                          stopOpacity={0}
                        />
                      </linearGradient>
                    </defs>
                    <CartesianGrid
                      strokeDasharray='3 3'
                      stroke={th.chartGrid}
                    />
                    <XAxis
                      dataKey='label'
                      tick={{ fill: th.chartTick, fontSize: 11 }}
                      tickLine={false}
                      axisLine={false}
                    />
                    <YAxis
                      tick={{ fill: th.chartTick, fontSize: 11 }}
                      tickLine={false}
                      axisLine={false}
                      tickFormatter={(v) => `$${(v / 1000).toFixed(0)}k`}
                      width={48}
                    />
                    <Tooltip content={<ChartTooltip {...ctProps} />} />
                    <Legend
                      wrapperStyle={{
                        fontSize: 11,
                        color: th.textMuted,
                        paddingTop: 8,
                      }}
                    />
                    <Area
                      type='monotone'
                      dataKey='debit'
                      name='Spent'
                      stroke={th.debit}
                      fill='url(#gD)'
                      strokeWidth={2}
                      dot={false}
                    />
                    <Area
                      type='monotone'
                      dataKey='credit'
                      name='Received'
                      stroke={th.credit}
                      fill='url(#gC)'
                      strokeWidth={2}
                      dot={false}
                    />
                  </AreaChart>
                </ResponsiveContainer>
              </div>

              <div
                style={{
                  background: th.bgCard,
                  border: `1px solid ${th.border}`,
                  borderRadius: 16,
                  padding: "20px 16px 12px",
                  boxShadow: "0 1px 4px rgba(0,0,0,0.05)",
                  minWidth: 0,
                  display: "flex",
                  flexDirection: "column",
                }}
              >
                <div style={{ fontWeight: 700, fontSize: 15, color: th.text }}>
                  Spending Mix
                </div>
                <div
                  style={{
                    fontSize: 12,
                    color: th.textMuted,
                    marginTop: 2,
                    marginBottom: 8,
                  }}
                >
                  Click a slice to drill in
                </div>
                <ResponsiveContainer width='100%' height={200}>
                  <PieChart>
                    <Pie
                      data={stats.categories.filter((c) => c.debit > 0)}
                      dataKey='debit'
                      nameKey='label'
                      cx='50%'
                      cy='50%'
                      outerRadius={80}
                      innerRadius={46}
                      paddingAngle={2}
                      onClick={(d) => {
                        setActiveCategory(d.label);
                        setActiveTab("transactions");
                      }}
                      style={{ cursor: "pointer" }}
                    >
                      {stats.categories
                        .filter((c) => c.debit > 0)
                        .map((_, i) => (
                          <Cell
                            key={i}
                            fill={PALETTE[i % PALETTE.length]}
                            opacity={0.9}
                          />
                        ))}
                    </Pie>
                    <Tooltip
                      formatter={(v) => fmt(v)}
                      contentStyle={{
                        background: th.tooltipBg,
                        border: `1px solid ${th.border}`,
                        borderRadius: 8,
                        fontSize: 12,
                        color: th.text,
                      }}
                    />
                  </PieChart>
                </ResponsiveContainer>
                {/* Inline legend */}
                <div
                  style={{
                    display: "flex",
                    flexDirection: "column",
                    gap: 5,
                    marginTop: 4,
                  }}
                >
                  {stats.categories
                    .filter((c) => c.debit > 0)
                    .slice(0, 6)
                    .map((cat, i) => (
                      <div
                        key={cat.label}
                        onClick={() => {
                          setActiveCategory(cat.label);
                          setActiveTab("transactions");
                        }}
                        style={{
                          display: "flex",
                          alignItems: "center",
                          gap: 6,
                          fontSize: 11,
                          cursor: "pointer",
                        }}
                      >
                        <span
                          style={{
                            width: 8,
                            height: 8,
                            borderRadius: 2,
                            background: PALETTE[i % PALETTE.length],
                            flexShrink: 0,
                          }}
                        />
                        <span
                          style={{
                            color: th.textMuted,
                            flex: 1,
                            overflow: "hidden",
                            textOverflow: "ellipsis",
                            whiteSpace: "nowrap",
                          }}
                        >
                          {cat.label}
                        </span>
                        <span
                          style={{
                            color: th.text,
                            fontFamily: "'DM Mono', monospace",
                          }}
                        >
                          {fmt(cat.debit)}
                        </span>
                      </div>
                    ))}
                  {stats.categories.filter((c) => c.debit > 0).length > 6 && (
                    <div
                      style={{
                        fontSize: 11,
                        color: th.textMuted,
                        paddingLeft: 14,
                      }}
                    >
                      +{stats.categories.filter((c) => c.debit > 0).length - 6}{" "}
                      more categories
                    </div>
                  )}
                </div>
              </div>
            </div>

            {/* ── Row 3: Monthly Spending (bar + line) ── */}
            <div
              style={{
                background: th.bgCard,
                border: `1px solid ${th.border}`,
                borderRadius: 16,
                padding: "20px 20px 16px",
                boxShadow: "0 1px 4px rgba(0,0,0,0.05)",
              }}
            >
              <div style={{ fontWeight: 700, fontSize: 16, color: th.text }}>
                Monthly Spending
              </div>
              <div
                style={{
                  fontSize: 12,
                  color: th.textMuted,
                  marginTop: 2,
                  marginBottom: 20,
                }}
              >
                Withdrawals by month with 3-month rolling average trend
              </div>

              <div
                className='dashboard-grid dashboard-grid--half'
                style={{
                  minWidth: 0,
                }}
              >
                {/* Bar chart — withdrawals by month */}
                <div>
                  <div
                    style={{
                      fontSize: 12,
                      fontWeight: 700,
                      color: th.textMuted,
                      textTransform: "uppercase",
                      letterSpacing: "0.08em",
                      marginBottom: 12,
                    }}
                  >
                    Withdrawals by Month
                  </div>
                  <ResponsiveContainer width='100%' height={220}>
                    <BarChart
                      data={stats.monthly}
                      margin={{ top: 4, right: 4, left: 0, bottom: 4 }}
                      barCategoryGap='30%'
                    >
                      <CartesianGrid
                        strokeDasharray='3 3'
                        stroke={th.chartGrid}
                        vertical={false}
                      />
                      <XAxis
                        dataKey='label'
                        tick={{ fill: th.chartTick, fontSize: 11 }}
                        tickLine={false}
                        axisLine={false}
                      />
                      <YAxis
                        tick={{ fill: th.chartTick, fontSize: 11 }}
                        tickLine={false}
                        axisLine={false}
                        tickFormatter={(v) => `$${(v / 1000).toFixed(0)}k`}
                        width={44}
                      />
                      <Tooltip content={<ChartTooltip {...ctProps} />} />
                      <Bar
                        dataKey='debit'
                        name='Withdrawals'
                        radius={[4, 4, 0, 0]}
                      >
                        {stats.monthly.map((m, i) => {
                          // Highlight the peak bar
                          const isPeak =
                            m.debit ===
                            Math.max(...stats.monthly.map((x) => x.debit));
                          return (
                            <Cell
                              key={i}
                              fill={isPeak ? th.debit : th.accent}
                              opacity={isPeak ? 1 : 0.65}
                            />
                          );
                        })}
                      </Bar>
                    </BarChart>
                  </ResponsiveContainer>
                </div>

                {/* Line chart — trend */}
                <div>
                  <div
                    style={{
                      fontSize: 12,
                      fontWeight: 700,
                      color: th.textMuted,
                      textTransform: "uppercase",
                      letterSpacing: "0.08em",
                      marginBottom: 12,
                    }}
                  >
                    Trend{" "}
                    <span
                      style={{
                        fontSize: 10,
                        fontWeight: 400,
                        color: th.textMuted,
                      }}
                    >
                      (with 3-month avg)
                    </span>
                  </div>
                  <ResponsiveContainer width='100%' height={220}>
                    <LineChart
                      data={stats.monthly}
                      margin={{ top: 4, right: 4, left: 0, bottom: 4 }}
                    >
                      <CartesianGrid
                        strokeDasharray='3 3'
                        stroke={th.chartGrid}
                      />
                      <XAxis
                        dataKey='label'
                        tick={{ fill: th.chartTick, fontSize: 11 }}
                        tickLine={false}
                        axisLine={false}
                      />
                      <YAxis
                        tick={{ fill: th.chartTick, fontSize: 11 }}
                        tickLine={false}
                        axisLine={false}
                        tickFormatter={(v) => `$${(v / 1000).toFixed(0)}k`}
                        width={44}
                      />
                      <Tooltip content={<ChartTooltip {...ctProps} />} />
                      <Legend
                        wrapperStyle={{
                          fontSize: 11,
                          color: th.textMuted,
                          paddingTop: 8,
                        }}
                      />
                      {stats.monthly.length > 0 && (
                        <ReferenceLine
                          y={
                            stats.monthly.reduce((s, m) => s + m.debit, 0) /
                            stats.monthly.length
                          }
                          stroke={th.textMuted}
                          strokeDasharray='4 4'
                          strokeOpacity={0.5}
                          label={{
                            value: "avg",
                            fill: th.textMuted,
                            fontSize: 10,
                            position: "right",
                          }}
                        />
                      )}
                      <Line
                        type='monotone'
                        dataKey='debit'
                        name='Withdrawals'
                        stroke={th.debit}
                        strokeWidth={2}
                        dot={{ fill: th.debit, r: 3 }}
                        activeDot={{ r: 5 }}
                      />
                      <Line
                        type='monotone'
                        dataKey='avg'
                        name='3-mo avg'
                        stroke={th.accent}
                        strokeWidth={2}
                        strokeDasharray='5 3'
                        dot={false}
                      />
                    </LineChart>
                  </ResponsiveContainer>
                </div>
              </div>

              {/* Monthly stats strip */}
              {stats.monthly.length > 0 &&
                (() => {
                  const debits = stats.monthly.map((m) => m.debit);
                  const avg = debits.reduce((s, v) => s + v, 0) / debits.length;
                  const peak = Math.max(...debits);
                  const low = Math.min(...debits);
                  const peakM = stats.monthly.find((m) => m.debit === peak);
                  const lowM = stats.monthly.find((m) => m.debit === low);
                  return (
                    <div
                      style={{
                        display: "flex",
                        gap: 12,
                        marginTop: 20,
                        flexWrap: "wrap",
                      }}
                    >
                      {[
                        {
                          label: "Monthly Average",
                          value: fmt(avg),
                          color: th.accent,
                        },
                        {
                          label: `Peak (${peakM?.label})`,
                          value: fmt(peak),
                          color: th.debit,
                        },
                        {
                          label: `Lowest (${lowM?.label})`,
                          value: fmt(low),
                          color: th.credit,
                        },
                        {
                          label: "Months Tracked",
                          value: stats.monthly.length,
                          color: "#8b5cf6",
                        },
                      ].map((s) => (
                        <div
                          key={s.label}
                          style={{
                            flex: "1 1 120px",
                            background: `${s.color}12`,
                            border: `1px solid ${s.color}30`,
                            borderRadius: 10,
                            padding: "10px 14px",
                          }}
                        >
                          <div
                            style={{
                              fontSize: 11,
                              color: th.textMuted,
                              marginBottom: 2,
                            }}
                          >
                            {s.label}
                          </div>
                          <div
                            style={{
                              fontSize: 16,
                              fontWeight: 700,
                              color: s.color,
                              fontFamily: "'DM Mono', monospace",
                            }}
                          >
                            {s.value}
                          </div>
                        </div>
                      ))}
                    </div>
                  );
                })()}
            </div>

            {/* Row 2: horizontal bar — full width */}
            <div
              style={{
                background: th.bgCard,
                border: `1px solid ${th.border}`,
                borderRadius: 16,
                padding: "20px 20px 12px",
                boxShadow: "0 1px 4px rgba(0,0,0,0.05)",
              }}
            >
              <div style={{ fontWeight: 700, fontSize: 15, color: th.text }}>
                Top Spending Categories
              </div>
              <div
                style={{
                  fontSize: 12,
                  color: th.textMuted,
                  marginTop: 2,
                  marginBottom: 12,
                }}
              >
                Total debit amount per category
              </div>
              <ResponsiveContainer
                width='100%'
                height={Math.max(
                  180,
                  Math.min(
                    stats.categories.filter((c) => c.debit > 0).length,
                    10,
                  ) * 36,
                )}
              >
                <BarChart
                  data={stats.categories
                    .filter((c) => c.debit > 0)
                    .slice(0, 10)}
                  layout='vertical'
                  margin={{ top: 4, right: 60, left: 8, bottom: 4 }}
                >
                  <CartesianGrid
                    strokeDasharray='3 3'
                    stroke={th.chartGrid}
                    horizontal={false}
                  />
                  <XAxis
                    type='number'
                    tick={{ fill: th.chartTick, fontSize: 11 }}
                    tickLine={false}
                    axisLine={false}
                    tickFormatter={(v) => `$${(v / 1000).toFixed(0)}k`}
                  />
                  <YAxis
                    type='category'
                    dataKey='label'
                    tick={{ fill: th.text, fontSize: 12 }}
                    tickLine={false}
                    axisLine={false}
                    width={120}
                  />
                  <Tooltip content={<ChartTooltip {...ctProps} />} />
                  <Bar dataKey='debit' name='Spent' radius={[0, 6, 6, 0]}>
                    {stats.categories
                      .filter((c) => c.debit > 0)
                      .slice(0, 10)
                      .map((_, i) => (
                        <Cell key={i} fill={PALETTE[i % PALETTE.length]} />
                      ))}
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            </div>

            {/* Category summary table */}
            <div
              style={{
                background: th.bgCard,
                border: `1px solid ${th.border}`,
                borderRadius: 16,
                padding: 20,
                boxShadow: "0 1px 4px rgba(0,0,0,0.05)",
              }}
            >
              <div
                style={{
                  fontWeight: 700,
                  fontSize: 15,
                  color: th.text,
                  marginBottom: 16,
                }}
              >
                Category Summary
              </div>
              <DataTable
                data={stats.categories}
                columns={catColumns}
                th={th}
                rowProps={(row) => ({
                  id: getCategoryRowId(row.original.label),
                  className: "clickable-row",
                  onClick: () => {
                    setReturnCategory(row.original.label);
                    setActiveCategory(row.original.label);
                    setActiveTab("transactions");
                  },
                })}
              />
            </div>

            {/* ── Top Merchants (bar chart + table) ── */}
            <div
              style={{
                background: th.bgCard,
                border: `1px solid ${th.border}`,
                borderRadius: 16,
                padding: "20px 20px 24px",
                boxShadow: "0 1px 4px rgba(0,0,0,0.05)",
              }}
            >
              <div
                style={{
                  display: "flex",
                  justifyContent: "space-between",
                  alignItems: "flex-start",
                  gap: 12,
                  flexWrap: "wrap",
                }}
              >
                <div>
                  <div
                    style={{
                      fontWeight: 700,
                      fontSize: 16,
                      color: th.text,
                    }}
                  >
                    Top Merchants
                  </div>
                  <div
                    style={{
                      fontSize: 12,
                      color: th.textMuted,
                      marginTop: 2,
                    }}
                  >
                    Top 10 merchants by spend, plus a full merchant visit and
                    spend table.
                  </div>
                </div>
                <div
                  style={{
                    fontSize: 12,
                    color: th.textMuted,
                    minWidth: 180,
                    textAlign: "right",
                  }}
                >
                  {stats.topMerchants.length} merchants shown in chart ·{" "}
                  {stats.merchants.length.toLocaleString()} merchants total
                </div>
              </div>

              <div style={{ marginTop: 18 }}>
                <ResponsiveContainer width='100%' height={420}>
                  <BarChart
                    data={stats.topMerchants}
                    layout='vertical'
                    margin={{ top: 4, right: 24, left: 8, bottom: 4 }}
                  >
                    <CartesianGrid
                      strokeDasharray='3 3'
                      stroke={th.chartGrid}
                      horizontal={false}
                    />
                    <XAxis
                      type='number'
                      tick={{ fill: th.chartTick, fontSize: 11 }}
                      tickLine={false}
                      axisLine={false}
                      tickFormatter={(v) => `$${(v / 1000).toFixed(0)}k`}
                    />
                    <YAxis
                      type='category'
                      dataKey='merchant'
                      tick={{ fill: th.text, fontSize: 12 }}
                      tickFormatter={(value) =>
                        String(value).length > 24
                          ? `${String(value).slice(0, 24).trim()}…`
                          : value
                      }
                      tickLine={false}
                      axisLine={false}
                      width={180}
                    />
                    <Tooltip content={<ChartTooltip {...ctProps} />} />
                    <Bar dataKey='spent' name='Spend' radius={[0, 6, 6, 0]}>
                      {stats.topMerchants.map((_, i) => (
                        <Cell key={i} fill={PALETTE[i % PALETTE.length]} />
                      ))}
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              </div>

              <div style={{ marginTop: 24 }}>
                <div
                  style={{
                    fontSize: 12,
                    fontWeight: 700,
                    letterSpacing: "0.08em",
                    textTransform: "uppercase",
                    color: th.textMuted,
                    marginBottom: 12,
                  }}
                >
                  Merchant spend and visit summary
                </div>
                <DataTable
                  data={stats.merchants}
                  columns={merchantColumns}
                  th={th}
                />
              </div>
            </div>

            {/* ── Key Observations ── */}
            {stats.observations.length > 0 && (
              <div
                style={{
                  background: th.bgCard,
                  border: `1px solid ${th.border}`,
                  borderRadius: 16,
                  padding: 24,
                  boxShadow: "0 1px 4px rgba(0,0,0,0.05)",
                }}
              >
                <div
                  style={{
                    display: "flex",
                    alignItems: "center",
                    gap: 10,
                    marginBottom: 20,
                  }}
                >
                  <span style={{ fontSize: 20 }}>🔍</span>
                  <div>
                    <div
                      style={{ fontWeight: 700, fontSize: 16, color: th.text }}
                    >
                      Key Observations
                    </div>
                    <div
                      style={{
                        fontSize: 12,
                        color: th.textMuted,
                        marginTop: 2,
                      }}
                    >
                      Automated insights from your spending data
                    </div>
                  </div>
                </div>
                <div
                  className='dashboard-grid dashboard-grid--cards'
                  style={{
                    minWidth: 0,
                  }}
                >
                  {stats.observations.map((obs, i) => {
                    const colors = {
                      info: {
                        bg: `${th.accent}0f`,
                        border: `${th.accent}30`,
                        title: th.accent,
                      },
                      success: {
                        bg: `${th.credit}0f`,
                        border: `${th.credit}30`,
                        title: th.credit,
                      },
                      warning: {
                        bg: `${th.debit}0f`,
                        border: `${th.debit}30`,
                        title: th.debit,
                      },
                    };
                    const c = colors[obs.type] || colors.info;
                    return (
                      <div
                        key={i}
                        style={{
                          background: c.bg,
                          border: `1px solid ${c.border}`,
                          borderRadius: 12,
                          padding: "14px 16px",
                          display: "flex",
                          gap: 12,
                          alignItems: "flex-start",
                        }}
                      >
                        <span
                          style={{
                            fontSize: 22,
                            lineHeight: 1,
                            flexShrink: 0,
                            marginTop: 1,
                          }}
                        >
                          {obs.icon}
                        </span>
                        <div>
                          <div
                            style={{
                              fontWeight: 700,
                              fontSize: 13,
                              color: c.title,
                              marginBottom: 4,
                            }}
                          >
                            {obs.title}
                          </div>
                          <div
                            style={{
                              fontSize: 12,
                              color: th.textMuted,
                              lineHeight: 1.6,
                            }}
                          >
                            {obs.body}
                          </div>
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>
            )}
          </div>
        )}

        {/* ── TRANSACTIONS TAB ── */}
        {activeTab === "transactions" && (
          <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
            <div
              style={{
                display: "flex",
                gap: 6,
                flexWrap: "wrap",
                alignItems: "center",
              }}
            >
              {returnCategory && (
                <button
                  onClick={() => setActiveTab("overview")}
                  style={{
                    background: th.btnBg,
                    color: th.textMuted,
                    border: `1px solid ${th.border}`,
                    borderRadius: 100,
                    padding: "4px 14px",
                    fontSize: 12,
                    fontWeight: 600,
                    cursor: "pointer",
                    fontFamily: "inherit",
                  }}
                >
                  ← Back to summary
                </button>
              )}
              <span
                style={{ fontSize: 12, color: th.textMuted, marginRight: 2 }}
              >
                Filter:
              </span>
              <button
                onClick={() => setActiveCategory(null)}
                style={{
                  background: !activeCategory ? th.accentDim : th.btnBg,
                  color: !activeCategory ? th.accent : th.textMuted,
                  border: !activeCategory
                    ? `1px solid ${th.accentBorder}`
                    : `1px solid ${th.border}`,
                  borderRadius: 100,
                  padding: "4px 14px",
                  fontSize: 12,
                  fontWeight: 600,
                  cursor: "pointer",
                  fontFamily: "inherit",
                }}
              >
                All
              </button>
              {stats.categories.map((cat, i) => (
                <button
                  key={cat.label}
                  onClick={() =>
                    setActiveCategory(
                      cat.label === activeCategory ? null : cat.label,
                    )
                  }
                  style={{
                    background:
                      activeCategory === cat.label
                        ? `${PALETTE[i % PALETTE.length]}22`
                        : th.btnBg,
                    color:
                      activeCategory === cat.label
                        ? PALETTE[i % PALETTE.length]
                        : th.textMuted,
                    border:
                      activeCategory === cat.label
                        ? `1px solid ${PALETTE[i % PALETTE.length]}55`
                        : `1px solid ${th.border}`,
                    borderRadius: 100,
                    padding: "4px 14px",
                    fontSize: 12,
                    fontWeight: 600,
                    cursor: "pointer",
                    fontFamily: "inherit",
                  }}
                >
                  {cat.label}
                </button>
              ))}
            </div>

            {activeCategory && (
              <div
                style={{
                  background: th.accentDim,
                  border: `1px solid ${th.accentBorder}`,
                  borderRadius: 12,
                  padding: "12px 16px",
                  display: "flex",
                  justifyContent: "space-between",
                  alignItems: "center",
                  flexWrap: "wrap",
                  gap: 8,
                }}
              >
                <span style={{ fontWeight: 700, color: th.accent }}>
                  {activeCategory}
                </span>
                <span style={{ color: th.textMuted, fontSize: 13 }}>
                  {filteredTx.length} transactions · Spent{" "}
                  {fmt(filteredTx.reduce((s, t) => s + t.debit, 0))} · Received{" "}
                  {fmt(filteredTx.reduce((s, t) => s + t.credit, 0))}
                </span>
              </div>
            )}

            <div
              style={{
                background: th.bgCard,
                border: `1px solid ${th.border}`,
                borderRadius: 16,
                padding: 20,
                boxShadow: "0 1px 4px rgba(0,0,0,0.05)",
              }}
            >
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

        .dashboard-header {
          display: flex;
          flex-wrap: wrap;
          align-items: center;
          justify-content: space-between;
          gap: 10px;
        }
        .tab-buttons {
          display: flex;
          flex-wrap: wrap;
          gap: 4px;
          justify-content: center;
        }
        .dashboard-grid--wide {
          display: grid;
          grid-template-columns: 2fr 1fr;
          gap: 16px;
          min-width: 0;
        }
        .dashboard-grid--half {
          display: grid;
          grid-template-columns: 1fr 1fr;
          gap: 24px;
          min-width: 0;
        }
        .dashboard-grid--cards {
          display: grid;
          grid-template-columns: repeat(auto-fill, minmax(320px, 1fr));
          gap: 12px;
          min-width: 0;
        }
        @media (max-width: 900px) {
          .dashboard-grid--wide,
          .dashboard-grid--half,
          .dashboard-grid--cards {
            grid-template-columns: 1fr;
          }
          .dashboard-header {
            justify-content: center;
          }
          .dashboard-header > div {
            width: 100%;
          }
          .dashboard-header .tab-buttons {
            justify-content: flex-start;
          }
        }
        @media (max-width: 640px) {
          .dashboard-header {
            padding: 14px 16px;
          }
          .dashboard-header > div {
            width: 100%;
          }
          .dashboard-grid--cards {
            grid-template-columns: 1fr;
          }
          .dashboard-header button {
            width: auto;
          }
        }
        .clickable-row {
          cursor: pointer;
        }
      `}</style>
    </div>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// ROOT
// ─────────────────────────────────────────────────────────────────────────────
function App() {
  const qc = useQueryClient();
  const [transactions, setTransactions] = useState(() => getStoredTransactions());
  const [dark, setDark] = useState(() => getStoredDarkMode());

  useEffect(() => {
    saveTransactions(transactions);
  }, [transactions]);

  useEffect(() => {
    saveDarkMode(dark);
  }, [dark]);

  const handleData = useCallback(
    (data) => {
      qc.setQueryData(["transactions"], data);
      setTransactions(data);
    },
    [qc],
  );

  const handleReset = useCallback(() => {
    qc.removeQueries({ queryKey: ["transactions"] });
    setTransactions(null);
  }, [qc]);

  if (!transactions)
    return (
      <UploadScreen
        onData={handleData}
        dark={dark}
        toggleDark={() => setDark((d) => !d)}
      />
    );
  return (
    <Dashboard
      transactions={transactions}
      onReset={handleReset}
      dark={dark}
      toggleDark={() => setDark((d) => !d)}
    />
  );
}

export default function Root() {
  return (
    <QueryClientProvider client={queryClient}>
      <App />
    </QueryClientProvider>
  );
}
