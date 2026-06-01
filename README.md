/**
 * MedicalTracker.jsx  —  v5
 * Changes from v4:
 *  1. Customer records table shows ALL 24 columns
 *  2. Removed "Branch completion rate" card; replaced with TPA breakdown chart
 *  3. Dynamic aging buckets built from actual min/max of the Ageing column (10-day buckets)
 *  4. Added Pivot page: two pivot tables matching the Excel screenshots
 *     - Pivot 1: Zone > GEO > TPA Name | Count | SumAssured (Crs) | Premium (Crs)
 *     - Pivot 2: Ageing bucket | Count | SumAssured (Crs) | Premium (Crs)
 *     Both with Disposition + Final Status slicer filters
 */

import { useState, useMemo, useCallback, useRef } from "react";
import * as XLSX from "xlsx";
import {
  PieChart, Pie, Cell, BarChart, Bar,
  XAxis, YAxis, Tooltip, ResponsiveContainer, Legend,
} from "recharts";

/* ─── Column name map ─────────────────────────────────────────────────── */
const COL_MAP = {
  "date"                          : "date",
  "tpa app no"                    : "tpaApplicationNo",
  "tpa app number"                : "tpaApplicationNo",
  "tpa application no"            : "tpaApplicationNo",
  "tpa application number"        : "tpaApplicationNo",
  "tpaapplicationnumber"          : "tpaApplicationNo",
  "tpaapplicationno"              : "tpaApplicationNo",
  "loan no"                       : "loanNumber",
  "loan number"                   : "loanNumber",
  "loannumber"                    : "loanNumber",
  "policy#"                       : "policy",
  "policy"                        : "policy",
  "policy no"                     : "policy",
  "policy no."                    : "policy",
  "policy number"                 : "policy",
  "name of the life assured"      : "nameOfLifeAssured",
  "name of life assured"          : "nameOfLifeAssured",
  "life assured name"             : "nameOfLifeAssured",
  "assured name"                  : "nameOfLifeAssured",
  "testcategory"                  : "testCategory",
  "test category"                 : "testCategory",
  "test categories"               : "testCategory",
  "tests"                         : "testCategory",
  "premium"                       : "premium",
  "sumassured"                    : "sumAssured",
  "sum assured"                   : "sumAssured",
  "age"                           : "age",
  "product type"                  : "productType",
  "producttype"                   : "productType",
  "medical type"                  : "medicalType",
  "medicaltype"                   : "medicalType",
  "med type"                      : "medicalType",
  "tpa name"                      : "tpaName",
  "tpaname"                       : "tpaName",
  "gender"                        : "gender",
  "sub status"                    : "substatus",
  "substatus"                     : "substatus",
  "sub-status"                    : "substatus",
  "appointment date"              : "appointmentDate",
  "appointment data"              : "appointmentDate",
  "appt date"                     : "appointmentDate",
  "appt. date"                    : "appointmentDate",
  "appointmentdate"               : "appointmentDate",
  "final status"                  : "finalStatus",
  "finalstatus"                   : "finalStatus",
  "aging"                         : "aging",
  "ageing"                        : "aging",
  "branch"                        : "branch",
  "customer contact no"           : "customerContact",
  "customer contact number"       : "customerContact",
  "customer contact"              : "customerContact",
  "contact number"                : "customerContact",
  "contact no"                    : "customerContact",
  "zone"                          : "zone",
  "state"                         : "state",
  "geo"                           : "geo",
  "disposition"                   : "disposition",
  "date (disposition date)"       : "dispositionDate",
  "disposition date"              : "dispositionDate",
  "dispositiondate"               : "dispositionDate",
  "timestamp"                     : "timestamp",
};

const NUMERIC_FIELDS = new Set(["premium", "sumAssured", "aging", "age"]);
const EXACT_ONLY_KEYS = new Set(["age", "geo", "date", "zone", "state", "branch", "gender", "premium", "policy", "tests"]);

function cleanHeader(raw) {
  return String(raw).replace(/^\uFEFF/, "").toLowerCase().trim();
}

function mapHeader(raw) {
  const key = cleanHeader(raw);
  if (!key) return null;
  if (COL_MAP[key] !== undefined) return COL_MAP[key];
  if (!EXACT_ONLY_KEYS.has(key)) {
    for (const [pattern, field] of Object.entries(COL_MAP)) {
      if (EXACT_ONLY_KEYS.has(pattern)) continue;
      if (pattern.length >= 8 && key.includes(pattern)) return field;
      if (pattern.length >= 8 && pattern.includes(key)) return field;
    }
  }
  return null;
}

function findHeaderRow(rows) {
  for (let i = 0; i < Math.min(rows.length, 10); i++) {
    const matched = rows[i].filter(cell => mapHeader(cell) !== null).length;
    if (matched >= 2) return i;
  }
  return 0;
}

function normaliseFinalStatus(raw) {
  if (!raw) return "";
  const s = String(raw).toLowerCase().trim();
  if (s.includes("complet")) return "completed";
  if (s.includes("not contact") || s.includes("non contact") || s.includes("not reachable") ||
      s.includes("not answer") || s.includes("no response") || s.includes("not respond") ||
      s.includes("contactable")) return "non contactable";
  if (s.includes("reschedul")) return "rescheduled";
  if (s.includes("pending") || s.includes("schedul") || s.includes("process")) return "pending";
  return s;
}

function parseExcel(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: "array", cellDates: true });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
        if (rows.length < 2) return resolve({ data: [], debug: { headerRow: 0, mappedHeaders: [] } });
        const headerRowIdx = findHeaderRow(rows);
        const rawHeaders = rows[headerRowIdx];
        const headers = rawHeaders.map(mapHeader);
        const mappedHeaders = rawHeaders.map((h, i) => ({
          raw: String(h), mapped: headers[i] || "— unmapped —", ok: !!headers[i],
        }));
        const data = rows.slice(headerRowIdx + 1)
          .filter(row => row.some(cell => cell !== "" && cell !== null && cell !== undefined))
          .map(row => {
            const obj = {};
            headers.forEach((key, i) => {
              if (!key) return;
              let val = row[i];
              if (NUMERIC_FIELDS.has(key)) val = parseFloat(String(val).replace(/,/g, "")) || 0;
              if (val instanceof Date) val = val.toLocaleDateString("en-IN", { day: "2-digit", month: "2-digit", year: "2-digit" });
              obj[key] = (val === null || val === undefined) ? "" : val;
            });
            if (obj.finalStatus !== undefined) obj.finalStatus = normaliseFinalStatus(obj.finalStatus);
            return obj;
          });
        resolve({ data, debug: { headerRow: headerRowIdx, mappedHeaders } });
      } catch (err) { reject(err); }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

/* ─── Palette ─────────────────────────────────────────────────────────── */
const STATUS_COLOR = {
  "completed"       : { bg: "#eaf8f0", color: "#1a6b42", dot: "#27a06b" },
  "pending"         : { bg: "#fef5e4", color: "#7d4e05", dot: "#d4870f" },
  "non contactable" : { bg: "#fdecea", color: "#7a1f1f", dot: "#c0392b" },
  "rescheduled"     : { bg: "#e8f2fc", color: "#0d3f73", dot: "#2471b8" },
};
const DISP_COLORS = ["#27a06b","#2471b8","#c0392b","#d4870f","#7c5cbf","#e07b39","#666"];
const AGING_COLORS_ARR = ["#27a06b","#4aab7a","#8bc34a","#ffc107","#ff9800","#f44336","#9c27b0","#e91e63","#795548","#607d8b"];

const agingColor = (d) => d >= 60 ? "#c0392b" : d >= 30 ? "#d4870f" : "#27a06b";
const fmtPrem = (n) => {
  if (!n || isNaN(n)) return "—";
  return n >= 100000 ? `₹${(n / 100000).toFixed(1)}L` : `₹${(n / 1000).toFixed(0)}K`;
};
const toCrs = (n) => {
  if (!n || isNaN(n)) return "0.00";
  return (n / 10000000).toFixed(2);
};

/* ─── Dynamic aging buckets ───────────────────────────────────────────── */
function buildAgingBuckets(data, bucketSize = 10) {
  const ages = data.map(r => r.aging).filter(a => typeof a === "number" && !isNaN(a));
  if (!ages.length) return [];
  const minAge = Math.floor(Math.min(...ages) / bucketSize) * bucketSize;
  const maxAge = Math.ceil(Math.max(...ages) / bucketSize) * bucketSize;
  const buckets = [];
  for (let lo = minAge; lo < maxAge; lo += bucketSize) {
    const hi = lo + bucketSize - 1;
    const label = `${lo}-${hi}`;
    const items = data.filter(r => r.aging >= lo && r.aging <= hi);
    buckets.push({
      name: label, value: items.length,
      sumAssured: items.reduce((s, r) => s + (r.sumAssured || 0), 0),
      premium: items.reduce((s, r) => s + (r.premium || 0), 0),
    });
  }
  return buckets;
}

/* ─── Small components ───────────────────────────────────────────────── */
function Badge({ value }) {
  const s = STATUS_COLOR[String(value).toLowerCase()] || { bg: "#f0f0f0", color: "#555" };
  return (
    <span style={{ display:"inline-block", padding:"2px 10px", borderRadius:20, fontSize:11, fontWeight:600, background:s.bg, color:s.color, whiteSpace:"nowrap" }}>
      {value || "—"}
    </span>
  );
}
function MetricCard({ label, value, sub, color }) {
  return (
    <div style={{ background:"#f7f8fa", borderRadius:10, padding:"14px 18px", minWidth:0 }}>
      <div style={{ fontSize:11, color:"#888", textTransform:"uppercase", letterSpacing:"0.05em", marginBottom:6 }}>{label}</div>
      <div style={{ fontSize:26, fontWeight:700, color:color||"#1a1a2e", lineHeight:1 }}>{value}</div>
      {sub && <div style={{ fontSize:11, color:"#aaa", marginTop:4 }}>{sub}</div>}
    </div>
  );
}
function Card({ title, children, style }) {
  return (
    <div style={{ background:"#fff", border:"1px solid #eee", borderRadius:12, padding:"18px 20px", ...style }}>
      {title && <div style={{ fontSize:13, fontWeight:600, color:"#555", marginBottom:14 }}>{title}</div>}
      {children}
    </div>
  );
}
const ChartTip = ({ active, payload, label }) => {
  if (!active || !payload?.length) return null;
  return (
    <div style={{ background:"#fff", border:"1px solid #e0e0e0", borderRadius:8, padding:"8px 14px", fontSize:12 }}>
      <div style={{ fontWeight:600, marginBottom:2 }}>{payload[0].payload?.fullName || label || payload[0].name}</div>
      <div style={{ color:payload[0].fill||"#333" }}>{payload[0].value} cases</div>
    </div>
  );
};

/* ─── Debug Panel ─────────────────────────────────────────────────────── */
function DebugPanel({ debug, onClose }) {
  if (!debug) return null;
  const unmapped = debug.mappedHeaders.filter(h => !h.ok && h.raw !== "");
  const mapped   = debug.mappedHeaders.filter(h => h.ok);
  return (
    <div style={{ background:"#1a1a2e", color:"#e0e0e0", borderRadius:12, padding:"16px 20px", marginBottom:16, fontSize:12 }}>
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:12 }}>
        <span style={{ fontWeight:600, color:"#fff" }}>Column detection — header on row {debug.headerRow + 1}</span>
        <button onClick={onClose} style={{ background:"none", border:"none", color:"#aaa", cursor:"pointer", fontSize:14 }}>✕</button>
      </div>
      <div style={{ display:"flex", gap:24, flexWrap:"wrap" }}>
        <div>
          <div style={{ color:"#27a06b", fontWeight:600, marginBottom:6 }}>✓ Mapped ({mapped.length})</div>
          {mapped.map((h,i) => (
            <div key={i} style={{ color:"#aaa", marginBottom:3 }}>
              <span style={{ color:"#fff" }}>{h.raw}</span> → <span style={{ color:"#7ec8e3" }}>{h.mapped}</span>
            </div>
          ))}
        </div>
        {unmapped.length > 0 && (
          <div>
            <div style={{ color:"#c0392b", fontWeight:600, marginBottom:6 }}>✕ Unmapped ({unmapped.length})</div>
            {unmapped.map((h,i) => <div key={i} style={{ color:"#e07b7b", marginBottom:3 }}>{h.raw}</div>)}
          </div>
        )}
      </div>
    </div>
  );
}

/* ─── Upload Zone ─────────────────────────────────────────────────────── */
function UploadZone({ onData, fileName }) {
  const [dragging, setDragging] = useState(false);
  const [parsing, setParsing]   = useState(false);
  const [err, setErr]           = useState(null);
  const inputRef = useRef();
  const handle = useCallback(async (file) => {
    if (!file) return;
    const ext = file.name.split(".").pop().toLowerCase();
    if (!["xlsx","xls","csv"].includes(ext)) { setErr("Please upload .xlsx, .xls, or .csv"); return; }
    setErr(null); setParsing(true);
    try {
      const result = await parseExcel(file);
      onData(result.data, file.name, result.debug);
    } catch(e) { setErr("Could not parse: " + e.message); }
    finally { setParsing(false); }
  }, [onData]);
  return (
    <div
      onDragOver={e=>{e.preventDefault();setDragging(true);}}
      onDragLeave={()=>setDragging(false)}
      onDrop={e=>{e.preventDefault();setDragging(false);handle(e.dataTransfer.files[0]);}}
      onClick={()=>inputRef.current.click()}
      style={{ border:`2px dashed ${dragging?"#2471b8":"#cdd5e0"}`, borderRadius:14, padding:"40px 24px", textAlign:"center", cursor:"pointer", background:dragging?"#f0f6ff":"#fafbfc", transition:"all 0.2s", userSelect:"none" }}
    >
      <input ref={inputRef} type="file" accept=".xlsx,.xls,.csv" style={{ display:"none" }} onChange={e=>handle(e.target.files[0])} />
      {parsing
        ? <div style={{ color:"#2471b8", fontSize:14, fontWeight:500 }}>⏳ Parsing file…</div>
        : <>
            <div style={{ fontSize:40, marginBottom:12 }}>📊</div>
            <div style={{ fontSize:15, fontWeight:600, color:"#333", marginBottom:6 }}>
              {fileName ? `Loaded: ${fileName}` : "Upload your medical tracker Excel"}
            </div>
            <div style={{ fontSize:12, color:"#999" }}>Drag & drop or click — .xlsx · .xls · .csv</div>
            {err && <div style={{ color:"#c0392b", fontSize:12, marginTop:10 }}>{err}</div>}
          </>
      }
    </div>
  );
}

/* ─── Pivot Table 1: Zone > GEO > TPA Name ───────────────────────────── */
function PivotByZoneGeoTpa({ data }) {
  const [dispFilter, setDispFilter] = useState("all");
  const [statusFilter, setStatusFilter] = useState("all");
  const [expandedZones, setExpandedZones] = useState({});
  const [expandedGeos, setExpandedGeos] = useState({});

  const dispositions = useMemo(() => ["all", ...new Set(data.map(r => r.disposition).filter(Boolean))], [data]);
  const statuses     = useMemo(() => ["all", ...new Set(data.map(r => r.finalStatus).filter(Boolean))], [data]);

  const filtered = useMemo(() => data.filter(r => {
    if (dispFilter !== "all" && r.disposition !== dispFilter) return false;
    if (statusFilter !== "all" && r.finalStatus !== statusFilter) return false;
    return true;
  }), [data, dispFilter, statusFilter]);

  // Build hierarchy: zone → geo → tpa
  const hierarchy = useMemo(() => {
    const h = {};
    filtered.forEach(r => {
      const z = r.zone || "—"; const g = r.geo || "—"; const t = r.tpaName || "—";
      if (!h[z]) h[z] = {};
      if (!h[z][g]) h[z][g] = {};
      if (!h[z][g][t]) h[z][g][t] = { count:0, sa:0, prem:0 };
      h[z][g][t].count++;
      h[z][g][t].sa   += r.sumAssured || 0;
      h[z][g][t].prem += r.premium    || 0;
    });
    return h;
  }, [filtered]);

  const grandTotal = useMemo(() => ({
    count: filtered.length,
    sa:    filtered.reduce((s,r) => s+(r.sumAssured||0), 0),
    prem:  filtered.reduce((s,r) => s+(r.premium||0),    0),
  }), [filtered]);

  const toggleZone = (z) => setExpandedZones(p => ({...p, [z]: !p[z]}));
  const toggleGeo  = (k) => setExpandedGeos(p =>  ({...p, [k]: !p[k]}));

  const thStyle = { padding:"8px 12px", textAlign:"right", fontSize:11, fontWeight:700, color:"#fff", background:"#1a3a5c", whiteSpace:"nowrap", borderRight:"1px solid #2a4a6c" };
  const thLeft  = { ...thStyle, textAlign:"left" };
  const tdStyle = (indent=0) => ({ padding:"7px 12px", fontSize:12, paddingLeft:indent+12, borderBottom:"1px solid #f0f0f0", textAlign:"left" });
  const tdNum   = { padding:"7px 12px", fontSize:12, textAlign:"right", borderBottom:"1px solid #f0f0f0", fontVariantNumeric:"tabular-nums" };

  return (
    <div>
      {/* Filters */}
      <div style={{ display:"flex", gap:10, marginBottom:16, flexWrap:"wrap", alignItems:"center" }}>
        <span style={{ fontSize:12, color:"#666", fontWeight:600 }}>Filters:</span>
        <select value={dispFilter} onChange={e=>setDispFilter(e.target.value)}
          style={{ fontSize:12, padding:"5px 10px", borderRadius:7, border:"1px solid #ddd", background:"#fff" }}>
          {dispositions.map(d => <option key={d} value={d}>{d === "all" ? "All Dispositions" : d}</option>)}
        </select>
        <select value={statusFilter} onChange={e=>setStatusFilter(e.target.value)}
          style={{ fontSize:12, padding:"5px 10px", borderRadius:7, border:"1px solid #ddd", background:"#fff" }}>
          {statuses.map(s => <option key={s} value={s}>{s === "all" ? "All Final Statuses" : s}</option>)}
        </select>
      </div>

      <div style={{ overflowX:"auto" }}>
        <table style={{ width:"100%", borderCollapse:"collapse", fontSize:12 }}>
          <thead>
            <tr>
              <th style={thLeft}>Zone</th>
              <th style={thLeft}>GEO</th>
              <th style={thLeft}>TPA Name</th>
              <th style={thStyle}>Count of Loan No</th>
              <th style={thStyle}>Sum of SumAssured (Crs)</th>
              <th style={thStyle}>Sum of Premium (Crs)</th>
            </tr>
          </thead>
          <tbody>
            {Object.entries(hierarchy).sort((a,b)=>a[0].localeCompare(b[0])).map(([zone, geos]) => {
              const zTot = { count:0, sa:0, prem:0 };
              Object.values(geos).forEach(tpas => Object.values(tpas).forEach(v => { zTot.count+=v.count; zTot.sa+=v.sa; zTot.prem+=v.prem; }));
              const zExp = expandedZones[zone];
              return [
                // Zone row
                <tr key={`z-${zone}`} style={{ background:"#e8f0fa", cursor:"pointer" }} onClick={()=>toggleZone(zone)}>
                  <td style={{ ...tdStyle(0), fontWeight:700, color:"#1a3a5c" }}>
                    <span style={{ marginRight:6 }}>{zExp ? "▼" : "▶"}</span>{zone}
                  </td>
                  <td style={{ ...tdStyle(0), color:"#888", fontStyle:"italic" }}></td>
                  <td style={{ ...tdStyle(0) }}></td>
                  <td style={{ ...tdNum, fontWeight:700 }}>{zTot.count}</td>
                  <td style={{ ...tdNum, fontWeight:700 }}>{toCrs(zTot.sa)}</td>
                  <td style={{ ...tdNum, fontWeight:700 }}>{toCrs(zTot.prem)}</td>
                </tr>,
                // GEO rows (if expanded)
                ...(zExp ? Object.entries(geos).sort((a,b)=>a[0].localeCompare(b[0])).flatMap(([geo, tpas]) => {
                  const gTot = { count:0, sa:0, prem:0 };
                  Object.values(tpas).forEach(v => { gTot.count+=v.count; gTot.sa+=v.sa; gTot.prem+=v.prem; });
                  const gKey = `${zone}|${geo}`;
                  const gExp = expandedGeos[gKey];
                  return [
                    <tr key={`g-${gKey}`} style={{ background:"#f4f8fc", cursor:"pointer" }} onClick={e=>{e.stopPropagation();toggleGeo(gKey);}}>
                      <td style={{ ...tdStyle(0) }}></td>
                      <td style={{ ...tdStyle(4), fontWeight:600, color:"#2a5a8c" }}>
                        <span style={{ marginRight:6 }}>{gExp ? "▼" : "▶"}</span>{geo}
                      </td>
                      <td style={{ ...tdStyle(4) }}></td>
                      <td style={{ ...tdNum, fontWeight:600 }}>{gTot.count}</td>
                      <td style={{ ...tdNum, fontWeight:600 }}>{toCrs(gTot.sa)}</td>
                      <td style={{ ...tdNum, fontWeight:600 }}>{toCrs(gTot.prem)}</td>
                    </tr>,
                    // TPA rows (if geo expanded)
                    ...(gExp ? Object.entries(tpas).sort((a,b)=>a[0].localeCompare(b[0])).map(([tpa, v]) => (
                      <tr key={`t-${gKey}|${tpa}`} style={{ background:"#fff" }}
                        onMouseEnter={e=>e.currentTarget.style.background="#fafeff"}
                        onMouseLeave={e=>e.currentTarget.style.background="#fff"}>
                        <td style={{ ...tdStyle(0) }}></td>
                        <td style={{ ...tdStyle(4) }}></td>
                        <td style={{ ...tdStyle(16), color:"#555" }}>{tpa}</td>
                        <td style={tdNum}>{v.count}</td>
                        <td style={tdNum}>{toCrs(v.sa)}</td>
                        <td style={tdNum}>{toCrs(v.prem)}</td>
                      </tr>
                    )) : [])
                  ];
                }) : [])
              ];
            })}
            {/* Grand Total */}
            <tr style={{ background:"#1a3a5c" }}>
              <td colSpan={3} style={{ padding:"9px 12px", fontWeight:700, color:"#fff", fontSize:13 }}>Grand Total</td>
              <td style={{ ...tdNum, fontWeight:700, color:"#fff", background:"#1a3a5c" }}>{grandTotal.count}</td>
              <td style={{ ...tdNum, fontWeight:700, color:"#fff", background:"#1a3a5c" }}>{toCrs(grandTotal.sa)}</td>
              <td style={{ ...tdNum, fontWeight:700, color:"#fff", background:"#1a3a5c" }}>{toCrs(grandTotal.prem)}</td>
            </tr>
          </tbody>
        </table>
      </div>
      <div style={{ fontSize:11, color:"#aaa", marginTop:8 }}>Click Zone / GEO rows to expand/collapse</div>
    </div>
  );
}

/* ─── Pivot Table 2: Ageing Buckets ──────────────────────────────────── */
function PivotByAgeing({ data }) {
  const [dispFilter, setDispFilter] = useState("all");
  const [statusFilter, setStatusFilter] = useState("all");

  const dispositions = useMemo(() => ["all", ...new Set(data.map(r => r.disposition).filter(Boolean))], [data]);
  const statuses     = useMemo(() => ["all", ...new Set(data.map(r => r.finalStatus).filter(Boolean))], [data]);

  const filtered = useMemo(() => data.filter(r => {
    if (dispFilter !== "all" && r.disposition !== dispFilter) return false;
    if (statusFilter !== "all" && r.finalStatus !== statusFilter) return false;
    return true;
  }), [data, dispFilter, statusFilter]);

  const buckets = useMemo(() => buildAgingBuckets(filtered, 10), [filtered]);

  const grandTotal = useMemo(() => ({
    count: filtered.length,
    sa:    filtered.reduce((s,r) => s+(r.sumAssured||0), 0),
    prem:  filtered.reduce((s,r) => s+(r.premium||0),    0),
  }), [filtered]);

  const thStyle = { padding:"8px 12px", textAlign:"right", fontSize:11, fontWeight:700, color:"#fff", background:"#1a3a5c", whiteSpace:"nowrap", borderRight:"1px solid #2a4a6c" };
  const thLeft  = { ...thStyle, textAlign:"left" };
  const tdNum   = { padding:"7px 12px", fontSize:12, textAlign:"right", borderBottom:"1px solid #f0f0f0", fontVariantNumeric:"tabular-nums" };

  return (
    <div>
      <div style={{ display:"flex", gap:10, marginBottom:16, flexWrap:"wrap", alignItems:"center" }}>
        <span style={{ fontSize:12, color:"#666", fontWeight:600 }}>Filters:</span>
        <select value={dispFilter} onChange={e=>setDispFilter(e.target.value)}
          style={{ fontSize:12, padding:"5px 10px", borderRadius:7, border:"1px solid #ddd", background:"#fff" }}>
          {dispositions.map(d => <option key={d} value={d}>{d === "all" ? "All Dispositions" : d}</option>)}
        </select>
        <select value={statusFilter} onChange={e=>setStatusFilter(e.target.value)}
          style={{ fontSize:12, padding:"5px 10px", borderRadius:7, border:"1px solid #ddd", background:"#fff" }}>
          {statuses.map(s => <option key={s} value={s}>{s === "all" ? "All Final Statuses" : s}</option>)}
        </select>
      </div>

      <div style={{ overflowX:"auto" }}>
        <table style={{ width:"100%", borderCollapse:"collapse", fontSize:12 }}>
          <thead>
            <tr>
              <th style={thLeft}>Ageing Bucket</th>
              <th style={thStyle}>Count of Loan No</th>
              <th style={thStyle}>Sum of SumAssured (Crs)</th>
              <th style={thStyle}>Sum of Premium (Crs)</th>
            </tr>
          </thead>
          <tbody>
            {buckets.map((b, i) => (
              <tr key={b.name} style={{ background: i % 2 === 0 ? "#fff" : "#f9fbff" }}
                onMouseEnter={e=>e.currentTarget.style.background="#eef5ff"}
                onMouseLeave={e=>e.currentTarget.style.background=i%2===0?"#fff":"#f9fbff"}>
                <td style={{ padding:"7px 12px", fontSize:12, fontWeight:600, borderBottom:"1px solid #f0f0f0" }}>
                  <span style={{ display:"inline-block", width:10, height:10, borderRadius:2, background:AGING_COLORS_ARR[i % AGING_COLORS_ARR.length], marginRight:8 }}></span>
                  {b.name}
                </td>
                <td style={tdNum}>{b.value}</td>
                <td style={tdNum}>{toCrs(b.sumAssured)}</td>
                <td style={tdNum}>{toCrs(b.premium)}</td>
              </tr>
            ))}
            <tr style={{ background:"#1a3a5c" }}>
              <td style={{ padding:"9px 12px", fontWeight:700, color:"#fff", fontSize:13 }}>Grand Total</td>
              <td style={{ ...tdNum, fontWeight:700, color:"#fff", background:"#1a3a5c" }}>{grandTotal.count}</td>
              <td style={{ ...tdNum, fontWeight:700, color:"#fff", background:"#1a3a5c" }}>{toCrs(grandTotal.sa)}</td>
              <td style={{ ...tdNum, fontWeight:700, color:"#fff", background:"#1a3a5c" }}>{toCrs(grandTotal.prem)}</td>
            </tr>
          </tbody>
        </table>
      </div>
      <div style={{ fontSize:11, color:"#aaa", marginTop:8 }}>Buckets auto-generated from actual min–max of the Ageing column (10-day intervals)</div>
    </div>
  );
}

/* ─── Pivot Page ──────────────────────────────────────────────────────── */
function PivotPage({ data }) {
  return (
    <div>
      <Card title="Pivot 1 — Zone › GEO › TPA Name breakdown" style={{ marginBottom:20 }}>
        <PivotByZoneGeoTpa data={data} />
      </Card>
      <Card title="Pivot 2 — Ageing bucket analysis">
        <PivotByAgeing data={data} />
      </Card>
    </div>
  );
}

/* ─── All 24 column definitions for the customer table ───────────────── */
const ALL_COLUMNS = [
  { col:"date",             label:"Date",            width:90  },
  { col:"tpaApplicationNo", label:"TPA App No.",      width:130 },
  { col:"loanNumber",       label:"Loan No.",         width:120 },
  { col:"policy",           label:"Policy",           width:100 },
  { col:"nameOfLifeAssured",label:"Name",             width:160 },
  { col:"testCategory",     label:"Test Category",    width:160 },
  { col:"premium",          label:"Premium",          width:80  },
  { col:"sumAssured",       label:"Sum Assured",      width:100 },
  { col:"age",              label:"Age",              width:50  },
  { col:"productType",      label:"Product Type",     width:90  },
  { col:"medicalType",      label:"Med. Type",        width:90  },
  { col:"tpaName",          label:"TPA",              width:100 },
  { col:"gender",           label:"Gender",           width:70  },
  { col:"substatus",        label:"Sub Status",       width:200 },
  { col:"appointmentDate",  label:"Appt. Date",       width:90  },
  { col:"finalStatus",      label:"Final Status",     width:130 },
  { col:"aging",            label:"Ageing",           width:65  },
  { col:"branch",           label:"Branch",           width:120 },
  { col:"customerContact",  label:"Contact No.",      width:110 },
  { col:"zone",             label:"Zone",             width:70  },
  { col:"state",            label:"State",            width:90  },
  { col:"geo",              label:"GEO",              width:90  },
  { col:"disposition",      label:"Disposition",      width:180 },
  { col:"dispositionDate",  label:"Disp. Date",       width:90  },
];

/* ─── Main Dashboard ─────────────────────────────────────────────────── */
export default function MedicalTracker() {
  const [allData, setAllData]           = useState([]);
  const [fileName, setFileName]         = useState("");
  const [debugInfo, setDebugInfo]       = useState(null);
  const [showDebug, setShowDebug]       = useState(false);
  const [activePage, setActivePage]     = useState("dashboard"); // "dashboard" | "pivot"
  const [search, setSearch]             = useState("");
  const [zoneFilter, setZoneFilter]     = useState("all");
  const [statusFilter, setStatusFilter] = useState("all");
  const [agingFilter, setAgingFilter]   = useState("all");
  const [sortCol, setSortCol]           = useState("aging");
  const [sortDir, setSortDir]           = useState("desc");
  const [page, setPage]                 = useState(1);
  const PAGE_SIZE = 15;

  const onData = useCallback((rows, name, debug) => {
    setAllData(rows); setFileName(name); setDebugInfo(debug); setShowDebug(true);
    setPage(1); setSearch(""); setZoneFilter("all"); setStatusFilter("all"); setAgingFilter("all");
  }, []);
  const handleFileChange = async (file) => {
    if (!file) return;
    try { const r = await parseExcel(file); onData(r.data, file.name, r.debug); } catch(e) { console.error(e); }
  };

  const zones    = useMemo(() => ["all", ...new Set(allData.map(r => r.zone).filter(Boolean))], [allData]);
  const statuses = useMemo(() => ["all", ...new Set(allData.map(r => r.finalStatus).filter(Boolean))], [allData]);

  const filtered = useMemo(() => allData.filter(r => {
    if (zoneFilter !== "all" && r.zone !== zoneFilter) return false;
    if (statusFilter !== "all" && r.finalStatus !== statusFilter) return false;
    if (agingFilter === "critical" && r.aging < 60) return false;
    if (agingFilter === "high" && (r.aging < 30 || r.aging >= 60)) return false;
    if (agingFilter === "normal" && r.aging >= 30) return false;
    if (search) {
      const q = search.toLowerCase();
      return (
        String(r.nameOfLifeAssured).toLowerCase().includes(q) ||
        String(r.loanNumber).toLowerCase().includes(q) ||
        String(r.tpaApplicationNo).toLowerCase().includes(q) ||
        String(r.branch).toLowerCase().includes(q) ||
        String(r.state).toLowerCase().includes(q)
      );
    }
    return true;
  }), [allData, zoneFilter, statusFilter, agingFilter, search]);

  const sorted = useMemo(() => [...filtered].sort((a, b) => {
    let av = a[sortCol], bv = b[sortCol];
    if (typeof av === "string") { av = av.toLowerCase(); bv = (bv||"").toLowerCase(); }
    if (av == null) return 1; if (bv == null) return -1;
    return sortDir === "asc" ? (av > bv ? 1 : -1) : (av < bv ? 1 : -1);
  }), [filtered, sortCol, sortDir]);

  const pageCount = Math.max(1, Math.ceil(sorted.length / PAGE_SIZE));
  const pageRows  = sorted.slice((page-1)*PAGE_SIZE, page*PAGE_SIZE);
  const handleSort = (col) => {
    if (sortCol === col) setSortDir(d => d==="asc"?"desc":"asc");
    else { setSortCol(col); setSortDir("desc"); }
    setPage(1);
  };

  /* Chart data */
  const statusData = useMemo(() => {
    const c = {};
    filtered.forEach(r => { if (r.finalStatus) c[r.finalStatus] = (c[r.finalStatus]||0)+1; });
    return Object.entries(c).map(([name,value]) => ({name,value}));
  }, [filtered]);

  const dispData = useMemo(() => {
    const c = {};
    filtered.forEach(r => { if (r.disposition) c[r.disposition] = (c[r.disposition]||0)+1; });
    return Object.entries(c).sort((a,b)=>b[1]-a[1])
      .map(([name,value]) => ({ name:name.length>16?name.slice(0,16)+"…":name, fullName:name, value }));
  }, [filtered]);

  const agingBuckets = useMemo(() => buildAgingBuckets(filtered, 10), [filtered]);
  const agingChartData = agingBuckets.map((b,i) => ({ ...b, fill:AGING_COLORS_ARR[i%AGING_COLORS_ARR.length] }));

  const tpaData = useMemo(() => {
    const c = {};
    filtered.forEach(r => { if (r.tpaName) c[r.tpaName] = (c[r.tpaName]||0)+1; });
    return Object.entries(c).sort((a,b)=>b[1]-a[1])
      .map(([name,value]) => ({ name, fullName:name, value }));
  }, [filtered]);

  const metrics = useMemo(() => {
    const total     = filtered.length;
    const completed = filtered.filter(r => r.finalStatus==="completed").length;
    const nc        = filtered.filter(r => r.finalStatus==="non contactable").length;
    const critical  = filtered.filter(r => r.aging>=60).length;
    const pending   = filtered.filter(r => r.finalStatus==="pending").length;
    const avgPrem   = total ? Math.round(filtered.reduce((s,r)=>s+(r.premium||0),0)/total) : 0;
    const compRate  = total ? Math.round((completed/total)*100) : 0;
    return { total, completed, nc, critical, pending, avgPrem, compRate };
  }, [filtered]);

  const Th = ({ col, label, width }) => (
    <th onClick={()=>handleSort(col)} style={{
      padding:"8px 12px", textAlign:"left", fontSize:11, fontWeight:600,
      color:"#666", whiteSpace:"nowrap", background:"#f7f8fa",
      borderBottom:"1px solid #eee", cursor:"pointer", width,
      userSelect:"none", position:"sticky", top:0, zIndex:1,
    }}>
      {label}{sortCol===col?(sortDir==="asc"?" ↑":" ↓"):""}
    </th>
  );

  const isEmpty = allData.length === 0;
  const navBtn = (id, label) => (
    <button onClick={()=>setActivePage(id)} style={{
      padding:"8px 20px", borderRadius:8, cursor:"pointer", fontSize:13, fontWeight:600,
      background: activePage===id ? "#1a3a5c" : "#fff",
      color: activePage===id ? "#fff" : "#555",
      border: activePage===id ? "1px solid #1a3a5c" : "1px solid #ddd",
      transition:"all 0.15s",
    }}>{label}</button>
  );

  return (
    <div style={{ fontFamily:"'DM Sans','Segoe UI',sans-serif", background:"#f4f5f7", minHeight:"100vh", padding:"24px 28px", color:"#1a1a2e" }}>

      {/* Header */}
      <div style={{ display:"flex", alignItems:"flex-start", justifyContent:"space-between", flexWrap:"wrap", gap:12, marginBottom:20 }}>
        <div>
          <h1 style={{ fontSize:22, fontWeight:700, margin:0, letterSpacing:"-0.02em" }}>Medical Checkup Tracker</h1>
          <p style={{ fontSize:13, color:"#888", margin:"4px 0 0" }}>Home Loan Insurance · Sales Team Monitoring</p>
        </div>
        {!isEmpty && (
          <div style={{ display:"flex", gap:8, flexWrap:"wrap" }}>
            {navBtn("dashboard","📊 Dashboard")}
            {navBtn("pivot","🔢 Pivot View")}
            <button onClick={()=>setShowDebug(v=>!v)} style={{ padding:"8px 14px", borderRadius:8, cursor:"pointer", background:"#fff", border:"1px solid #ddd", fontSize:12, color:"#555", fontWeight:500 }}>
              {showDebug?"Hide":"Show"} column map
            </button>
            <label style={{ display:"inline-flex", alignItems:"center", gap:7, padding:"8px 16px", borderRadius:8, cursor:"pointer", background:"#fff", border:"1px solid #ddd", fontSize:13, fontWeight:500, color:"#333" }}>
              📂 Upload new file
              <input type="file" accept=".xlsx,.xls,.csv" style={{ display:"none" }} onChange={e=>{handleFileChange(e.target.files[0]);e.target.value="";}} />
            </label>
          </div>
        )}
      </div>

      {isEmpty ? (
        <div style={{ maxWidth:560, margin:"60px auto" }}>
          <UploadZone onData={onData} fileName={fileName} />
          <p style={{ textAlign:"center", fontSize:12, color:"#bbb", marginTop:16 }}>Parsed entirely in the browser — data never leaves your device.</p>
        </div>
      ) : (
        <>
          <div style={{ fontSize:12, color:"#888", marginBottom:12, display:"flex", alignItems:"center", gap:8 }}>
            <span style={{ color:"#27a06b", fontWeight:600 }}>✓</span>
            {fileName} — <strong style={{ color:"#555" }}>{allData.length.toLocaleString()} records loaded</strong>
          </div>

          {showDebug && debugInfo && <DebugPanel debug={debugInfo} onClose={()=>setShowDebug(false)} />}

          {/* ── PIVOT PAGE ── */}
          {activePage === "pivot" && <PivotPage data={allData} />}

          {/* ── DASHBOARD PAGE ── */}
          {activePage === "dashboard" && <>
            {/* Filters */}
            <div style={{ display:"flex", gap:8, flexWrap:"wrap", marginBottom:20 }}>
              {[
                { id:"zone",   label:"Zone",   val:zoneFilter,   opts:zones,    set:v=>{setZoneFilter(v);setPage(1);} },
                { id:"status", label:"Status", val:statusFilter, opts:statuses, set:v=>{setStatusFilter(v);setPage(1);} },
                { id:"aging",  label:"Aging",  val:agingFilter,
                  opts:[{value:"all",label:"All Aging"},{value:"critical",label:"Critical (60+ days)"},{value:"high",label:"High (30–59 days)"},{value:"normal",label:"Normal (<30 days)"}],
                  set:v=>{setAgingFilter(v);setPage(1);} },
              ].map(f => (
                <select key={f.id} value={f.val} onChange={e=>f.set(e.target.value)}
                  style={{ fontSize:12, padding:"6px 10px", borderRadius:8, border:"1px solid #ddd", background:"#fff", color:"#333", cursor:"pointer" }}>
                  {f.opts.map(o => typeof o==="string"
                    ? <option key={o} value={o}>{o==="all"?`All ${f.label}s`:o}</option>
                    : <option key={o.value} value={o.value}>{o.label}</option>
                  )}
                </select>
              ))}
            </div>

            {/* Metrics */}
            <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(140px,1fr))", gap:10, marginBottom:20 }}>
              <MetricCard label="Total Cases"     value={metrics.total}        sub="active pipeline" />
              <MetricCard label="Completed"       value={metrics.completed}    sub={`${metrics.compRate}% rate`}  color="#27a06b" />
              <MetricCard label="Non Contactable" value={metrics.nc}           sub="needs follow-up"              color="#c0392b" />
              <MetricCard label="Critical Aging"  value={metrics.critical}     sub="60+ days pending"             color="#c0392b" />
              <MetricCard label="Pending"         value={metrics.pending}      sub="appointment awaited"          color="#d4870f" />
              <MetricCard label="Avg Premium"     value={fmtPrem(metrics.avgPrem)} sub="per policy" />
            </div>

            {/* Charts Row 1 */}
            <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:14, marginBottom:14 }}>
              <Card title="Final status distribution">
                <ResponsiveContainer width="100%" height={200}>
                  <PieChart>
                    <Pie data={statusData} cx="50%" cy="50%" innerRadius={52} outerRadius={78} paddingAngle={3} dataKey="value">
                      {statusData.map((e,i) => <Cell key={i} fill={STATUS_COLOR[e.name]?.dot||DISP_COLORS[i%DISP_COLORS.length]} />)}
                    </Pie>
                    <Tooltip content={<ChartTip />} />
                    <Legend iconSize={9} iconType="square" formatter={v=><span style={{fontSize:11,color:"#555"}}>{v}</span>} />
                  </PieChart>
                </ResponsiveContainer>
              </Card>
              <Card title="Sales disposition breakdown">
                <ResponsiveContainer width="100%" height={200}>
                  <BarChart data={dispData} margin={{top:0,right:8,left:-20,bottom:0}}>
                    <XAxis dataKey="name" tick={{fontSize:10}} interval={0} angle={-12} textAnchor="end" height={38} />
                    <YAxis tick={{fontSize:10}} allowDecimals={false} />
                    <Tooltip content={<ChartTip />} />
                    <Bar dataKey="value" radius={[4,4,0,0]}>
                      {dispData.map((_,i) => <Cell key={i} fill={DISP_COLORS[i%DISP_COLORS.length]} />)}
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              </Card>
            </div>

            {/* Charts Row 2 */}
            <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:14, marginBottom:14 }}>
              <Card title={`Dynamic aging buckets (${agingChartData.length > 0 ? `${agingChartData[0]?.name.split("-")[0]}–${agingChartData[agingChartData.length-1]?.name.split("-")[1]} days` : "—"})`}>
                <ResponsiveContainer width="100%" height={200}>
                  <BarChart data={agingChartData} margin={{top:0,right:8,left:-20,bottom:0}}>
                    <XAxis dataKey="name" tick={{fontSize:10}} interval={0} angle={-12} textAnchor="end" height={38} />
                    <YAxis tick={{fontSize:10}} allowDecimals={false} />
                    <Tooltip content={<ChartTip />} />
                    <Bar dataKey="value" radius={[4,4,0,0]}>
                      {agingChartData.map((e,i) => <Cell key={i} fill={e.fill} />)}
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              </Card>
              <Card title="TPA-wise case distribution">
                <ResponsiveContainer width="100%" height={200}>
                  <BarChart data={tpaData} margin={{top:0,right:8,left:-20,bottom:0}}>
                    <XAxis dataKey="name" tick={{fontSize:11}} />
                    <YAxis tick={{fontSize:10}} allowDecimals={false} />
                    <Tooltip content={<ChartTip />} />
                    <Bar dataKey="value" radius={[4,4,0,0]}>
                      {tpaData.map((_,i) => <Cell key={i} fill={DISP_COLORS[i%DISP_COLORS.length]} />)}
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              </Card>
            </div>

            {/* Customer Table — ALL 24 columns */}
            <Card>
              <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:14, flexWrap:"wrap", gap:8 }}>
                <span style={{ fontSize:13, fontWeight:600, color:"#555" }}>Customer records</span>
                <div style={{ display:"flex", alignItems:"center", gap:10 }}>
                  <input type="text" placeholder="Search name / loan / TPA / branch / state…"
                    value={search} onChange={e=>{setSearch(e.target.value);setPage(1);}}
                    style={{ fontSize:12, padding:"6px 12px", borderRadius:8, border:"1px solid #ddd", width:260, outline:"none" }} />
                  <span style={{ fontSize:11, color:"#aaa", whiteSpace:"nowrap" }}>{sorted.length.toLocaleString()} records</span>
                </div>
              </div>

              <div style={{ overflowX:"auto" }}>
                <table style={{ width:"100%", borderCollapse:"collapse", fontSize:12 }}>
                  <thead>
                    <tr>
                      {ALL_COLUMNS.map(c => <Th key={c.col} col={c.col} label={c.label} width={c.width} />)}
                    </tr>
                  </thead>
                  <tbody>
                    {pageRows.length === 0 ? (
                      <tr><td colSpan={ALL_COLUMNS.length} style={{ textAlign:"center", padding:"40px", color:"#ccc" }}>No records match the current filters.</td></tr>
                    ) : pageRows.map((r, i) => (
                      <tr key={i} style={{ borderBottom:"1px solid #f5f5f5" }}
                        onMouseEnter={e=>e.currentTarget.style.background="#fafafa"}
                        onMouseLeave={e=>e.currentTarget.style.background="transparent"}>
                        <td style={{ padding:"8px 12px", color:"#777", whiteSpace:"nowrap" }}>{r.date}</td>
                        <td style={{ padding:"8px 12px", color:"#555", whiteSpace:"nowrap" }}>{String(r.tpaApplicationNo||"").slice(-16)}</td>
                        <td style={{ padding:"8px 12px", color:"#555", whiteSpace:"nowrap" }}>{String(r.loanNumber||"").slice(-16)}</td>
                        <td style={{ padding:"8px 12px", color:"#777" }}>{r.policy}</td>
                        <td style={{ padding:"8px 12px", fontWeight:500, whiteSpace:"nowrap" }} title={r.nameOfLifeAssured}>{r.nameOfLifeAssured}</td>
                        <td style={{ padding:"8px 12px", color:"#666", maxWidth:160, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }} title={r.testCategory}>{r.testCategory}</td>
                        <td style={{ padding:"8px 12px", whiteSpace:"nowrap" }}>{fmtPrem(r.premium)}</td>
                        <td style={{ padding:"8px 12px", whiteSpace:"nowrap" }}>{fmtPrem(r.sumAssured)}</td>
                        <td style={{ padding:"8px 12px" }}>{r.age}</td>
                        <td style={{ padding:"8px 12px", color:"#666" }}>{r.productType}</td>
                        <td style={{ padding:"8px 12px", textTransform:"capitalize" }}>{r.medicalType}</td>
                        <td style={{ padding:"8px 12px", color:"#555", whiteSpace:"nowrap" }}>{r.tpaName}</td>
                        <td style={{ padding:"8px 12px", textTransform:"capitalize" }}>{r.gender}</td>
                        <td style={{ padding:"8px 12px", color:"#777", maxWidth:200, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }} title={r.substatus}>{r.substatus}</td>
                        <td style={{ padding:"8px 12px", color:"#555", whiteSpace:"nowrap" }}>{r.appointmentDate||"—"}</td>
                        <td style={{ padding:"8px 12px" }}><Badge value={r.finalStatus} /></td>
                        <td style={{ padding:"8px 12px", fontWeight:700, color:agingColor(r.aging), whiteSpace:"nowrap" }}>{r.aging||0}d</td>
                        <td style={{ padding:"8px 12px", color:"#555", whiteSpace:"nowrap" }}>{r.branch}</td>
                        <td style={{ padding:"8px 12px", color:"#666" }}>{r.customerContact}</td>
                        <td style={{ padding:"8px 12px" }}>{r.zone}</td>
                        <td style={{ padding:"8px 12px", color:"#777" }}>{r.state}</td>
                        <td style={{ padding:"8px 12px", color:"#777" }}>{r.geo}</td>
                        <td style={{ padding:"8px 12px", maxWidth:180, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap", color:"#666" }} title={r.disposition}>{r.disposition}</td>
                        <td style={{ padding:"8px 12px", color:"#777", whiteSpace:"nowrap" }}>{r.dispositionDate||"—"}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>

              {/* Pagination */}
              {pageCount > 1 && (
                <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between", marginTop:16, flexWrap:"wrap", gap:8 }}>
                  <span style={{ fontSize:12, color:"#aaa" }}>
                    Showing {((page-1)*PAGE_SIZE)+1}–{Math.min(page*PAGE_SIZE,sorted.length)} of {sorted.length.toLocaleString()}
                  </span>
                  <div style={{ display:"flex", gap:6 }}>
                    {[
                      { label:"««",     action:()=>setPage(1),                         disabled:page===1 },
                      { label:"← Prev", action:()=>setPage(p=>Math.max(1,p-1)),        disabled:page===1 },
                      { label:`${page} / ${pageCount}`, action:null, disabled:true, plain:true },
                      { label:"Next →", action:()=>setPage(p=>Math.min(pageCount,p+1)),disabled:page===pageCount },
                      { label:"»»",     action:()=>setPage(pageCount),                 disabled:page===pageCount },
                    ].map((btn,i) => btn.plain
                      ? <span key={i} style={{ padding:"5px 12px", fontSize:12, color:"#555" }}>{btn.label}</span>
                      : <button key={i} onClick={btn.action} disabled={btn.disabled}
                          style={{ padding:"5px 12px", borderRadius:6, border:"1px solid #ddd", background:"#fff", fontSize:12, cursor:btn.disabled?"not-allowed":"pointer", color:btn.disabled?"#ccc":"#333" }}>
                          {btn.label}
                        </button>
                    )}
                  </div>
                </div>
              )}
            </Card>
          </>}
        </>
      )}
    </div>
  );
}
