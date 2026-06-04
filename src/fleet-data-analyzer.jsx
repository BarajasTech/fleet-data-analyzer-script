import React, { useState, useMemo, useRef, useEffect, useCallback } from "react";
import * as XLSX from "xlsx";
import {
  ComposedChart, Line, Area, XAxis, YAxis, CartesianGrid, Tooltip, Legend,
  ReferenceLine, ReferenceDot, ReferenceArea, BarChart, Bar, ResponsiveContainer, Cell,
} from "recharts";

/* ============================================================
   FLEET DATA ANALYZER — two-vehicle edition
   Event Recorder + Fault Log per vehicle (vehicle 2 optional).
   Combined overlay graph or separate stacked graphs, with
   shared time axis, drag-pan, wheel-zoom, and minimap scroll.
   Handles TELOC exports (title row, split Date/Time columns)
   and fault logs with Excel datetimes + mixed-car exports.
   ============================================================ */

const FONTS = `
@import url('https://fonts.googleapis.com/css2?family=Barlow+Condensed:wght@500;600;700&family=IBM+Plex+Mono:wght@400;500;600&display=swap');
`;

/* Siemens corporate palette: Petrol #009999 (primary), Deep Blue #000028 (text),
   Light Sand #F3F3F0 (background), functional red/green/blue accents */
const C = {
  bg: "#f3f3f0", panel: "#ffffff", panelEdge: "#e0e0da",
  ink: "#000028", dim: "#66667e", faint: "#c0c0cc",
  amber: "#009999", red: "#d72339", green: "#1b8038", cyan: "#00557c", violet: "#7353e5",
};
const SIG_COLORS = [
  "#009999", "#00557c", "#d72339", "#7353e5",
  "#e96401", "#1b8038", "#99700a", "#cc3377",
  "#006b5e", "#5e5e8e", "#4e7b0f", "#0e6fc0",
  "#962b6e", "#2f8f83", "#b0451b", "#44546a",
];
const VEH = {
  1: { tag: "#009999", fault: "#d72339", dash: undefined },
  2: { tag: "#00557c", fault: "#7353e5", dash: "7 4" },
};

/* ---------------- timestamp parsing ---------------- */
/* snap timestamps within 2 ms of an exact second (Excel serial float drift, e.g. :22.999 → :23.000) */
function snapMs(t) {
  const s = Math.round(t / 1000) * 1000;
  return Math.abs(t - s) <= 2 ? s : t;
}
function parseTs(v) {
  if (v == null || v === "") return null;
  if (v instanceof Date) { const t = v.getTime(); return isNaN(t) ? null : snapMs(t); }
  if (typeof v === "number") {
    if (v > 20000 && v < 80000) return snapMs(Math.round((v - 25569) * 86400000));
    if (v > 1e12) return v;
    if (v > 1e9) return v * 1000;
    return null;
  }
  if (typeof v === "string") {
    const s = v.trim();
    if (/^\d{1,2}:\d{2}/.test(s)) return null;
    let t = Date.parse(s);
    if (!isNaN(t)) return t;
    const m = s.match(/^(\d{1,2})[./](\d{1,2})[./](\d{2,4})[ T]?(\d{1,2})?:?(\d{1,2})?:?(\d{1,2})?/);
    if (m) {
      const yr = m[3].length === 2 ? 2000 + +m[3] : +m[3];
      t = new Date(yr, +m[2] - 1, +m[1], +(m[4] || 0), +(m[5] || 0), +(m[6] || 0)).getTime();
      if (!isNaN(t)) return t;
    }
  }
  return null;
}
function parseTimeOfDay(v) {
  if (v == null || v === "") return null;
  if (v instanceof Date) return v.getHours() * 3600000 + v.getMinutes() * 60000 + v.getSeconds() * 1000 + v.getMilliseconds();
  if (typeof v === "number") {
    if (v >= 0 && v < 1) return Math.round(v * 86400000);
    if (v >= 20000) return parseTimeOfDay(new Date(Math.round((v - 25569) * 86400000)));
    return null;
  }
  if (typeof v === "string") {
    const m = v.trim().match(/^(\d{1,2}):(\d{2})(?::(\d{2})(?:[.,](\d{1,3}))?)?\s*(AM|PM)?$/i);
    if (!m) return null;
    let h = +m[1];
    if (m[5]) { const pm = /pm/i.test(m[5]); if (pm && h < 12) h += 12; if (!pm && h === 12) h = 0; }
    return h * 3600000 + +m[2] * 60000 + (+(m[3] || 0)) * 1000 + +((m[4] || "0").padEnd(3, "0"));
  }
  return null;
}
function buildTs(dateVal, timeVal) {
  const base = parseTs(dateVal);
  if (base == null) return null;
  if (timeVal == null) return base;
  const tod = parseTimeOfDay(timeVal);
  if (tod == null) return base;
  const d = new Date(base);
  d.setHours(0, 0, 0, 0);
  return d.getTime() + tod;
}
function fmtTime(t, spanMs) {
  const d = new Date(t), p = (n) => String(n).padStart(2, "0");
  if (spanMs != null && spanMs < 36e5 * 20) return `${p(d.getHours())}:${p(d.getMinutes())}:${p(d.getSeconds())}`;
  return `${p(d.getMonth() + 1)}/${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}`;
}
function fmtFull(t) {
  const d = new Date(t), p = (n) => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}:${p(d.getSeconds())}`;
}
function fmtDay(t) { const d = new Date(t), p = (n) => String(n).padStart(2, "0"); return `${p(d.getMonth() + 1)}/${p(d.getDate())}`; }

/* ---------------- sheet parsing ---------------- */
function detectHeaderRow(aoa) {
  let best = 0, bestScore = -1;
  for (let i = 0; i < Math.min(aoa.length, 15); i++) {
    const row = aoa[i] || [];
    const strs = row.filter((c) => typeof c === "string" && c.trim() !== "");
    const uniq = new Set(strs.map((s) => s.trim().toLowerCase()));
    const score = uniq.size + (row.filter((c) => c != null && c !== "").length > 2 ? 1 : 0);
    if (score > bestScore) { bestScore = score; best = i; }
  }
  return best;
}
function analyzeRows(rows) {
  if (!rows.length) return { cols: [], numericCols: [], textCols: [], tsGuess: null, timeGuess: null };
  const cols = Object.keys(rows[0]);
  const sample = rows.slice(0, Math.min(rows.length, 300));
  const scores = {}, numeric = [], text = [];
  let timeGuess = null;
  for (const c of cols) {
    let tsOk = 0, numOk = 0, todOk = 0, nonNull = 0;
    for (const r of sample) {
      const v = r[c];
      if (v == null || v === "") continue;
      nonNull++;
      if (parseTs(v) != null && (v instanceof Date || typeof v === "string" || (typeof v === "number" && v > 20000))) tsOk++;
      if (typeof v === "number") numOk++;
      if (typeof v === "string" && /^\d{1,2}:\d{2}/.test(v.trim())) todOk++;
    }
    if (!nonNull) continue;
    const nameBoost = /datetime|timestamp|^date$|date /i.test(c) ? 0.35 : /time|zeit|datum|date/i.test(c) ? 0.2 : 0;
    scores[c] = tsOk / nonNull + nameBoost;
    if (todOk / nonNull > 0.8 && !timeGuess) timeGuess = c;
    if (numOk / nonNull > 0.8) numeric.push(c);
    else if (tsOk / nonNull < 0.5 && todOk / nonNull < 0.5) text.push(c);
  }
  const tsGuess = cols.reduce((b, c) => ((scores[c] || 0) > (scores[b] || 0) ? c : b), cols[0]);
  return { cols, numericCols: numeric, textCols: text, tsGuess: (scores[tsGuess] || 0) > 0.4 ? tsGuess : null, timeGuess };
}
function defaultSignals(numericCols, rows) {
  // check a sample of values so binary flags (0/1) don't fill the analog defaults
  const isBinary = (c) => {
    if (!rows?.length) return false;
    const step = Math.max(1, Math.floor(rows.length / 500));
    let seen = 0;
    for (let i = 0; i < rows.length; i += step) {
      const v = rows[i][c];
      if (v == null || v === "" || typeof v !== "number") continue;
      seen++;
      if (v !== 0 && v !== 1) return false;
    }
    return seen > 0;
  };
  const pri = [/speed/i, /current|amp/i, /volt|_v\b|\(v/i, /press/i, /temp/i, /force/i, /load/i];
  const picked = [];
  // pass 1: analog channels matching priority names
  for (const re of pri) for (const c of numericCols) {
    if (re.test(c) && !picked.includes(c) && !/record|id|index/i.test(c) && !isBinary(c)) {
      picked.push(c);
      if (picked.length >= 4) return picked;
    }
  }
  // pass 2: any remaining analog channel
  for (const c of numericCols) {
    if (!picked.includes(c) && !/record|id|index/i.test(c) && !isBinary(c)) picked.push(c);
    if (picked.length >= 4) return picked;
  }
  // pass 3: fall back to anything
  for (const c of numericCols) { if (!picked.includes(c)) picked.push(c); if (picked.length >= 4) break; }
  return picked.slice(0, 4);
}

/* ---------------- demo data ---------------- */
function makeDemoVehicle(carNo, seedShift) {
  const start = Date.now() - 6 * 3600 * 1000;
  const rec = [];
  for (let i = 0; i < 4320; i++) {
    const t = start + i * 5000, phase = ((i + seedShift) % 360) / 360;
    const speed = Math.max(0, 55 * Math.sin(phase * Math.PI) + (Math.random() - 0.5) * 4);
    rec.push({
      Timestamp: new Date(t),
      Speed_mph: +speed.toFixed(1),
      LineVoltage_V: +(750 + 30 * Math.sin((i + seedShift) / 40) + (Math.random() - 0.5) * 12 - (speed > 40 ? 25 : 0)).toFixed(0),
      MotorCurrent_A: +(speed * 6.5 + (Math.random() - 0.5) * 30 + (phase < 0.25 ? 120 : 0)).toFixed(0),
      BrakeCylPressure_psi: +(phase > 0.85 ? 60 + Math.random() * 10 : 5 + Math.random() * 3).toFixed(1),
      Driving: phase > 0.02 && phase < 0.8 ? 1 : 0,
      L_Brake: phase > 0.8 ? 1 : 0,
      ATP_SafetyBrake: phase > 0.92 && phase < 0.96 ? 1 : 0,
      DoorsClosed: phase > 0.02 && phase < 0.97 ? 1 : 0,
    });
  }
  const codes = [["740", "ATP Safety Brake"], ["745", "Event recorder faulty"], ["294", "Tow mode active"], ["879", "Door 7 isolated"], ["740", "ATP Safety Brake"]];
  const faults = [];
  for (let i = 0; i < 30; i++) {
    const [code, desc] = codes[Math.floor(Math.random() * codes.length)];
    faults.push({ DateTime: new Date(start + Math.random() * 6 * 3600 * 1000), CodeID: code, Fault: desc, CarNo: carNo });
  }
  faults.sort((a, b) => a.DateTime - b.DateTime);
  return { rec, faults };
}

/* ---------------- UI atoms ---------------- */
const Label = ({ children }) => (
  <div style={{ fontFamily: "'Barlow Condensed', sans-serif", fontWeight: 600, fontSize: 13, letterSpacing: "0.14em", textTransform: "uppercase", color: C.dim, marginBottom: 6 }}>{children}</div>
);
const Select = ({ value, onChange, options, placeholder }) => (
  <select value={value || ""} onChange={(e) => onChange(e.target.value || null)}
    style={{ width: "100%", background: "#f3f3f0", color: C.ink, border: `1px solid ${C.faint}`, borderRadius: 4, padding: "8px 10px", fontFamily: "'IBM Plex Mono', monospace", fontSize: 12, outline: "none" }}>
    <option value="">{placeholder || "— select —"}</option>
    {options.map((o) => (<option key={o} value={o}>{o}</option>))}
  </select>
);
function Btn({ children, onClick, primary, disabled, small }) {
  return (
    <button onClick={onClick} disabled={disabled}
      style={{
        fontFamily: "'Barlow Condensed', sans-serif", fontWeight: 700, letterSpacing: "0.12em", textTransform: "uppercase",
        fontSize: small ? 12 : 14, padding: small ? "6px 12px" : "10px 22px", cursor: disabled ? "not-allowed" : "pointer",
        background: primary ? (disabled ? "#cfe6e6" : C.amber) : "transparent",
        color: primary ? (disabled ? "#5d8f8f" : "#ffffff") : C.amber,
        border: `1px solid ${primary ? "transparent" : C.amber}`,
        borderRadius: 3, opacity: disabled ? 0.5 : 1, transition: "all .15s",
      }}>{children}</button>
  );
}
const Chip = ({ active, onClick, children, color }) => (
  <div onClick={onClick} style={{
    fontFamily: "'IBM Plex Mono', monospace", fontSize: 11, padding: "4px 10px", borderRadius: 3, cursor: "pointer",
    border: `1px solid ${active ? (color || C.amber) : C.faint}`, color: active ? (color || C.amber) : C.dim,
    background: active ? "#e9f1f1" : "transparent", whiteSpace: "nowrap",
  }}>{children}</div>
);

/* ---------------- signal multi-select dropdown ---------------- */
function SignalDropdown({ label, color, allSignals, selected, onChange }) {
  const [open, setOpen] = useState(false);
  const [q, setQ] = useState("");
  const filtered = useMemo(() => {
    const f = q.trim().toLowerCase();
    return f ? allSignals.filter((c) => c.toLowerCase().includes(f)) : allSignals;
  }, [allSignals, q]);
  const list = useMemo(() => [...selected.filter((s) => filtered.includes(s) || !q), ...filtered.filter((c) => !selected.includes(c))], [filtered, selected, q]);
  const toggle = (c) => onChange(selected.includes(c) ? selected.filter((x) => x !== c) : [...selected, c]);
  return (
    <div style={{ position: "relative" }}>
      <button onClick={() => setOpen(!open)}
        style={{
          fontFamily: "'Barlow Condensed', sans-serif", fontWeight: 700, letterSpacing: "0.12em", textTransform: "uppercase",
          fontSize: 12, padding: "6px 12px", cursor: "pointer", background: open ? "#e9f1f1" : "transparent",
          color: color, border: `1px solid ${color}`, borderRadius: 3,
        }}>
        {label} ({selected.length}/{allSignals.length}) {open ? "▴" : "▾"}
      </button>
      {open && (
        <div style={{ position: "absolute", top: "calc(100% + 6px)", right: 0, zIndex: 50, width: 340, maxWidth: "85vw", background: "#ffffff", border: `1px solid ${C.faint}`, borderRadius: 6, boxShadow: "0 8px 30px rgba(40,50,60,0.18)", padding: 10 }}>
          <div style={{ display: "flex", gap: 6, marginBottom: 8 }}>
            <input autoFocus value={q} onChange={(e) => setQ(e.target.value)} placeholder="Search signals…"
              style={{ flex: 1, background: "#f3f3f0", color: C.ink, border: `1px solid ${C.faint}`, borderRadius: 4, padding: "6px 9px", fontFamily: "'IBM Plex Mono', monospace", fontSize: 11, outline: "none" }} />
            <span onClick={() => setOpen(false)} style={{ color: C.dim, cursor: "pointer", padding: "4px 6px" }}>✕</span>
          </div>
          <div style={{ display: "flex", gap: 6, marginBottom: 8 }}>
            <Btn small onClick={() => onChange([...new Set([...selected, ...filtered])])}>Select all{q ? " matching" : ""} ({filtered.length})</Btn>
            <Btn small onClick={() => onChange(q ? selected.filter((s) => !filtered.includes(s)) : [])}>Clear{q ? " matching" : ""}</Btn>
          </div>
          <div style={{ maxHeight: 280, overflowY: "auto", display: "flex", flexDirection: "column" }}>
            {list.slice(0, 300).map((c) => {
              const on = selected.includes(c);
              const idx = selected.indexOf(c);
              return (
                <div key={c} onClick={() => toggle(c)}
                  style={{ display: "flex", alignItems: "center", gap: 8, padding: "5px 6px", borderRadius: 3, cursor: "pointer", background: on ? "#e9f1f1" : "transparent" }}>
                  <span style={{ width: 13, height: 13, borderRadius: 2, flexShrink: 0, border: `1px solid ${on ? SIG_COLORS[idx % SIG_COLORS.length] : C.faint}`, background: on ? SIG_COLORS[idx % SIG_COLORS.length] : "transparent" }} />
                  <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 11, color: on ? C.ink : C.dim, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{c}</span>
                </div>
              );
            })}
            {!list.length && <div style={{ color: C.dim, fontSize: 11, fontFamily: "'IBM Plex Mono', monospace", padding: 6 }}>No matches</div>}
          </div>
          <div style={{ fontSize: 10, color: C.dim, fontFamily: "'IBM Plex Mono', monospace", marginTop: 6 }}>{allSignals.length} channels · all selectable{selected.length > 12 ? " · tip: >12 lines gets crowded" : ""}</div>
        </div>
      )}
    </div>
  );
}

/* ---------------- file panel ---------------- */
function FilePanel({ title, accent, data, setData, kind }) {
  const inputRef = useRef(null);
  const [error, setError] = useState(null);
  const [busy, setBusy] = useState(false);
  const [sigSearch, setSigSearch] = useState("");

  const loadSheet = (wb, name, headerOverride) => {
    const ws = wb.Sheets[name];
    const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, range: 0, blankrows: false });
    const headerRow = headerOverride != null ? headerOverride : detectHeaderRow(aoa.slice(0, 15));
    const rows = XLSX.utils.sheet_to_json(ws, { defval: null, range: headerRow });
    return { rows, headerRow, det: analyzeRows(rows) };
  };
  const applySheet = (base, wb, name, headerOverride) => {
    const { rows, headerRow, det } = loadSheet(wb, name, headerOverride);
    return {
      ...base, wb, sheet: name, headerRow, rows,
      cols: det.cols, numericCols: det.numericCols, textCols: det.textCols,
      tsCol: det.tsGuess, timeCol: det.timeGuess && det.timeGuess !== det.tsGuess ? det.timeGuess : null,
      signals: kind === "rec" ? defaultSignals(det.numericCols, rows) : [],
      codeCol: kind === "fault" ? (det.cols.find((c) => /code/i.test(c)) || det.textCols.find((c) => /fault|event|err/i.test(c)) || det.textCols[0] || null) : null,
      descCol: kind === "fault" ? (det.textCols.find((c) => /fault|desc|text|message|name/i.test(c)) || null) : null,
      vehCol: kind === "fault" ? (det.cols.find((c) => /car|veh|unit|train/i.test(c)) || null) : null,
      vehPick: null,
    };
  };
  const handleFile = async (file) => {
    setError(null); setBusy(true);
    try {
      // CSV/TSV: read as UTF-8 text so headers like "Δt" aren't mangled by codepage guessing
      const isText = /\.(csv|tsv|txt)$/i.test(file.name);
      const wb = isText
        ? XLSX.read(await file.text(), { type: "string", cellDates: true })
        : XLSX.read(await file.arrayBuffer(), { cellDates: true });
      const sheets = wb.SheetNames;
      let pick = sheets[0], maxR = -1;
      for (const s of sheets) {
        const ref = wb.Sheets[s]?.["!ref"];
        if (!ref) continue;
        const r = XLSX.utils.decode_range(ref).e.r;
        if (r > maxR) { maxR = r; pick = s; }
      }
      setData(applySheet({ fileName: file.name, sheets }, wb, pick, null));
    } catch (e) { setError("Could not parse file: " + e.message); }
    setBusy(false);
  };

  const vehVals = useMemo(() => {
    if (!data?.vehCol || !data?.rows) return [];
    const s = new Set();
    for (const r of data.rows) {
      let v = r[data.vehCol];
      if (v == null || v === "") continue;
      v = typeof v === "number" ? String(Math.round(v)) : String(v);
      if (v.length < 12) s.add(v);
    }
    return [...s].sort();
  }, [data]);

  const filteredSignals = useMemo(() => {
    if (!data) return [];
    const q = sigSearch.trim().toLowerCase();
    const pool = q ? data.numericCols.filter((c) => c.toLowerCase().includes(q)) : data.numericCols;
    const rest = pool.filter((c) => !data.signals.includes(c)).slice(0, 28);
    return [...new Set([...data.signals, ...rest])].filter((c) => !q || c.toLowerCase().includes(q) || data.signals.includes(c));
  }, [data, sigSearch]);

  return (
    <div style={{ flex: 1, minWidth: 290, background: C.panel, border: `1px solid ${C.panelEdge}`, borderTop: `3px solid ${accent}`, borderRadius: 6, padding: 18 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline", marginBottom: 12 }}>
        <div style={{ fontFamily: "'Barlow Condensed', sans-serif", fontWeight: 700, fontSize: 17, letterSpacing: "0.1em", textTransform: "uppercase", color: accent }}>{title}</div>
        {data && <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 11, color: C.dim }}>{data.rows.length.toLocaleString()} rows · {data.cols.length} cols</div>}
      </div>

      {!data ? (
        <div onClick={() => inputRef.current?.click()}
          onDragOver={(e) => e.preventDefault()}
          onDrop={(e) => { e.preventDefault(); const f = e.dataTransfer.files?.[0]; if (f) handleFile(f); }}
          style={{ border: `1.5px dashed ${C.faint}`, borderRadius: 6, padding: "30px 14px", textAlign: "center", cursor: "pointer", color: C.dim, fontFamily: "'IBM Plex Mono', monospace", fontSize: 12 }}>
          <div style={{ fontSize: 24, marginBottom: 6, color: accent }}>{busy ? "…" : "⬆"}</div>
          {busy ? "Parsing…" : "Drop .xlsx / .xls or click"}
          <input ref={inputRef} type="file" accept=".xlsx,.xls,.xlsm,.csv,.tsv,.txt" style={{ display: "none" }} onChange={(e) => { const f = e.target.files?.[0]; if (f) handleFile(f); }} />
        </div>
      ) : (
        <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
          <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 12, color: C.ink, background: "#f3f3f0", border: `1px solid ${C.faint}`, borderRadius: 4, padding: "7px 10px", display: "flex", justifyContent: "space-between", gap: 8 }}>
            <span style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>📄 {data.fileName}</span>
            <span style={{ color: C.red, cursor: "pointer" }} onClick={() => setData(null)}>✕</span>
          </div>

          <div style={{ display: "flex", gap: 10 }}>
            {data.sheets.length > 1 && (
              <div style={{ flex: 2 }}><Label>Sheet</Label><Select value={data.sheet} onChange={(v) => setData(applySheet(data, data.wb, v, null))} options={data.sheets} /></div>
            )}
            <div style={{ flex: 1 }}>
              <Label>Header row</Label>
              <Select value={String(data.headerRow + 1)} onChange={(v) => setData(applySheet(data, data.wb, data.sheet, +v - 1))} options={Array.from({ length: 12 }, (_, i) => String(i + 1))} />
            </div>
          </div>

          <div style={{ display: "flex", gap: 10 }}>
            <div style={{ flex: 1 }}><Label>Date / timestamp</Label><Select value={data.tsCol} onChange={(v) => setData({ ...data, tsCol: v })} options={data.cols} /></div>
            <div style={{ flex: 1 }}><Label>Time (if separate)</Label><Select value={data.timeCol} onChange={(v) => setData({ ...data, timeCol: v })} options={data.cols} placeholder="— none —" /></div>
          </div>

          {kind === "rec" && (
            <div>
              <Label>Signal channels — {data.numericCols.length} numeric found · {data.signals.length} selected</Label>
              <input value={sigSearch} onChange={(e) => setSigSearch(e.target.value)} placeholder="Search signals… e.g. speed, brake"
                style={{ width: "100%", boxSizing: "border-box", background: "#f3f3f0", color: C.ink, border: `1px solid ${C.faint}`, borderRadius: 4, padding: "7px 10px", fontFamily: "'IBM Plex Mono', monospace", fontSize: 12, outline: "none", marginBottom: 8 }} />
              <div style={{ display: "flex", gap: 6, marginBottom: 8 }}>
                <Btn small onClick={() => {
                  const q = sigSearch.trim().toLowerCase();
                  const pool = q ? data.numericCols.filter((c) => c.toLowerCase().includes(q)) : data.numericCols;
                  setData({ ...data, signals: [...new Set([...data.signals, ...pool])] });
                }}>Select all{sigSearch.trim() ? " matching" : ""}</Btn>
                <Btn small onClick={() => {
                  const q = sigSearch.trim().toLowerCase();
                  setData({ ...data, signals: q ? data.signals.filter((s) => !s.toLowerCase().includes(q)) : [] });
                }}>Clear{sigSearch.trim() ? " matching" : ""}</Btn>
              </div>
              <div style={{ display: "flex", flexWrap: "wrap", gap: 6, maxHeight: 150, overflowY: "auto" }}>
                {filteredSignals.map((c) => {
                  const on = data.signals.includes(c);
                  const idx = data.signals.indexOf(c);
                  return (
                    <div key={c}
                      onClick={() => {
                        const s = on ? data.signals.filter((x) => x !== c) : [...data.signals, c];
                        setData({ ...data, signals: s });
                      }}
                      style={{
                        fontFamily: "'IBM Plex Mono', monospace", fontSize: 11, padding: "5px 10px", borderRadius: 3, cursor: "pointer",
                        border: `1px solid ${on ? SIG_COLORS[idx % SIG_COLORS.length] : C.faint}`,
                        color: on ? SIG_COLORS[idx % SIG_COLORS.length] : C.dim,
                        background: on ? "#e9f1f1" : "transparent",
                      }}>{c}</div>
                  );
                })}
                {!data.numericCols.length && <div style={{ color: C.red, fontSize: 12, fontFamily: "'IBM Plex Mono', monospace" }}>No numeric columns — check header row</div>}
              </div>
            </div>
          )}

          {kind === "fault" && (
            <>
              <div style={{ display: "flex", gap: 10 }}>
                <div style={{ flex: 1 }}><Label>Fault code column</Label><Select value={data.codeCol} onChange={(v) => setData({ ...data, codeCol: v })} options={data.cols} /></div>
                <div style={{ flex: 1 }}><Label>Description (optional)</Label><Select value={data.descCol} onChange={(v) => setData({ ...data, descCol: v })} options={data.cols} placeholder="— none —" /></div>
              </div>
              <div style={{ display: "flex", gap: 10 }}>
                <div style={{ flex: 1 }}><Label>Vehicle / car column</Label><Select value={data.vehCol} onChange={(v) => setData({ ...data, vehCol: v, vehPick: null })} options={data.cols} placeholder="— none —" /></div>
                <div style={{ flex: 1 }}>
                  <Label>Filter to vehicle</Label>
                  <Select value={data.vehPick} onChange={(v) => setData({ ...data, vehPick: v })} options={vehVals} placeholder={data.vehCol ? "— all —" : "pick column first"} />
                </div>
              </div>
            </>
          )}
        </div>
      )}
      {error && <div style={{ marginTop: 10, color: C.red, fontFamily: "'IBM Plex Mono', monospace", fontSize: 12 }}>{error}</div>}
    </div>
  );
}

/* ---------------- VCU data logs panel ---------------- */
function VcuPanel({ accent, data, setData }) {
  const inputRef = useRef(null);
  const [error, setError] = useState(null);
  const handleFile = async (file) => {
    setError(null);
    try {
      const events = parseVcuText(await file.text());
      if (!events.length) { setError("No diagnostic entries found — expected 'Diagnostic code / Appeared / Disappeared' blocks."); return; }
      setData({ fileName: file.name, events });
    } catch (e) { setError("Could not parse file: " + e.message); }
  };
  const codes = data ? new Set(data.events.map((e) => e.code)).size : 0;
  const active = data ? data.events.filter((e) => e.open).length : 0;
  return (
    <div style={{ flex: 1, minWidth: 290, background: C.panel, border: `1px solid ${C.panelEdge}`, borderTop: `3px solid ${accent}`, borderRadius: 6, padding: 18 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline", marginBottom: 12 }}>
        <div style={{ fontFamily: "'Barlow Condensed', sans-serif", fontWeight: 700, fontSize: 17, letterSpacing: "0.1em", textTransform: "uppercase", color: accent }}>VCU Data Logs</div>
        {data && <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 11, color: C.dim }}>{data.events.length} events</div>}
      </div>
      {!data ? (
        <div onClick={() => inputRef.current?.click()}
          onDragOver={(e) => e.preventDefault()}
          onDrop={(e) => { e.preventDefault(); const f = e.dataTransfer.files?.[0]; if (f) handleFile(f); }}
          style={{ border: `1.5px dashed ${C.faint}`, borderRadius: 6, padding: "30px 14px", textAlign: "center", cursor: "pointer", color: C.dim, fontFamily: "'IBM Plex Mono', monospace", fontSize: 12 }}>
          <div style={{ fontSize: 24, marginBottom: 6, color: accent }}>⬆</div>
          Drop VCU dump (.txt) or click — optional
          <input ref={inputRef} type="file" accept=".txt,.log,.dat" style={{ display: "none" }} onChange={(e) => { const f = e.target.files?.[0]; if (f) handleFile(f); }} />
        </div>
      ) : (
        <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
          <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 12, color: C.ink, background: "#f3f3f0", border: `1px solid ${C.faint}`, borderRadius: 4, padding: "7px 10px", display: "flex", justifyContent: "space-between", gap: 8 }}>
            <span style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>📄 {data.fileName}</span>
            <span style={{ color: C.red, cursor: "pointer" }} onClick={() => setData(null)}>✕</span>
          </div>
          <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 11, color: C.dim }}>
            {codes} codes · {fmtFull(data.events[0].t)} → {fmtFull(data.events[data.events.length - 1].t)}
            {active > 0 && <span style={{ color: C.red, fontWeight: 600 }}> · {active} still ACTIVE</span>}
          </div>
          <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 10, color: C.dim }}>
            VCU faults carry Appeared→Disappeared durations and render as spans in the ⚠ channel.
          </div>
        </div>
      )}
      {error && <div style={{ marginTop: 10, color: C.red, fontFamily: "'IBM Plex Mono', monospace", fontSize: 12 }}>{error}</div>}
    </div>
  );
}

/* ---------------- tooltip ---------------- */
const DarkTooltip = ({ active, payload, label, labelMap }) => {
  if (!active || !payload?.length) return null;
  const realLabel = typeof label === "number" && labelMap ? labelMap(label) : label;
  return (
    <div style={{ background: "#ffffff", border: `1px solid ${C.faint}`, borderRadius: 4, padding: "8px 12px", fontFamily: "'IBM Plex Mono', monospace", fontSize: 11, maxWidth: 320 }}>
      <div style={{ color: C.dim, marginBottom: 4 }}>{typeof realLabel === "number" ? fmtFull(realLabel) : realLabel}</div>
      {payload.map((p, i) => (
        <div key={i} style={{ color: p.color || C.ink }}>{p.name}: <b>{typeof p.value === "number" ? p.value.toLocaleString(undefined, { maximumFractionDigits: 2 }) : p.value}</b></div>
      ))}
    </div>
  );
};

/* ---------------- TELOC-style digital logic lane ---------------- */
function DigitalLane({ label, color, tagColor, pts, domain, faults, selFault }) {
  const [a, b] = domain;
  const W = 1000, H = 30, span = (b - a) || 1;
  const x = (t) => ((t - a) / span) * W;
  let d = "", prevY = null;
  for (const p of pts) {
    if (p.v == null) { prevY = null; continue; }
    const X = +x(p.t).toFixed(1), Y = p.v >= 0.5 ? 6 : H - 6;
    if (prevY == null) d += `M${X},${Y}`;
    else if (prevY !== Y) d += `L${X},${prevY}L${X},${Y}`;
    else d += `L${X},${Y}`;
    prevY = Y;
  }
  return (
    <div className="lane-row" data-label={label} data-color={color} style={{ display: "flex", alignItems: "stretch", borderTop: `1px solid ${C.panelEdge}` }}>
      <div style={{ flex: "0 0 190px", minWidth: 0, fontSize: 10, fontFamily: "'IBM Plex Mono', monospace", color, padding: "0 8px", borderRight: `1px solid ${C.panelEdge}`, borderLeft: `3px solid ${tagColor || "transparent"}`, display: "flex", alignItems: "center", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", background: "#f7f7f4" }} title={label}>
        {label}
      </div>
      <svg viewBox={`0 0 ${W} ${H}`} preserveAspectRatio="none" style={{ flex: 1, height: 28, display: "block", background: "#ffffff" }}>
        {faults.map((f) => {
          const fx = x(f.t);
          if (fx < 0 || fx > W) return null;
          const sel = `${f.vi}-${f.id}` === selFault;
          return <line key={`${f.vi}-${f.id}`} x1={fx} x2={fx} y1={0} y2={H} stroke={sel ? C.amber : VEH[f.vi].fault} strokeOpacity={sel ? 1 : 0.45} strokeWidth={sel ? 2 : 1} vectorEffect="non-scaling-stroke" strokeDasharray={sel ? undefined : "3 3"} />;
        })}
        <path d={d} fill="none" stroke={color} strokeWidth={1.3} vectorEffect="non-scaling-stroke" />
      </svg>
    </div>
  );
}

/* ---------------- fault occurrences as a pulse-train signal lane ---------------- */
function FaultLane({ label, color, tagColor, events, domain }) {
  const [a, b] = domain;
  const W = 1000, H = 30, span = (b - a) || 1;
  const base = H - 6, top = 6, hw = 2;
  const x = (t) => Math.max(-6, Math.min(W + 6, ((t - a) / span) * W));
  // union of event intervals (instant events get a min-width pulse; open events run to window end)
  const ivs = events.map((e) => {
    const x0 = x(e.t);
    const x1 = e.open ? W : x(e.tEnd != null ? e.tEnd : e.t);
    let lo2 = x0 - hw, hi2 = Math.max(x1, x0) + hw;
    return [lo2, hi2];
  }).sort((p, q) => p[0] - q[0]);
  const merged = [];
  for (const iv of ivs) {
    const L = merged[merged.length - 1];
    if (L && iv[0] <= L[1]) L[1] = Math.max(L[1], iv[1]);
    else merged.push([...iv]);
  }
  let d = `M0,${base}`;
  for (const [s0, s1] of merged) {
    if (s1 < 0 || s0 > W) continue;
    d += `L${s0.toFixed(1)},${base}L${s0.toFixed(1)},${top}L${s1.toFixed(1)},${top}L${s1.toFixed(1)},${base}`;
  }
  d += `L${W},${base}`;
  return (
    <div className="lane-row" data-label={`${label} ×${events.length}`} data-color={color} style={{ display: "flex", alignItems: "stretch", borderTop: `1px solid ${C.panelEdge}` }}>
      <div style={{ flex: "0 0 190px", minWidth: 0, fontSize: 10, fontFamily: "'IBM Plex Mono', monospace", color, padding: "0 8px", borderRight: `1px solid ${C.panelEdge}`, borderLeft: `3px solid ${tagColor || "transparent"}`, display: "flex", alignItems: "center", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", background: "#f7f7f4", fontWeight: 600 }} title={`${label} — ${events.length} in window`}>
        {label}&nbsp;<span style={{ color: C.dim, fontWeight: 400 }}>×{events.length}</span>
      </div>
      <svg viewBox={`0 0 ${W} ${H}`} preserveAspectRatio="none" style={{ flex: 1, height: 28, display: "block", background: "#ffffff" }}>
        <path d={d} fill="none" stroke={color} strokeWidth={1.3} vectorEffect="non-scaling-stroke" />
      </svg>
    </div>
  );
}

/* ---------------- fault-lane code picker ---------------- */
function FaultLaneDropdown({ allCodes, selected, onChange, label = "⚠ Fault lanes", color = C.red, hint = "Each selected code becomes its own pulse lane per vehicle", emptyMeansAll = false, maxSelect = Infinity }) {
  const capped = (arr) => (arr.length > maxSelect ? arr.slice(arr.length - maxSelect) : arr);
  const [open, setOpen] = useState(false);
  const [q, setQ] = useState("");
  const list = useMemo(() => {
    const f = q.trim().toLowerCase();
    return f ? allCodes.filter((p) => p.code.toLowerCase().includes(f) || (p.desc || "").toLowerCase().includes(f)) : allCodes;
  }, [allCodes, q]);
  const toggle = (code) => onChange(capped(selected.includes(code) ? selected.filter((c) => c !== code) : [...selected, code]));
  return (
    <div style={{ position: "relative" }}>
      <button onClick={() => setOpen(!open)}
        style={{
          fontFamily: "'Barlow Condensed', sans-serif", fontWeight: 700, letterSpacing: "0.12em", textTransform: "uppercase",
          fontSize: 12, padding: "6px 12px", cursor: "pointer", background: open ? "#e9f1f1" : "transparent",
          color: color, border: `1px solid ${color}`, borderRadius: 3,
        }}>
        {label} ({selected.length || (emptyMeansAll ? "all" : 0)}) {open ? "▴" : "▾"}
      </button>
      {open && (
        <div style={{ position: "absolute", top: "calc(100% + 6px)", right: 0, zIndex: 50, width: 360, maxWidth: "85vw", background: "#ffffff", border: `1px solid ${C.faint}`, borderRadius: 6, boxShadow: "0 8px 24px rgba(40,50,60,0.18)", padding: 10 }}>
          <div style={{ display: "flex", gap: 6, marginBottom: 8 }}>
            <input autoFocus value={q} onChange={(e) => setQ(e.target.value)} placeholder="Search fault codes…"
              style={{ flex: 1, background: "#f3f3f0", color: C.ink, border: `1px solid ${C.faint}`, borderRadius: 4, padding: "6px 9px", fontFamily: "'IBM Plex Mono', monospace", fontSize: 11, outline: "none" }} />
            <span onClick={() => setOpen(false)} style={{ color: C.dim, cursor: "pointer", padding: "4px 6px" }}>✕</span>
          </div>
          <div style={{ display: "flex", gap: 6, marginBottom: 8 }}><Btn small onClick={() => onChange(capped(allCodes.map((p) => p.code)))}>Select all{Number.isFinite(maxSelect) ? ` (max ${maxSelect})` : ""}</Btn><Btn small onClick={() => onChange([])}>Clear</Btn></div>
          <div style={{ maxHeight: 260, overflowY: "auto", display: "flex", flexDirection: "column" }}>
            {list.slice(0, 200).map((p) => {
              const on = selected.includes(p.code);
              return (
                <div key={p.code} onClick={() => toggle(p.code)}
                  style={{ display: "flex", alignItems: "center", gap: 8, padding: "5px 6px", borderRadius: 3, cursor: "pointer", background: on ? "#e9f1f1" : "transparent" }}>
                  <span style={{ width: 13, height: 13, borderRadius: 2, flexShrink: 0, border: `1px solid ${on ? color : C.faint}`, background: on ? color : "transparent" }} />
                  <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 11, color: on ? C.ink : C.dim, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{p.code}{p.desc ? ` — ${p.desc}` : ""} ({p.count})</span>
                </div>
              );
            })}
            {!list.length && <div style={{ color: C.dim, fontSize: 11, fontFamily: "'IBM Plex Mono', monospace", padding: 6 }}>No matches</div>}
          </div>
          <div style={{ fontSize: 10, color: C.dim, fontFamily: "'IBM Plex Mono', monospace", marginTop: 6 }}>{hint}</div>
        </div>
      )}
    </div>
  );
}

/* ---------------- continuous-time mapping (gap compression) ---------------- */
/* Detects intervals that contain recorder data and builds a piecewise mapping
   real time <-> virtual time where the gaps between intervals are collapsed
   to thin visual seams, so traces render continuously. */
function buildSegments(recDatas) {
  const ivs = [];
  for (const rd of recDatas) {
    if (!rd?.length) continue;
    // adaptive gap threshold: 25x the median sample interval, min 90 s —
    // event recorders log sparsely while idle, so only true recording voids count as gaps
    const dts = [];
    const step = Math.max(1, Math.floor(rd.length / 500));
    for (let i = step; i < rd.length; i += step) dts.push(rd[i].t - rd[i - step].t);
    dts.sort((a, b) => a - b);
    const med = (dts[Math.floor(dts.length / 2)] || 1000) / step;
    const thr = Math.max(med * 25, 90000);
    let a = rd[0].t, prev = rd[0].t;
    for (let i = 1; i < rd.length; i++) {
      const t = rd[i].t;
      if (t - prev > thr) { ivs.push([a, prev]); a = t; }
      prev = t;
    }
    ivs.push([a, prev]);
  }
  if (!ivs.length) return [];
  ivs.sort((x, y) => x[0] - y[0]);
  const out = [ivs[0].slice()];
  for (let i = 1; i < ivs.length; i++) {
    const last = out[out.length - 1];
    if (ivs[i][0] <= last[1] + 10000) last[1] = Math.max(last[1], ivs[i][1]);
    else out.push(ivs[i].slice());
  }
  return out;
}
function buildTimeMap(segs) {
  if (!segs.length) return null;
  const total = segs.reduce((a, [x, y]) => a + (y - x), 0) || 1;
  const pad = Math.max(total * 0.012, 1500); // visual seam width in virtual ms
  const map = []; const gapsV = [];
  let v = 0;
  for (let i = 0; i < segs.length; i++) {
    const [a, b] = segs[i];
    map.push({ a, b, vA: v });
    v += (b - a);
    if (i < segs.length - 1) { gapsV.push({ v0: v, v1: v + pad, ra: b, rb: segs[i + 1][0] }); v += pad; }
  }
  const vEnd = v;
  const toV = (t) => {
    if (t <= segs[0][0]) return Math.max(t - segs[0][0], -pad);
    for (let i = 0; i < map.length; i++) {
      const m = map[i];
      if (t <= m.b) {
        if (t >= m.a) return m.vA + (t - m.a);
        const g = gapsV[i - 1];
        return g.v0 + (g.v1 - g.v0) * (t - g.ra) / Math.max(g.rb - g.ra, 1);
      }
    }
    return vEnd + Math.min(t - segs[segs.length - 1][1], pad);
  };
  const fromV = (v2) => {
    if (v2 <= 0) return segs[0][0] + v2;
    for (let i = 0; i < map.length; i++) {
      const m = map[i];
      if (v2 <= m.vA + (m.b - m.a)) {
        if (v2 >= m.vA) return m.a + (v2 - m.vA);
        const g = gapsV[i - 1];
        return g.ra + (g.rb - g.ra) * (v2 - g.v0) / Math.max(g.v1 - g.v0, 1);
      }
    }
    return segs[segs.length - 1][1] + (v2 - vEnd);
  };
  return { toV, fromV, vEnd, gapsV, segs };
}

/* ---------------- VCU diagnostic-memory text parser ---------------- */
/* Parses VCU dump text: "Diagnostic code: NNN : desc / Appeared: MM-DD-YYYY hh:mm:ss / Disappeared: ..."
   Disappeared 00-00-0000 means the fault is still active. Covers main + subsystem sections. */
function parseVcuText(text) {
  const events = [];
  const re = /Diagnostic code:\s*(\d+)\s*:\s*([^\n\r]+)[\r\n]+\s*Appeared\s*:\s*(\d{2})-(\d{2})-(\d{4})\s+(\d{1,2}):(\d{2}):(\d{2})\s*[\r\n]+\s*Disappeared\s*:\s*(\d{2})-(\d{2})-(\d{4})\s+(\d{1,2}):(\d{2}):(\d{2})/g;
  let m;
  while ((m = re.exec(text))) {
    const t = new Date(+m[5], +m[3] - 1, +m[4], +m[6], +m[7], +m[8]).getTime();
    if (isNaN(t)) continue;
    let tEnd = null, open = false;
    if (m[11] === "0000") open = true;
    else {
      tEnd = new Date(+m[11], +m[9] - 1, +m[10], +m[12], +m[13], +m[14]).getTime();
      if (isNaN(tEnd) || tEnd < t) tEnd = t;
    }
    events.push({ t, tEnd, open, code: m[1], desc: m[2].trim() });
  }
  events.sort((a, b) => a.t - b.t);
  return events;
}
function fmtDur(ms) {
  const s = Math.round(ms / 1000);
  if (s < 90) return s + "s";
  if (s < 5400) return Math.round(s / 60) + "m";
  return (s / 3600).toFixed(1) + "h";
}

/* ---------------- processing helpers ---------------- */
function processRec(rec) {
  if (!rec?.rows?.length || !rec.tsCol) return null;
  const out = [];
  for (const r of rec.rows) {
    const t = buildTs(r[rec.tsCol], rec.timeCol ? r[rec.timeCol] : null);
    if (t == null) continue;
    const row = { t };
    for (const s of rec.signals) {
      const v = r[s];
      row[s] = typeof v === "number" ? v : (v != null && v !== "" && !isNaN(+v) ? +v : null);
    }
    out.push(row);
  }
  out.sort((a, b) => a.t - b.t);
  return out.length ? out : null;
}
function processFault(fault) {
  if (!fault?.rows?.length || !fault.tsCol || !fault.codeCol) return null;
  // optional operational-context columns common in Expert/VCU exports
  const spdCol = (fault.cols || []).find((c) => /speed/i.test(c));
  const upCol = (fault.cols || []).find((c) => /up\s*time/i.test(c));
  const odoCol = (fault.cols || []).find((c) => /miles|odometer|\bkm\b/i.test(c));
  const num = (v) => (typeof v === "number" ? v : (v != null && v !== "" && !isNaN(+v) ? +v : null));
  const out = [];
  fault.rows.forEach((r, i) => {
    const t = buildTs(r[fault.tsCol], fault.timeCol ? r[fault.timeCol] : null);
    if (t == null) return;
    const code = r[fault.codeCol];
    if (code == null || code === "") return;
    let veh = null;
    if (fault.vehCol) {
      const v = r[fault.vehCol];
      veh = v == null ? null : (typeof v === "number" ? String(Math.round(v)) : String(v));
    }
    if (fault.vehPick && veh !== fault.vehPick) return;
    out.push({
      t, id: i,
      code: typeof code === "number" ? String(Math.round(code)) : String(code),
      desc: fault.descCol ? String(r[fault.descCol] ?? "") : "",
      veh,
      spd: spdCol ? num(r[spdCol]) : null,
      up: upCol ? num(r[upCol]) : null,
      odo: odoCol ? num(r[odoCol]) : null,
    });
  });
  out.sort((a, b) => a.t - b.t);
  return out.length ? out : null;
}
function computeRanges(recData, signals) {
  const r = {};
  if (!recData) return r;
  for (const s of signals) {
    let mn = Infinity, mx = -Infinity;
    for (const row of recData) { const v = row[s]; if (v != null) { if (v < mn) mn = v; if (v > mx) mx = v; } }
    r[s] = [mn, mx === mn ? mn + 1 : mx];
  }
  return r;
}
/* signals whose values are only 0/1 render as TELOC-style logic lanes */
function detectBinary(recData, signals) {
  const set = new Set();
  if (!recData?.length) return set;
  const step = Math.max(1, Math.floor(recData.length / 4000));
  for (const s of signals) {
    let seen = 0, ok = true;
    for (let i = 0; i < recData.length; i += step) {
      const v = recData[i][s];
      if (v == null) continue;
      seen++;
      if (v !== 0 && v !== 1) { ok = false; break; }
    }
    if (ok && seen > 0) set.add(s);
  }
  return set;
}
function windowRows(recData, signals, domain, normalize, ranges, prefix) {
  if (!recData || !domain) return [];
  const [a, b] = domain;
  // binary search the first in-window sample, then pad one sample past each edge
  // so lines extend to the chart borders while panning (axis clipping trims the rest)
  let lo = 0, hiB = recData.length;
  while (lo < hiB) { const m = (lo + hiB) >> 1; if (recData[m].t < a) lo = m + 1; else hiB = m; }
  let j = lo;
  while (j < recData.length && recData[j].t <= b) j++;
  const start = Math.max(0, lo - 1), end = Math.min(recData.length, j + 1);
  const inWin = recData.slice(start, end);
  // point budget scales down as more signals are plotted, keeping total points bounded
  const perSeries = Math.max(250, Math.floor(8000 / Math.max(signals.length, 1)));
  const stride = Math.max(1, Math.ceil(inWin.length / perSeries));
  const sampled = stride === 1 ? inWin : inWin.filter((_, i) => i % stride === 0 || i === inWin.length - 1);
  return sampled.map((r) => {
    const o = { t: r.t };
    for (const s of signals) {
      const key = prefix ? `${prefix}${s}` : s;
      if (r[s] == null) { o[key] = null; continue; }
      if (normalize) {
        const [mn, mx] = ranges[s] || [0, 1];
        o[key] = +(((r[s] - mn) / (mx - mn)) * 100).toFixed(2);
      } else o[key] = r[s];
    }
    return o;
  });
}

/* ============================================================ MAIN ============================================================ */
export default function FleetDataAnalyzer() {
  const [rec1, setRec1] = useState(null);
  const [fault1, setFault1] = useState(null);
  const [rec2, setRec2] = useState(null);
  const [fault2, setFault2] = useState(null);
  const [vcu1, setVcu1] = useState(null);
  const [vcu2, setVcu2] = useState(null);
  const [name1, setName1] = useState("Vehicle 1");
  const [name2, setName2] = useState("Vehicle 2");
  const [stage, setStage] = useState("setup");
  const [tab, setTab] = useState("timeline");
  const [viewMode, setViewMode] = useState("combined"); // combined | separate
  const [editOpen, setEditOpen] = useState(false);
  const [chartCfg, setChartCfg] = useState({ title: "", showLegend: true, showGrid: true, showFaults: true, yMin: "", yMax: "", faultStyle: "overlay" });
  const [seriesCfg, setSeriesCfg] = useState({}); // per-series style overrides keyed by cfgKey
  const [domain, setDomain] = useState(null);
  const [selFault, setSelFault] = useState(null); // "vi-id"
  const [normalize, setNormalize] = useState(true);
  const [continuous, setContinuous] = useState(false); // gap-compressed time axis
  const [visExpert, setVisExpert] = useState([]); // Expert fault codes shown; empty = all Expert faults
  const [visVcu, setVisVcu] = useState([]);       // VCU fault codes shown; empty = all VCU faults
  const totalSel = visExpert.length + visVcu.length;
  const srcShown = useCallback((f) => {
    const sel = f.src === "vcu" ? visVcu : visExpert;
    return !sel.length || sel.includes(f.code);
  }, [visExpert, visVcu]);
  const clearFaultSel = () => { setVisExpert([]); setVisVcu([]); };
  /* only explicitly selected codes are drawn on the graph (channel/markers); empty = clean graph */
  const graphShown = useCallback((f) => (f.src === "vcu" ? visVcu : visExpert).includes(f.code), [visExpert, visVcu]);
  /* toggle a code in whichever source(s) contain it (used by Pareto/table clicks) */
  const toggleCode = (code) => {
    const inExpert = expertCodes.some((p) => p.code === code);
    const inVcu = vcuCodes.some((p) => p.code === code);
    if (inExpert) setVisExpert((p) => (p.includes(code) ? p.filter((x) => x !== code) : [...p, code]));
    if (inVcu) setVisVcu((p) => (p.includes(code) ? p.filter((x) => x !== code) : [...p, code]));
  };
  const [syncOpen, setSyncOpen] = useState(false);
  const [offset1, setOffset1] = useState(0); // fault-log time offset vs recorder, seconds
  const [offset2, setOffset2] = useState(0);
  const [faultSearchE, setFaultSearchE] = useState(""); // Expert fault log search
  const [faultSearchV, setFaultSearchV] = useState(""); // VCU fault log search
  const [statsScope, setStatsScope] = useState("all"); // all | v1 | v2
  /* AI analysis (Fault Statistics tab) */
  const [aiMessages, setAiMessages] = useState([]);
  const [aiInput, setAiInput] = useState("");
  const [aiBusy, setAiBusy] = useState(false);
  const [aiError, setAiError] = useState(null);
  const [aiKey, setAiKey] = useState("");      // Anthropic API key for self-hosted deployments; memory-only
  const [aiKeyOpen, setAiKeyOpen] = useState(false);
  const [aiProvider, setAiProvider] = useState("claude"); // claude | ollama
  const [ollamaUrl, setOllamaUrl] = useState("http://localhost:11434");
  const [ollamaModel, setOllamaModel] = useState("llama3.1");
  const [ollamaModels, setOllamaModels] = useState([]);
  const listOllamaModels = async () => {
    try {
      const res = await fetch(`${ollamaUrl.replace(/\/$/, "")}/api/tags`);
      const data = await res.json();
      const names = (data.models || []).map((m) => m.name);
      setOllamaModels(names);
      if (names.length && !names.includes(ollamaModel)) setOllamaModel(names[0]);
      setAiError(names.length ? null : "Connected to Ollama, but no models installed — run e.g. `ollama pull llama3.1`.");
    } catch (e) {
      setAiError("Couldn't reach Ollama at " + ollamaUrl + ". Is it running? For a hosted page you must set OLLAMA_ORIGINS (see note below).");
    }
  };

  /* ----- processed per vehicle ----- */
  const recData1 = useMemo(() => processRec(rec1), [rec1]);
  const recData2 = useMemo(() => processRec(rec2), [rec2]);
  /* Expert rows + VCU events merged into one fault stream per vehicle.
     VCU ids are offset so they never collide with Expert row indexes. */
  const mergeVehicleFaults = (expert, vcu, offSec) => {
    const out = [];
    if (expert) out.push(...expert.map((f) => ({ ...f, src: "expert", tEnd: null, open: false })));
    if (vcu?.events?.length) out.push(...vcu.events.map((e, i) => ({ t: e.t, tEnd: e.tEnd, open: e.open, id: 1000000 + i, code: e.code, desc: e.desc, veh: null, src: "vcu" })));
    if (!out.length) return null;
    out.sort((a, b) => a.t - b.t);
    const o = Math.round(offSec * 1000);
    return o ? out.map((f) => ({ ...f, t: f.t + o, tEnd: f.tEnd != null ? f.tEnd + o : null })) : out;
  };
  const faultData1 = useMemo(() => mergeVehicleFaults(processFault(fault1), vcu1, offset1), [fault1, vcu1, offset1]);
  const faultData2 = useMemo(() => mergeVehicleFaults(processFault(fault2), vcu2, offset2), [fault2, vcu2, offset2]);

  const v2Active = !!(recData2?.length || faultData2?.length);

  const recSpan = useMemo(() => {
    const ts = [];
    if (recData1?.length) ts.push(recData1[0].t, recData1[recData1.length - 1].t);
    if (recData2?.length) ts.push(recData2[0].t, recData2[recData2.length - 1].t);
    return ts.length ? [Math.min(...ts), Math.max(...ts)] : null;
  }, [recData1, recData2]);
  const recSpan1 = recData1?.length ? [recData1[0].t, recData1[recData1.length - 1].t] : null;
  const recSpan2 = recData2?.length ? [recData2[0].t, recData2[recData2.length - 1].t] : null;

  const fullDomain = useMemo(() => {
    const ts = [];
    if (recSpan) ts.push(...recSpan);
    for (const fd of [faultData1, faultData2]) if (fd?.length) ts.push(fd[0].t, fd[fd.length - 1].t);
    return ts.length ? [Math.min(...ts), Math.max(...ts)] : null;
  }, [recSpan, faultData1, faultData2]);

  /* the recorder recording is the authoritative timeline: the graph never leaves it */
  const navFull = recSpan || fullDomain;
  const activeDomain = domain || recSpan || fullDomain;
  const spanMs = activeDomain ? activeDomain[1] - activeDomain[0] : null;

  /* continuous-time mapping: collapse no-data gaps */
  const segments = useMemo(() => buildSegments([recData1, recData2]), [recData1, recData2]);
  const tmap = useMemo(() => (continuous && segments.length ? buildTimeMap(segments) : null), [continuous, segments]);
  const tv = useCallback((t) => (tmap ? tmap.toV(t) : t), [tmap]);
  const tr = useCallback((v) => (tmap ? tmap.fromV(v) : v), [tmap]);
  const vActive = activeDomain ? [tv(activeDomain[0]), tv(activeDomain[1])] : null;

  /* analog vs digital (0/1) split — digital signals render as logic lanes */
  const binSet1 = useMemo(() => detectBinary(recData1, rec1?.signals || []), [recData1, rec1]);
  const binSet2 = useMemo(() => detectBinary(recData2, rec2?.signals || []), [recData2, rec2]);
  const analog1 = useMemo(() => (rec1?.signals || []).filter((s) => !binSet1.has(s)), [rec1, binSet1]);
  const analog2 = useMemo(() => (rec2?.signals || []).filter((s) => !binSet2.has(s)), [rec2, binSet2]);
  const digital1 = useMemo(() => (rec1?.signals || []).filter((s) => binSet1.has(s)), [rec1, binSet1]);
  const digital2 = useMemo(() => (rec2?.signals || []).filter((s) => binSet2.has(s)), [rec2, binSet2]);

  const ranges1 = useMemo(() => computeRanges(recData1, analog1), [recData1, analog1]);
  const ranges2 = useMemo(() => computeRanges(recData2, analog2), [recData2, analog2]);

  /* combined rows (analog only): both vehicles merged with prefixed keys */
  const combinedRows = useMemo(() => {
    const p1 = v2Active ? "V1 " : "";
    const a = windowRows(recData1, analog1, activeDomain, normalize, ranges1, p1);
    const b = v2Active ? windowRows(recData2, analog2, activeDomain, normalize, ranges2, "V2 ") : [];
    return [...a, ...b].sort((x, y) => x.t - y.t);
  }, [recData1, recData2, analog1, analog2, activeDomain, normalize, ranges1, ranges2, v2Active]);

  const sepRows1 = useMemo(() => (viewMode === "separate" ? windowRows(recData1, analog1, activeDomain, normalize, ranges1, "") : []), [viewMode, recData1, analog1, activeDomain, normalize, ranges1]);
  const sepRows2 = useMemo(() => (viewMode === "separate" ? windowRows(recData2, analog2, activeDomain, normalize, ranges2, "") : []), [viewMode, recData2, analog2, activeDomain, normalize, ranges2]);

  /* digital lane data: raw windowed values per vehicle */
  const laneRows1 = useMemo(() => windowRows(recData1, digital1, activeDomain, false, {}, ""), [recData1, digital1, activeDomain]);
  const laneRows2 = useMemo(() => windowRows(recData2, digital2, activeDomain, false, {}, ""), [recData2, digital2, activeDomain]);
  const lanePts = useCallback((laneRows, s) => laneRows.map((r) => ({ t: r.t, v: r[s] })), []);

  /* series definitions — color index follows position in the vehicle's full selection */
  const mkDef = (s, i, prefix, dash) => ({ dataKey: prefix ? `${prefix}${s}` : s, cfgKey: prefix ? `${prefix}${s}` : s, sig: s, defColor: SIG_COLORS[i % SIG_COLORS.length], defDash: dash });

  /* analog series for the combined chart */
  const combinedSeries = useMemo(() => {
    const p1 = v2Active ? "V1 " : "";
    const out = (rec1?.signals || []).map((s, i) => mkDef(s, i, p1, VEH[1].dash)).filter((d) => analog1.includes(d.sig));
    if (v2Active) out.push(...(rec2?.signals || []).map((s, i) => mkDef(s, i, "V2 ", VEH[2].dash)).filter((d) => analog2.includes(d.sig)));
    return out;
  }, [rec1, rec2, v2Active, analog1, analog2]);

  /* separate-mode analog series defs (dataKey unprefixed, cfgKey prefixed) */
  const sepSeries1 = useMemo(() => (rec1?.signals || []).map((s, i) => ({ ...mkDef(s, i, "", undefined), cfgKey: v2Active ? `V1 ${s}` : s })).filter((d) => analog1.includes(d.sig)), [rec1, v2Active, analog1]);
  const sepSeries2 = useMemo(() => (rec2?.signals || []).map((s, i) => ({ ...mkDef(s, i, "", undefined), cfgKey: `V2 ${s}` })).filter((d) => analog2.includes(d.sig)), [rec2, analog2]);

  /* digital lane defs per vehicle (cfgKey always prefixed in 2-vehicle mode for the editor) */
  const laneDefs1 = useMemo(() => (rec1?.signals || []).map((s, i) => ({ ...mkDef(s, i, "", undefined), cfgKey: v2Active ? `V1 ${s}` : s })).filter((d) => digital1.includes(d.sig)), [rec1, v2Active, digital1]);
  const laneDefs2 = useMemo(() => (rec2?.signals || []).map((s, i) => ({ ...mkDef(s, i, "", undefined), cfgKey: `V2 ${s}` })).filter((d) => digital2.includes(d.sig)), [rec2, digital2]);

  /* series visible in the current layout (drives the edit panel) — analog first, then lanes */
  const editSeries = useMemo(() => {
    const lanes = [...laneDefs1, ...(v2Active ? laneDefs2 : [])];
    if (!v2Active || viewMode === "combined") return [...combinedSeries, ...lanes];
    return [...sepSeries1, ...sepSeries2, ...lanes];
  }, [v2Active, viewMode, combinedSeries, sepSeries1, sepSeries2, laneDefs1, laneDefs2]);

  const DASHES = { solid: undefined, dashed: "7 4", dotted: "2 4" };
  const resolveSeries = useCallback((d) => {
    const o = seriesCfg[d.cfgKey] || {};
    return {
      color: o.color || d.defColor,
      dash: o.style ? DASHES[o.style] : d.defDash,
      width: o.width ? +o.width : 1.6,
      curve: o.curve === "smooth" ? "monotone" : o.curve === "step" ? "stepAfter" : "linear",
      area: o.curve === "area",
      dots: !!o.dots,
      label: o.label || d.cfgKey,
    };
  }, [seriesCfg]);
  const updateSeries = (cfgKey, patch) => setSeriesCfg((p) => ({ ...p, [cfgKey]: { ...p[cfgKey], ...patch } }));

  /* faults: merged + per-vehicle visible */
  const mergedFaults = useMemo(() => {
    const a = (faultData1 || []).map((f) => ({ ...f, vi: 1 }));
    const b = (faultData2 || []).map((f) => ({ ...f, vi: 2 }));
    return [...a, ...b].sort((x, y) => x.t - y.t);
  }, [faultData1, faultData2]);

  const visFaults = useCallback((vi) => {
    const src = vi === 1 ? faultData1 : faultData2;
    if (!src || !activeDomain) return [];
    const [a, b] = activeDomain;
    return src.filter((f) => { const end = f.open ? Infinity : (f.tEnd ?? f.t); return f.t <= b && end >= a && srcShown(f); }).map((f) => ({ ...f, vi }));
  }, [faultData1, faultData2, activeDomain, srcShown]);
  const visibleFaults1 = useMemo(() => visFaults(1), [visFaults]);
  const visibleFaults2 = useMemo(() => visFaults(2), [visFaults]);

  /* ----- stats ----- */
  const statsSource = useMemo(() => {
    if (statsScope === "v1") return faultData1 || [];
    if (statsScope === "v2") return faultData2 || [];
    return mergedFaults;
  }, [statsScope, faultData1, faultData2, mergedFaults]);

  const stats = useMemo(() => {
    const fd = statsSource;
    if (!fd?.length) return null;
    const byCode = {};
    for (const f of fd) {
      if (!byCode[f.code]) byCode[f.code] = { code: f.code, count: 0, desc: f.desc };
      byCode[f.code].count++;
      if (!byCode[f.code].desc && f.desc) byCode[f.code].desc = f.desc;
    }
    const pareto = Object.values(byCode).sort((a, b) => b.count - a.count);
    const spanMs2 = Math.max(fd[fd.length - 1].t - fd[0].t, 3600000);
    const spanDays = spanMs2 / 86400000;
    const hourly = spanMs2 < 3 * 86400000;
    const bucketMs = hourly ? 3600000 : 86400000;
    const buckets = {};
    for (const f of fd) { const k = Math.floor(f.t / bucketMs) * bucketMs; buckets[k] = (buckets[k] || 0) + 1; }
    const trend = Object.entries(buckets).map(([k, v]) => ({ t: +k, count: v })).sort((a, b) => a.t - b.t);
    const mtbf = fd.length > 1 ? spanMs2 / (fd.length - 1) : null;
    return { pareto, total: fd.length, spanDays, perDay: fd.length / spanDays, trend, hourly, mtbf };
  }, [statsSource]);

  /* fault dropdown options: union of all codes */
  const allCodes = useMemo(() => {
    const byCode = {};
    for (const f of mergedFaults) {
      if (!byCode[f.code]) byCode[f.code] = { code: f.code, count: 0, desc: f.desc };
      byCode[f.code].count++;
      if (!byCode[f.code].desc && f.desc) byCode[f.code].desc = f.desc;
    }
    return Object.values(byCode).sort((a, b) => b.count - a.count);
  }, [mergedFaults]);

  /* per-source code lists for the two selectors */
  const codesBySrc = useCallback((src) => {
    const byCode = {};
    for (const f of mergedFaults) {
      if ((f.src || "expert") !== src) continue;
      if (!byCode[f.code]) byCode[f.code] = { code: f.code, count: 0, desc: f.desc };
      byCode[f.code].count++;
      if (!byCode[f.code].desc && f.desc) byCode[f.code].desc = f.desc;
    }
    return Object.values(byCode).sort((a, b) => b.count - a.count);
  }, [mergedFaults]);
  const expertCodes = useMemo(() => codesBySrc("expert"), [codesBySrc]);
  const vcuCodes = useMemo(() => codesBySrc("vcu"), [codesBySrc]);

  const ready = recData1?.length && faultData1?.length;

  const zoomToFault = (f) => { setSelFault(`${f.vi}-${f.id}`); setDomain([f.t - 120000, f.t + 120000]); setTab("timeline"); };

  /* ----- signal snapshot at the selected fault ----- */
  const selSnapshot = useMemo(() => {
    if (!selFault) return null;
    const [viStr, idStr] = selFault.split("-");
    const vi = +viStr, id = +idStr;
    const src = vi === 1 ? faultData1 : faultData2;
    const f = src?.find((x) => x.id === id);
    if (!f) return null;
    const recD = vi === 1 ? recData1 : recData2;
    const recCfg = vi === 1 ? rec1 : rec2;
    const binSet = vi === 1 ? binSet1 : binSet2;
    if (!recD?.length) return { f, vi, rows: [], dt: null, outside: true };
    // nearest recorder sample (binary search; both arrays sorted by t)
    let lo = 0, hi = recD.length - 1;
    while (hi - lo > 1) { const m = (lo + hi) >> 1; if (recD[m].t <= f.t) lo = m; else hi = m; }
    const near = Math.abs(recD[lo].t - f.t) <= Math.abs(recD[hi].t - f.t) ? lo : hi;
    const row = recD[near], dt = row.t - f.t;
    const outside = f.t < recD[0].t || f.t > recD[recD.length - 1].t;
    const rows = (recCfg?.signals || []).map((s) => ({ s, v: row[s], binary: binSet.has(s) }));
    return { f, vi, row, dt, rows, outside };
  }, [selFault, faultData1, faultData2, recData1, recData2, rec1, rec2, binSet1, binSet2]);

  /* color lookup for snapshot dots / readout */
  const seriesColorOf = useCallback((vi, s) => {
    const cfgKey = v2Active ? `V${vi} ${s}` : s;
    const all = [...combinedSeries, ...sepSeries1, ...sepSeries2, ...laneDefs1, ...laneDefs2];
    const def = all.find((d) => d.cfgKey === cfgKey && d.sig === s) || { cfgKey, sig: s, defColor: C.ink, defDash: undefined };
    return resolveSeries(def).color;
  }, [v2Active, combinedSeries, sepSeries1, sepSeries2, laneDefs1, laneDefs2, resolveSeries]);

  /* dots drawn on the analog chart at the fault instant */
  const snapDots = useMemo(() => {
    if (!selSnapshot?.row || selSnapshot.outside) return [];
    const { vi, row, f } = selSnapshot;
    const sigs = vi === 1 ? analog1 : analog2;
    const rng = vi === 1 ? ranges1 : ranges2;
    return sigs.map((s) => {
      const v = row[s];
      if (v == null) return null;
      const y = normalize ? ((v - (rng[s]?.[0] ?? 0)) / (((rng[s]?.[1] ?? 1) - (rng[s]?.[0] ?? 0)) || 1)) * 100 : v;
      return { x: f.t, y, color: seriesColorOf(vi, s), vi };
    }).filter(Boolean);
  }, [selSnapshot, analog1, analog2, ranges1, ranges2, normalize, seriesColorOf]);

  /* step through faults (respecting the Faults-shown selection) */
  const navList = useMemo(() => mergedFaults.filter((f) => srcShown(f)), [mergedFaults, srcShown]);
  const navFault = (dir) => {
    if (!navList.length) return;
    let idx = selFault ? navList.findIndex((f) => `${f.vi}-${f.id}` === selFault) : -1;
    idx = idx === -1 ? (dir > 0 ? 0 : navList.length - 1) : Math.min(Math.max(idx + dir, 0), navList.length - 1);
    zoomToFault(navList[idx]);
  };

  /* ----- AI analysis: build a compact data context and query the Anthropic API ----- */
  const buildAiContext = () => {
    const L = [];
    L.push(`Fleet: light-rail vehicles. Vehicle 1: ${name1}${v2Active ? `, Vehicle 2: ${name2}` : ""}.`);
    if (recSpan) L.push(`TELOC event recorder coverage: ${fmtFull(recSpan[0])} to ${fmtFull(recSpan[1])}. Plotted signals: ${[...(rec1?.signals || []), ...(rec2?.signals || [])].slice(0, 14).join(", ")}.`);

    /* per-source summaries */
    for (const srcKey of ["expert", "vcu"]) {
      const evs = mergedFaults.filter((f) => (f.src || "expert") === srcKey);
      if (!evs.length) continue;
      const by = {};
      evs.forEach((f) => { if (!by[f.code]) by[f.code] = { c: 0, d: f.desc }; by[f.code].c++; if (!by[f.code].d) by[f.code].d = f.desc; });
      const top = Object.entries(by).sort((a, b) => b[1].c - a[1].c).slice(0, 15).map(([k, v]) => `${k} x${v.c} ${v.d}`).join("; ");
      const spanDays = Math.max((evs[evs.length - 1].t - evs[0].t) / 86400000, 0.04);
      L.push(`${srcKey === "vcu" ? "VCU diagnostic memory (with Appeared/Disappeared durations)" : "Expert fault log (instantaneous events, 1s resolution)"}: ${evs.length} events across ${Object.keys(by).length} codes, ${fmtFull(evs[0].t)} to ${fmtFull(evs[evs.length - 1].t)} (${(evs.length / spanDays).toFixed(1)}/day). Most frequent: ${top}.`);
      const active = evs.filter((f) => f.open);
      if (active.length) L.push(`STILL-ACTIVE (uncleared) VCU faults: ${active.map((f) => `${f.code} ${f.desc} since ${fmtFull(f.t)}`).join("; ")}.`);
      const longest = evs.filter((f) => f.tEnd != null).sort((a, b) => (b.tEnd - b.t) - (a.tEnd - a.t)).slice(0, 5);
      if (longest.length && srcKey === "vcu") L.push(`Longest-duration VCU faults: ${longest.map((f) => `${f.code} ${f.desc} ${fmtDur(f.tEnd - f.t)} (${fmtFull(f.t)})`).join("; ")}.`);
    }

    /* recorder operational stats over the current window */
    const opsFor = (recD, recCfg, binSet, vn) => {
      if (!recD?.length || !activeDomain) return;
      const [a, b] = activeDomain;
      const rows = recD.filter((r) => r.t >= a && r.t <= b);
      if (!rows.length) return;
      const stride = Math.max(1, Math.ceil(rows.length / 2000));
      const sample = rows.filter((_, i) => i % stride === 0);
      const sigs = (recCfg?.signals || []).filter((sg) => !binSet.has(sg)).slice(0, 6);
      const parts = sigs.map((sg) => {
        let mn = Infinity, mx = -Infinity, sum = 0, n = 0;
        for (const r of sample) { const v = r[sg]; if (v == null) continue; n++; sum += v; if (v < mn) mn = v; if (v > mx) mx = v; }
        return n ? `${sg}: min ${mn.toFixed(1)}, max ${mx.toFixed(1)}, avg ${(sum / n).toFixed(1)}` : null;
      }).filter(Boolean);
      let motion = "";
      const spdSig = sigs.find((sg) => /speed/i.test(sg));
      if (spdSig) {
        let m = 0, n = 0;
        for (const r of sample) { const v = r[spdSig]; if (v == null) continue; n++; if (v > 1) m++; }
        if (n) motion = ` Vehicle in motion ~${Math.round((100 * m) / n)}% of the window.`;
      }
      if (parts.length) L.push(`${vn} recorder stats over the current window: ${parts.join("; ")}.${motion}`);
    };
    opsFor(recData1, rec1, binSet1, name1);
    if (v2Active) opsFor(recData2, rec2, binSet2, name2);

    /* window fault detail with operational context */
    const winF = [...visibleFaults1, ...visibleFaults2].sort((a, b) => a.t - b.t).slice(0, 50);
    if (winF.length && activeDomain) {
      L.push(`Faults intersecting the current analysis window ${fmtFull(activeDomain[0])} to ${fmtFull(activeDomain[1])}:\n` +
        winF.map((f) => `${fmtFull(f.t)} [${f.src === "vcu" ? "VCU" : "EXP"}] ${f.code} ${f.desc}` +
          (f.tEnd != null ? ` (duration ${fmtDur(f.tEnd - f.t)})` : "") +
          (f.open ? " (STILL ACTIVE)" : "") +
          (f.spd != null ? ` @${f.spd.toFixed(0)}mph` : "") +
          (f.up != null ? ` VCU-uptime ${Math.round(f.up)}s` : "")).join("\n"));
      /* clusters: events within 120s of each other often share one incident */
      const clusters = [];
      let cur = [winF[0]];
      for (let i = 1; i < winF.length; i++) {
        if (winF[i].t - cur[cur.length - 1].t <= 120000) cur.push(winF[i]);
        else { if (cur.length > 1) clusters.push(cur); cur = [winF[i]]; }
      }
      if (cur.length > 1) clusters.push(cur);
      if (clusters.length) L.push(`Fault clusters in window (events within 120s, likely one incident each): ` +
        clusters.slice(0, 8).map((c) => `${fmtFull(c[0].t)}: ${c.map((f) => f.code).join(" → ")}`).join(" | "));
    }

    /* selected event deep context */
    if (selSnapshot?.f) {
      const f = selSnapshot.f;
      let sel = `SELECTED EVENT: ${f.code} ${f.desc} at ${fmtFull(f.t)}${f.tEnd != null ? `, cleared after ${fmtDur(f.tEnd - f.t)}` : f.open ? ", STILL ACTIVE" : ""}${f.spd != null ? `, fault-log speed ${f.spd.toFixed(0)}mph` : ""}${f.up != null ? `, VCU uptime ${Math.round(f.up)}s${f.up < 120 ? " (recent VCU restart!)" : ""}` : ""}.`;
      if (!selSnapshot.outside && selSnapshot.row) {
        sel += ` Recorder signals at that instant (${(selSnapshot.dt / 1000).toFixed(3)}s sample offset): ` +
          selSnapshot.rows.filter((r) => r.v != null).slice(0, 24).map((r) => `${r.s}=${r.binary ? (r.v >= 0.5 ? "HIGH" : "LOW") : r.v}`).join(", ") + ".";
        /* neighbours: faults within ±10 min of the selected event */
        const near = mergedFaults.filter((x) => Math.abs(x.t - f.t) <= 600000 && x.id !== f.id).slice(0, 12);
        if (near.length) sel += ` Events within ±10min: ${near.map((x) => `${fmtFull(x.t)} ${x.code} ${x.desc}`).join("; ")}.`;
      } else sel += " (outside recorder coverage — no signal data at this instant).";
      L.push(sel);
    }
    if (offset1 !== 0 || offset2 !== 0) L.push(`Time-sync offsets applied to fault timestamps: V1 ${offset1}s, V2 ${offset2}s.`);
    return L.join("\n\n").slice(0, 12000);
  };

  const askAi = async (q) => {
    const question = (q ?? aiInput).trim();
    if (!question || aiBusy) return;
    setAiBusy(true); setAiError(null); setAiInput("");
    const history = [...aiMessages, { role: "user", content: question }];
    setAiMessages(history);
    const instruction = [
      "You are an expert rail vehicle reliability engineer analyzing light-rail vehicle (LRV) fleet diagnostics for a transit operator.",
      "",
      "DATA SOURCES:",
      "- TELOC event recorder: high-rate recordings of vehicle signals (speed, line/inverter voltage, brake cylinder pressures, ATP flags, door circuits). Millisecond resolution; records sparsely while the vehicle is idle.",
      "- Expert fault log: VCU (Vehicle Control Unit) fault events exported per car — code, description, sometimes speed/odometer/VCU-uptime at occurrence. 1-second timestamp resolution; events are instants.",
      "- VCU diagnostic memory dump: fault events with Appeared/Disappeared timestamps (durations). 'STILL ACTIVE' means the fault never cleared in the dump.",
      "",
      "DOMAIN NOTES:",
      "- ATP = Automatic Train Protection; an ATP safety brake application forces speed to zero. Safety brake events while already stationary usually reflect door/standstill logic rather than emergency braking.",
      "- BCU = Brake Control Unit, one per truck (Power Truck A/B, Center Truck C). 'Brakes cutout' = that truck's brakes isolated — degraded braking, often follows a BCU fault.",
      "- MVB/WTB = train communication buses to the VOBC; bus faults frequently cascade into consequential faults on other subsystems within seconds.",
      "- Low VCU uptime at a fault (<120s) means the VCU recently restarted — the fault may be a power-up artifact.",
      "- Door emergency-release and Passenger Emergency Stop events are usually operational/passenger actions, not equipment failures. CCTV/NVR faults are non-safety. Diagnostic-memory 'Environment' entries at fixed times daily are routine housekeeping.",
      "",
      "ANALYSIS APPROACH: distinguish root-cause from consequential faults using ordering, durations, and subsystem relationships; treat events in the same second or tight cluster as one incident; correlate fault instants with recorder signal states; flag repeating patterns and systemic issues. Be concise and technical, suitable for an 8D/RCA report. Base everything STRICTLY on the data below — explicitly say when something cannot be determined from it.",
      "",
      "DATA:",
      "",
      buildAiContext(),
    ].join("\n");

    /* ---- Ollama (local, free, data stays on your machine) ---- */
    if (aiProvider === "ollama") {
      try {
        const res = await fetch(`${ollamaUrl.replace(/\/$/, "")}/api/chat`, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            model: ollamaModel,
            stream: false,
            messages: [{ role: "system", content: instruction }, ...history],
          }),
        });
        const data = await res.json();
        if (data.error) throw new Error(data.error);
        const text = (data.message?.content || "").trim();
        setAiMessages([...history, { role: "assistant", content: text || "(no response)" }]);
      } catch (e) {
        const msg = String(e?.message || e);
        setAiError(/failed to fetch|networkerror|load failed/i.test(msg)
          ? "Couldn't reach Ollama at " + ollamaUrl + ". Check it's running (`ollama serve`) and that OLLAMA_ORIGINS allows this page's origin."
          : msg);
        setAiMessages(history.slice(0, -1));
        setAiInput(question);
      }
      setAiBusy(false);
      return;
    }

    const payload = JSON.stringify({
      model: "claude-sonnet-4-20250514",
      max_tokens: 1000,
      messages: [
        { role: "user", content: instruction },
        { role: "assistant", content: "Understood — I will analyze the provided LRV diagnostic data as a rail reliability engineer, based strictly on the data." },
        ...history,
      ],
    });
    let lastErr = null;
    for (let attempt = 0; attempt < 2; attempt++) {
      try {
        // Inside Claude.ai artifacts no key is needed (platform bridge).
        // Self-hosted (e.g. GitHub Pages): supply an API key — uses Anthropic's
        // CORS-enabled direct browser access. The key lives only in memory.
        const headers = aiKey.trim()
          ? { "Content-Type": "application/json", "x-api-key": aiKey.trim(), "anthropic-version": "2023-06-01", "anthropic-dangerous-direct-browser-access": "true" }
          : { "Content-Type": "application/json" };
        const res = await fetch("https://api.anthropic.com/v1/messages", {
          method: "POST",
          headers,
          body: payload,
        });
        const data = await res.json();
        if (data.error) throw new Error(data.error.message || "API error");
        const text = (data.content || []).map((c) => (c.type === "text" ? c.text : "")).join("\n").trim();
        setAiMessages([...history, { role: "assistant", content: text || "(no response)" }]);
        setAiBusy(false);
        return;
      } catch (e) {
        lastErr = e;
        // network-level failure: brief pause, then one retry
        if (attempt === 0 && /fetch|network/i.test(String(e?.message || e))) {
          await new Promise((r) => setTimeout(r, 1200));
          continue;
        }
        break;
      }
    }
    const msg = String(lastErr?.message || lastErr);
    setAiError(/failed to fetch|networkerror|load failed/i.test(msg)
      ? (aiKey.trim()
        ? "Couldn't reach the Anthropic API (network error, retried once). Check your connection and that the API key is valid, then try again."
        : "Couldn't reach the AI service. If this app is self-hosted (e.g. GitHub Pages), click 'API key' below and enter an Anthropic API key — the keyless mode only works inside Claude.ai.")
      : msg);
    if (!aiKey.trim() && /failed to fetch|networkerror|load failed/i.test(msg)) setAiKeyOpen(true);
    setAiMessages(history.slice(0, -1));
    setAiInput(question);
    setAiBusy(false);
  };

  /* ----- pan / zoom / scroll ----- */
  const chartWrapRef = useRef(null);
  const dragRef = useRef(null);
  const mmDragRef = useRef(null);

  /* ----- time-frame selection ----- */
  const [selectMode, setSelectMode] = useState(false); // next drag on chart selects a range
  const [band, setBand] = useState(null);              // live rubber-band {x1,x2} in px
  const bandRef = useRef(null);
  const [rangeOpen, setRangeOpen] = useState(false);
  const [rangeFrom, setRangeFrom] = useState("");
  const [rangeTo, setRangeTo] = useState("");
  const toLocalInput = (t) => {
    const d = new Date(t), p = (n) => String(n).padStart(2, "0");
    return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}T${p(d.getHours())}:${p(d.getMinutes())}:${p(d.getSeconds())}`;
  };
  const applyRange = () => {
    const a = new Date(rangeFrom).getTime(), b = new Date(rangeTo).getTime();
    if (isNaN(a) || isNaN(b) || b <= a) return;
    setDomain(clampDomain(a, b));
    setSelFault(null);
  };
  const applyDuration = (ms) => {
    if (!activeDomain) return;
    const end = activeDomain[1];
    setDomain(clampDomain(end - ms, end));
    setRangeFrom(toLocalInput(end - ms)); setRangeTo(toLocalInput(end));
  };

  const clampDomain = useCallback((a, b) => {
    const bounds = navFull;
    if (!bounds) return [a, b];
    const minSpan = 2000, full = bounds[1] - bounds[0];
    let span = Math.min(Math.max(b - a, minSpan), full);
    let na = a;
    if (na < bounds[0]) na = bounds[0];
    if (na + span > bounds[1]) na = bounds[1] - span;
    return [na, na + span];
  }, [navFull]);

  const panBy = (frac) => {
    if (!activeDomain) return;
    const v0 = tv(activeDomain[0]), v1 = tv(activeDomain[1]), s = v1 - v0;
    setDomain(clampDomain(tr(v0 + s * frac), tr(v1 + s * frac)));
  };
  const zoomBy = (factor, anchor = 0.5) => {
    if (!activeDomain) return;
    const v0 = tv(activeDomain[0]), v1 = tv(activeDomain[1]);
    const c = v0 + (v1 - v0) * anchor, s = (v1 - v0) * factor;
    setDomain(clampDomain(tr(c - s * anchor), tr(c + s * (1 - anchor))));
  };

  useEffect(() => {
    const el = chartWrapRef.current;
    if (!el) return;
    const onWheel = (e) => {
      e.preventDefault();
      const rect = el.getBoundingClientRect();
      const anchor = Math.min(Math.max((e.clientX - rect.left) / rect.width, 0), 1);
      zoomBy(e.deltaY > 0 ? 1.25 : 0.8, anchor);
    };
    el.addEventListener("wheel", onWheel, { passive: false });
    return () => el.removeEventListener("wheel", onWheel);
  });

  const chartPointerDown = (e) => {
    if (!chartWrapRef.current || !activeDomain) return;
    e.currentTarget.setPointerCapture?.(e.pointerId);
    const rect = chartWrapRef.current.getBoundingClientRect();
    if (selectMode || e.shiftKey) {
      bandRef.current = { startX: e.clientX, rect };
      setBand({ x1: e.clientX - rect.left, x2: e.clientX - rect.left });
      return;
    }
    dragRef.current = { x: e.clientX, vdom: [tv(activeDomain[0]), tv(activeDomain[1])], w: rect.width };
  };
  const chartPointerMove = (e) => {
    if (bandRef.current) {
      const { startX, rect } = bandRef.current;
      setBand({ x1: Math.min(startX, e.clientX) - rect.left, x2: Math.max(startX, e.clientX) - rect.left });
      return;
    }
    const d = dragRef.current;
    if (!d) return;
    const dvt = -((e.clientX - d.x) / d.w) * (d.vdom[1] - d.vdom[0]);
    setDomain(clampDomain(tr(d.vdom[0] + dvt), tr(d.vdom[1] + dvt)));
  };
  const chartPointerUp = () => {
    if (bandRef.current) {
      const { rect } = bandRef.current;
      bandRef.current = null;
      if (band && Math.abs(band.x2 - band.x1) > 8 && activeDomain) {
        const v0 = tv(activeDomain[0]), v1 = tv(activeDomain[1]), vspan = v1 - v0;
        setDomain(clampDomain(tr(v0 + (band.x1 / rect.width) * vspan), tr(v0 + (band.x2 / rect.width) * vspan)));
        setSelFault(null);
      }
      setBand(null);
      setSelectMode(false);
      return;
    }
    dragRef.current = null;
  };

  const mmTicks = useMemo(() => {
    if (!mergedFaults.length || !navFull) return [];
    const inNav = mergedFaults.filter((f) => f.t >= navFull[0] && f.t <= navFull[1]);
    const stride = Math.max(1, Math.ceil(inNav.length / 350));
    const span = navFull[1] - navFull[0] || 1;
    return inNav.filter((_, i) => i % stride === 0).map((f) => ({ frac: (f.t - navFull[0]) / span, vi: f.vi }));
  }, [mergedFaults, navFull]);

  const mmPointerDown = (e) => {
    if (!navFull || !activeDomain) return;
    const el = e.currentTarget;
    el.setPointerCapture?.(e.pointerId);
    const rect = el.getBoundingClientRect();
    const frac = (e.clientX - rect.left) / rect.width;
    const tFull = navFull[1] - navFull[0];
    const span = activeDomain[1] - activeDomain[0];
    const winA = (activeDomain[0] - navFull[0]) / tFull, winB = (activeDomain[1] - navFull[0]) / tFull;
    if (frac >= winA && frac <= winB) {
      mmDragRef.current = { x: e.clientX, dom: [...activeDomain], w: rect.width, tFull };
    } else {
      const c = navFull[0] + frac * tFull;
      const nd = clampDomain(c - span / 2, c + span / 2);
      setDomain(nd);
      mmDragRef.current = { x: e.clientX, dom: nd, w: rect.width, tFull };
    }
  };
  const mmPointerMove = (e) => {
    const d = mmDragRef.current;
    if (!d) return;
    const dt = ((e.clientX - d.x) / d.w) * d.tFull;
    setDomain(clampDomain(d.dom[0] + dt, d.dom[1] + dt));
  };
  const mmPointerUp = () => { mmDragRef.current = null; };

  const loadDemo = () => {
    const d1 = makeDemoVehicle("224", 0), d2 = makeDemoVehicle("211", 95);
    const mk = (rows, kind, extras) => {
      const det = analyzeRows(rows);
      return {
        fileName: kind === "rec" ? "demo_event_recorder.xlsx" : "demo_fault_log.xlsx",
        sheets: ["Sheet1"], sheet: "Sheet1", headerRow: 0, rows,
        cols: det.cols, numericCols: det.numericCols, textCols: det.textCols,
        timeCol: null, signals: kind === "rec" ? det.numericCols : [],
        ...extras,
      };
    };
    setRec1(mk(d1.rec, "rec", { tsCol: "Timestamp", codeCol: null, descCol: null, vehCol: null, vehPick: null }));
    setFault1(mk(d1.faults, "fault", { tsCol: "DateTime", codeCol: "CodeID", descCol: "Fault", vehCol: "CarNo", vehPick: null }));
    setRec2(mk(d2.rec, "rec", { tsCol: "Timestamp", codeCol: null, descCol: null, vehCol: null, vehPick: null }));
    setFault2(mk(d2.faults, "fault", { tsCol: "DateTime", codeCol: "CodeID", descCol: "Fault", vehCol: "CarNo", vehPick: null }));
    setName1("Car 224"); setName2("Car 211");
  };

  /* ---------- chart renderer (combined or per-vehicle) ---------- */
  const yDomain = normalize
    ? [0, 100]
    : [chartCfg.yMin !== "" && !isNaN(+chartCfg.yMin) ? +chartCfg.yMin : "auto", chartCfg.yMax !== "" && !isNaN(+chartCfg.yMax) ? +chartCfg.yMax : "auto"];

  const renderChart = (rows, seriesDefs, faults, height, dots = [], coverage = []) => {
    let prows = tmap ? rows.map((r) => ({ ...r, tReal: r.t, t: tv(r.t) })) : rows;
    const dom = tmap ? vActive : activeDomain;
    const showOverlay = chartCfg.showFaults && chartCfg.faultStyle !== "markers";
    const showMarkers = chartCfg.showFaults && chartCfg.faultStyle !== "overlay";
    // graph shows only selected fault codes; the clicked/selected fault is always highlighted
    const gFaults = totalSel ? faults.filter(graphShown) : [];
    // Expert/fault data as a discrete channel laid on top of the recorder data:
    // a logic-style band near the top of the plot (low = no fault, high pulse = fault logged)
    let pulseSeries = [];
    if (showOverlay && gFaults.length && dom) {
      let H = 100;
      if (!normalize) {
        let mx = -Infinity;
        for (const r of rows) for (const d of seriesDefs) { const v = r[d.dataKey]; if (v != null && v > mx) mx = v; }
        H = mx > -Infinity ? mx : 1;
      }
      const bands = { 1: [0.875, 0.985], 2: [0.73, 0.84] }; // per-vehicle band as fraction of scale
      const w = Math.max((dom[1] - dom[0]) / 400, 20); // pulse half-width in axis units
      const groups = {};
      for (const f of gFaults.slice(0, 400)) (groups[f.vi] = groups[f.vi] || []).push(f);
      const extra = [];
      const X = (t) => (tmap ? tv(t) : t);
      pulseSeries = Object.entries(groups).map(([vi, fs]) => {
        const key = v2Active ? `⚠ FAULTS V${vi}` : "⚠ FAULTS";
        const lo = +(H * bands[vi][0]).toFixed(3), hi = +(H * bands[vi][1]).toFixed(3);
        // spans: Expert events are instants; VCU events run Appeared→Disappeared (open = still active)
        const ivs = fs.map((f) => {
          const t0 = X(f.t);
          const t1 = f.open ? dom[1] : (f.tEnd != null ? X(f.tEnd) : t0);
          return [t0 - w, Math.max(t1, t0) + w];
        }).sort((p, q) => p[0] - q[0]);
        const merged = [];
        for (const iv of ivs) {
          const L = merged[merged.length - 1];
          if (L && iv[0] <= L[1]) L[1] = Math.max(L[1], iv[1]);
          else merged.push([...iv]);
        }
        extra.push({ t: dom[0], [key]: lo });
        for (const [s0, s1] of merged) {
          extra.push({ t: s0, [key]: lo }, { t: s0 + 1, [key]: hi }, { t: s1 - 1, [key]: hi }, { t: s1, [key]: lo });
        }
        extra.push({ t: dom[1], [key]: lo });
        return { key, color: VEH[+vi].fault };
      });
      prows = [...prows, ...extra].sort((a, b) => a.t - b.t);
    }
    return (
    <ResponsiveContainer width="100%" height={height}>
      <ComposedChart data={prows} margin={{ top: 10, right: 20, bottom: 5, left: 0 }}>
        {chartCfg.showGrid && <CartesianGrid stroke={C.faint} strokeDasharray="2 6" vertical={false} />}
        <XAxis dataKey="t" type="number" domain={dom} tickFormatter={(t) => fmtTime(tr(t), spanMs)} stroke={C.dim} tick={{ fontSize: 10, fontFamily: "'IBM Plex Mono', monospace" }} tickCount={7} allowDataOverflow />
        <YAxis stroke={C.dim} tick={{ fontSize: 10, fontFamily: "'IBM Plex Mono', monospace" }} width={52} domain={yDomain} allowDataOverflow={!normalize} padding={{ top: 4, bottom: 8 }} />
        <Tooltip content={<DarkTooltip labelMap={tr} />} />
        {coverage.map((c, i) => {
          if (!c.span || !activeDomain) return null;
          const x1 = Math.max(c.span[0], activeDomain[0]), x2 = Math.min(c.span[1], activeDomain[1]);
          if (x2 <= x1) return null;
          return <ReferenceArea key={`cov${i}`} x1={tv(x1)} x2={tv(x2)} fill={c.color} fillOpacity={0.05} stroke={c.color} strokeOpacity={0.25} strokeDasharray="2 4" />;
        })}
        {tmap && tmap.gapsV.map((g, i) => {
          if (!dom || g.v1 < dom[0] || g.v0 > dom[1]) return null;
          return <ReferenceArea key={`gap${i}`} x1={Math.max(g.v0, dom[0])} x2={Math.min(g.v1, dom[1])} fill="#9aa3ad" fillOpacity={0.18} stroke="#9aa3ad" strokeOpacity={0.5} strokeDasharray="2 2" />;
        })}
        {chartCfg.showLegend && <Legend wrapperStyle={{ fontSize: 11, fontFamily: "'IBM Plex Mono', monospace" }} />}
        {seriesDefs.map((d) => {
          const s = resolveSeries(d);
          return s.area ? (
            <Area key={d.cfgKey} dataKey={d.dataKey} name={s.label} type="monotone" stroke={s.color} fill={s.color} fillOpacity={0.15} strokeWidth={s.width} dot={false} isAnimationActive={false} connectNulls />
          ) : (
            <Line key={d.cfgKey} dataKey={d.dataKey} name={s.label} type={s.curve} stroke={s.color} strokeDasharray={s.dash} strokeWidth={s.width} dot={s.dots ? { r: 2, strokeWidth: 0, fill: s.color } : false} isAnimationActive={false} connectNulls />
          );
        })}
        {pulseSeries.map((p) => (
          <Line key={p.key} dataKey={p.key} name={p.key} type="linear" stroke={p.color} strokeWidth={1.4} dot={false} isAnimationActive={false} connectNulls />
        ))}
        {chartCfg.showFaults && faults.slice(0, 250).map((f) => {
          const isSel = `${f.vi}-${f.id}` === selFault;
          if (isSel ? false : !(showMarkers && totalSel && graphShown(f))) return null;
          return (
            <ReferenceLine key={`${f.vi}-${f.id}`} x={tv(f.t)}
              stroke={isSel ? C.amber : VEH[f.vi].fault}
              strokeWidth={isSel ? 2 : 1}
              strokeDasharray={isSel ? "0" : "4 3"}
              strokeOpacity={isSel ? 1 : 0.6} />
          );
        })}
        {dots.map((d, i) => (
          <ReferenceDot key={i} x={tv(d.x)} y={d.y} r={4.5} fill={d.color} stroke="#ffffff" strokeWidth={1.5} ifOverflow="discard" isFront />
        ))}
      </ComposedChart>
    </ResponsiveContainer>
    );
  };

  /* ---------- faults rendered as pulse-train signal lanes ---------- */
  const renderFaultLanes = (vi, vName) => {
    const src = vi === 1 ? faultData1 : faultData2;
    if (!src?.length || !activeDomain) return null;
    const [a, b] = activeDomain;
    const slack = (b - a) * 0.01;
    const inWin = (arr) => arr.filter((f) => {
      const end = f.open ? Infinity : (f.tEnd ?? f.t);
      return f.t <= b + slack && end >= a - slack;
    });
    const toEvents = (arr) => arr.map((f) => (tmap
      ? { t: tv(f.t), tEnd: f.tEnd != null ? tv(f.tEnd) : null, open: f.open }
      : { t: f.t, tEnd: f.tEnd, open: f.open }));
    // lanes only for explicitly selected codes (the ⚠ channel in the chart covers "all")
    const lanes = [];
    const sel = [
      ...visExpert.map((code) => ({ code, src: "expert" })),
      ...visVcu.map((code) => ({ code, src: "vcu" })),
    ].slice(0, 400);
    // pre-group events by source+code once, so hundreds of lanes stay fast while panning
    const byKey = {};
    for (const f of src) { const k = (f.src || "expert") + "-" + f.code; (byKey[k] = byKey[k] || []).push(f); }
    sel.forEach(({ code, src: fsrc }, i) => {
      const pool = fsrc === "vcu" ? vcuCodes : expertCodes;
      const meta = pool.find((p) => p.code === code);
      lanes.push({
        key: `${fsrc}-${code}`,
        label: `${code}${meta?.desc ? ` ${meta.desc}` : ""}${fsrc === "vcu" ? " [VCU]" : ""} · ${vName}`,
        color: SIG_COLORS[(i + 4) % SIG_COLORS.length],
        events: toEvents(inWin(byKey[`${fsrc}-${code}`] || [])),
      });
    });
    if (!lanes.length) return null;
    const scrollable = lanes.length > 10;
    return (
      <div style={{ border: `1px solid ${C.panelEdge}`, borderTop: "none", overflow: "hidden" }}>
        {scrollable && (
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "4px 10px", background: "#f3f3f0", borderBottom: `1px solid ${C.panelEdge}`, fontFamily: "'Barlow Condensed', sans-serif", fontWeight: 700, fontSize: 11, letterSpacing: "0.1em", textTransform: "uppercase", color: C.dim }}>
            <span>Fault lanes · {vName} ({lanes.length})</span>
            <span style={{ fontFamily: "'IBM Plex Mono', monospace", textTransform: "none", letterSpacing: 0, fontWeight: 400, fontSize: 9 }}>scroll for more ↓</span>
          </div>
        )}
        <div style={{ maxHeight: scrollable ? 320 : undefined, overflowY: scrollable ? "auto" : undefined }}>
          {lanes.map((l) => <FaultLane key={l.key} label={l.label} color={l.color} tagColor={v2Active ? VEH[vi].tag : "transparent"} events={l.events} domain={tmap ? vActive : activeDomain} />)}
        </div>
      </div>
    );
  };

  /* ---------- TELOC-style digital lanes block ---------- */
  const renderLanes = (defs, laneRows, faults, vi) => defs.length > 0 && (
    <div style={{ border: `1px solid ${C.panelEdge}`, borderTop: "none", borderRadius: "0 0 4px 4px", overflow: "hidden", marginBottom: 4 }}>
      {defs.map((d) => {
        const s = resolveSeries(d);
        const pts = tmap ? lanePts(laneRows, d.sig).map((p) => ({ ...p, t: tv(p.t) })) : lanePts(laneRows, d.sig);
        const lfaults = (chartCfg.showFaults ? faults : []).map((f) => (tmap ? { ...f, t: tv(f.t) } : f));
        return <DigitalLane key={d.cfgKey} label={s.label} color={s.color} tagColor={v2Active ? VEH[vi].tag : "transparent"} pts={pts} domain={tmap ? vActive : activeDomain} faults={lfaults} selFault={selFault} />;
      })}
    </div>
  );

  /* ---------- export current graph(s) as PNG ---------- */
  const exportPNG = () => {
    const svgs = chartWrapRef.current?.querySelectorAll("svg");
    if (!svgs?.length) return;
    const pad = 16, scale = 2;
    const dims = [...svgs].map((s) => s.getBoundingClientRect());
    const W = Math.max(...dims.map((d) => d.width)) + pad * 2;
    const titleH = chartCfg.title ? 42 : 0;
    const H = dims.reduce((a, d) => a + d.height, 0) + pad * (svgs.length + 1) + titleH;
    const canvas = document.createElement("canvas");
    canvas.width = W * scale; canvas.height = H * scale;
    const ctx = canvas.getContext("2d");
    ctx.scale(scale, scale);
    ctx.fillStyle = "#ffffff"; ctx.fillRect(0, 0, W, H);
    if (chartCfg.title) {
      ctx.fillStyle = C.ink; ctx.font = "700 18px 'Barlow Condensed', sans-serif";
      ctx.fillText(chartCfg.title.toUpperCase(), pad, pad + 16);
    }
    let y = pad + titleH, loaded = 0;
    [...svgs].forEach((svg, i) => {
      const laneParent = svg.closest(".lane-row");
      const xml = new XMLSerializer().serializeToString(svg);
      const img = new Image();
      const yPos = y; y += dims[i].height + (laneParent ? 2 : pad);
      const xPos = laneParent ? pad + 190 : pad;
      if (laneParent) {
        ctx.fillStyle = laneParent.dataset.color || C.ink;
        ctx.font = "10px 'IBM Plex Mono', monospace";
        ctx.fillText((laneParent.dataset.label || "").slice(0, 30), pad + 4, yPos + dims[i].height / 2 + 3);
      }
      img.onload = () => {
        ctx.drawImage(img, xPos, yPos, laneParent ? Math.max(dims[i].width - 190, 50) : dims[i].width, dims[i].height);
        if (++loaded === svgs.length) {
          const a = document.createElement("a");
          a.download = (chartCfg.title || "fleet-graph").replace(/\s+/g, "_") + ".png";
          a.href = canvas.toDataURL("image/png");
          a.click();
        }
      };
      img.src = "data:image/svg+xml;charset=utf-8," + encodeURIComponent(xml);
    });
  };

  return (
    <div style={{ minHeight: "100vh", background: C.bg, color: C.ink, fontFamily: "'IBM Plex Mono', monospace", backgroundImage: "radial-gradient(circle at 20% 0%, rgba(0,153,153,0.05), transparent 40%), repeating-linear-gradient(0deg, transparent, transparent 3px, rgba(20,30,40,0.012) 3px, rgba(20,30,40,0.012) 4px)" }}>
      <style>{FONTS}</style>

      <div style={{ borderBottom: `1px solid ${C.panelEdge}`, padding: "18px 28px", display: "flex", alignItems: "center", justifyContent: "space-between", flexWrap: "wrap", gap: 12 }}>
        <div>
          <div style={{ fontFamily: "'Barlow Condensed', sans-serif", fontWeight: 700, fontSize: 26, letterSpacing: "0.16em", textTransform: "uppercase" }}>
            <span style={{ color: C.amber }}>▮</span> Fleet Data Analyzer
          </div>
          <div style={{ fontSize: 11, color: C.dim, letterSpacing: "0.05em" }}>Event recorder × fault log correlation · up to 2 vehicles</div>
        </div>
        <div style={{ display: "flex", gap: 10 }}>
          {stage === "analyze" && <Btn small onClick={() => { setStage("setup"); setDomain(null); setSelFault(null); clearFaultSel(); }}>⟵ Data setup</Btn>}
          {stage === "setup" && <Btn small onClick={loadDemo}>Load demo (2 vehicles)</Btn>}
        </div>
      </div>

      {/* ===================== SETUP ===================== */}
      {stage === "setup" && (
        <div style={{ padding: 28, maxWidth: 1200, margin: "0 auto" }}>
          {[
            { n: 1, name: name1, setName: setName1, rec: rec1, setRec: setRec1, fault: fault1, setFault: setFault1, vcu: vcu1, setVcu: setVcu1, required: true },
            { n: 2, name: name2, setName: setName2, rec: rec2, setRec: setRec2, fault: fault2, setFault: setFault2, vcu: vcu2, setVcu: setVcu2, required: false },
          ].map((v) => (
            <div key={v.n} style={{ marginBottom: 26 }}>
              <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 12 }}>
                <span style={{ width: 10, height: 10, borderRadius: 2, background: VEH[v.n].tag }} />
                <input value={v.name} onChange={(e) => v.setName(e.target.value)}
                  style={{ fontFamily: "'Barlow Condensed', sans-serif", fontWeight: 700, fontSize: 20, letterSpacing: "0.12em", textTransform: "uppercase", color: VEH[v.n].tag, background: "transparent", border: "none", borderBottom: `1px dashed ${C.faint}`, outline: "none", width: 220 }} />
                <span style={{ fontSize: 11, color: C.dim }}>{v.required ? "required" : "optional — leave empty for single-vehicle analysis"}</span>
              </div>
              <div style={{ display: "flex", gap: 18, flexWrap: "wrap" }}>
                <FilePanel title="Event Recorder" accent={VEH[v.n].tag} data={v.rec} setData={v.setRec} kind="rec" />
                <FilePanel title="Expert Data Logs" accent={VEH[v.n].fault} data={v.fault} setData={v.setFault} kind="fault" />
                <VcuPanel accent={C.violet} data={v.vcu} setData={v.setVcu} />
              </div>
            </div>
          ))}
          <div style={{ display: "flex", alignItems: "center", gap: 16, flexWrap: "wrap" }}>
            <Btn primary disabled={!ready || !rec1?.signals?.length} onClick={() => setStage("analyze")}>Analyze ▶</Btn>
            {!ready && <span style={{ color: C.dim, fontSize: 12 }}>Vehicle 1 needs the Event Recorder plus at least one fault source (Expert Data or VCU Data Logs). If one export contains multiple cars, use "Filter to vehicle" to split it.</span>}
            {ready && !rec1?.signals?.length ? <span style={{ color: C.red, fontSize: 12 }}>Select at least one signal channel for Vehicle 1.</span> : null}
          </div>
        </div>
      )}

      {/* ===================== ANALYZE ===================== */}
      {stage === "analyze" && ready && (
        <div style={{ padding: "20px 28px" }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-end", flexWrap: "wrap", gap: 10 }}>
            <div style={{ display: "flex", gap: 4 }}>
              {[["timeline", "Timeline Correlation"], ["stats", "Fault Statistics"]].map(([k, label]) => (
                <div key={k} onClick={() => setTab(k)}
                  style={{
                    fontFamily: "'Barlow Condensed', sans-serif", fontWeight: 700, fontSize: 15, letterSpacing: "0.12em", textTransform: "uppercase",
                    padding: "9px 20px", cursor: "pointer", borderRadius: "4px 4px 0 0",
                    color: tab === k ? C.amber : C.dim, background: tab === k ? C.panel : "transparent",
                    border: `1px solid ${tab === k ? C.panelEdge : "transparent"}`, borderBottom: "none",
                  }}>{label}</div>
              ))}
            </div>

            <div style={{ display: "flex", gap: 12, alignItems: "center", paddingBottom: 6, flexWrap: "wrap" }}>
              {v2Active && tab === "timeline" && (
                <div style={{ display: "flex", gap: 6, alignItems: "center" }}>
                  <span style={{ fontSize: 10, color: C.dim, letterSpacing: "0.1em", textTransform: "uppercase", fontFamily: "'Barlow Condensed', sans-serif", fontWeight: 600 }}>Layout:</span>
                  <Chip active={viewMode === "combined"} onClick={() => setViewMode("combined")}>1 graph</Chip>
                  <Chip active={viewMode === "separate"} onClick={() => setViewMode("separate")}>2 graphs</Chip>
                </div>
              )}
              {expertCodes.length > 0 && <FaultLaneDropdown allCodes={expertCodes} selected={visExpert} onChange={setVisExpert} label="Expert faults" hint="Pick codes to draw them on the graph (⚠ channel + a lane per code). Nothing selected = clean graph; all faults stay in the list." />}
              {vcuCodes.length > 0 && <FaultLaneDropdown allCodes={vcuCodes} selected={visVcu} onChange={setVisVcu} label="VCU faults" color={C.violet} hint="Pick codes to draw them on the graph as duration spans (⚠ channel + a lane per code). Nothing selected = clean graph." />}
              {recData1?.length > 0 && rec1 && (
                <SignalDropdown label={v2Active ? `${name1} sig` : "Signals"} color={VEH[1].tag} allSignals={rec1.numericCols} selected={rec1.signals} onChange={(s) => setRec1({ ...rec1, signals: s })} />
              )}
              {v2Active && recData2?.length > 0 && rec2 && (
                <SignalDropdown label={`${name2} sig`} color={VEH[2].tag} allSignals={rec2.numericCols} selected={rec2.signals} onChange={(s) => setRec2({ ...rec2, signals: s })} />
              )}
            </div>
          </div>

          {/* ---------- TIMELINE ---------- */}
          {tab === "timeline" && (
            <div style={{ display: "flex", flexDirection: "column", gap: 18, marginTop: 14 }}>
              <div style={{ width: "100%", minWidth: 0 }}>
                <div style={{ background: C.panel, border: `1px solid ${C.panelEdge}`, borderRadius: 6, padding: 18 }}>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 10, marginBottom: 8 }}>
                    <div style={{ fontSize: 11, color: C.dim }}>
                      {fmtFull(activeDomain[0])} → {fmtFull(activeDomain[1])} · {visibleFaults1.length + visibleFaults2.length} fault{visibleFaults1.length + visibleFaults2.length !== 1 ? "s" : ""} in window
                      {totalSel > 0 && <span style={{ color: C.amber }}> · graphing {totalSel} fault code{totalSel > 1 ? "s" : ""} <span style={{ cursor: "pointer" }} onClick={clearFaultSel}>✕</span></span>}
                    </div>
                    <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                      <Btn small onClick={() => panBy(-0.5)}>◀</Btn>
                      <Btn small onClick={() => panBy(0.5)}>▶</Btn>
                      <Btn small onClick={() => zoomBy(0.5)}>＋</Btn>
                      <Btn small onClick={() => zoomBy(2)}>−</Btn>
                      <Btn small onClick={() => setNormalize(!normalize)}>{normalize ? "Raw values" : "Normalize %"}</Btn>
                      <Btn small onClick={() => setContinuous(!continuous)}>{continuous ? "Real time" : "Continuous"}</Btn>
                      <Btn small onClick={() => { setDomain(recSpan ? [...recSpan] : null); setSelFault(null); }}>Full recording</Btn>
                      <Btn small onClick={() => {
                        if (!rangeOpen && activeDomain) { setRangeFrom(toLocalInput(activeDomain[0])); setRangeTo(toLocalInput(activeDomain[1])); }
                        setRangeOpen(!rangeOpen);
                      }}>⏱ Time range</Btn>
                      <Btn small onClick={() => setSelectMode(!selectMode)}>{selectMode ? "✦ drag on chart…" : "⛶ Select on chart"}</Btn>
                      <Btn small onClick={() => setSyncOpen(!syncOpen)}>{syncOpen ? "✓ Sync" : "⟲ Time sync"}</Btn>
                      <Btn small onClick={() => setEditOpen(!editOpen)}>{editOpen ? "✓ Done" : "✎ Edit graph"}</Btn>
                      <Btn small onClick={exportPNG}>⤓ PNG</Btn>
                    </div>
                  </div>

                  {/* ---------- time sync panel ---------- */}
                  {syncOpen && (
                    <div style={{ background: "#f3f3f0", border: `1px solid ${C.faint}`, borderRadius: 6, padding: 12, marginBottom: 12 }}>
                      <div style={{ fontSize: 11, color: C.ink, marginBottom: 10, fontFamily: "'IBM Plex Mono', monospace" }}>
                        Fault-log ↔ recorder time alignment. Both files are parsed at full precision (recorder: milliseconds; fault log: 1-second resolution, so ±0.5 s is inherent).
                        If the VCU clock and recorder clock drift, shift the fault timestamps until a fault pulse lines up with its recorder flag edge — positive = faults move later.
                      </div>
                      {[
                        { n: 1, name: name1, off: offset1, setOff: setOffset1, has: !!faultData1?.length },
                        { n: 2, name: name2, off: offset2, setOff: setOffset2, has: !!faultData2?.length },
                      ].filter((v) => v.has).map((v) => (
                        <div key={v.n} style={{ display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap", marginBottom: 6 }}>
                          <span style={{ fontFamily: "'Barlow Condensed', sans-serif", fontWeight: 700, fontSize: 13, letterSpacing: "0.1em", textTransform: "uppercase", color: VEH[v.n].tag, width: 110 }}>{v.name}</span>
                          {[-60, -10, -1, -0.5].map((d) => <Chip key={d} onClick={() => v.setOff(+(v.off + d).toFixed(1))}>{d}s</Chip>)}
                          <input type="number" step="0.1" value={v.off} onChange={(e) => v.setOff(e.target.value === "" ? 0 : +e.target.value)}
                            style={{ width: 90, background: "#fff", color: v.off !== 0 ? C.amber : C.ink, fontWeight: v.off !== 0 ? 600 : 400, border: `1px solid ${v.off !== 0 ? C.amber : C.faint}`, borderRadius: 4, padding: "6px 9px", fontFamily: "'IBM Plex Mono', monospace", fontSize: 12, outline: "none", textAlign: "center" }} />
                          <span style={{ fontSize: 10, color: C.dim }}>sec</span>
                          {[0.5, 1, 10, 60].map((d) => <Chip key={d} onClick={() => v.setOff(+(v.off + d).toFixed(1))}>+{d}s</Chip>)}
                          <Chip onClick={() => v.setOff(0)} color={C.red}>reset</Chip>
                          {v.off !== 0 && <span style={{ fontSize: 10, color: C.amber, fontFamily: "'IBM Plex Mono', monospace" }}>faults shifted {v.off > 0 ? "+" : ""}{v.off}s</span>}
                        </div>
                      ))}
                      <div style={{ fontSize: 10, color: C.dim, marginTop: 4 }}>
                        Tip: add a fault code as a lane, plot its matching recorder flag (e.g. 740 vs ATP_SafetyBrake), zoom to one event with Shift+drag, and nudge until the edges align. The offset applies everywhere — markers, lanes, list, and stats.
                      </div>
                    </div>
                  )}

                  {/* ---------- time range panel ---------- */}
                  {rangeOpen && (
                    <div style={{ background: "#f3f3f0", border: `1px solid ${C.faint}`, borderRadius: 6, padding: 12, marginBottom: 12, display: "flex", gap: 12, flexWrap: "wrap", alignItems: "flex-end" }}>
                      <div>
                        <Label>From</Label>
                        <input type="datetime-local" step="1" value={rangeFrom} onChange={(e) => setRangeFrom(e.target.value)}
                          style={{ background: "#fff", color: C.ink, border: `1px solid ${C.faint}`, borderRadius: 4, padding: "6px 9px", fontFamily: "'IBM Plex Mono', monospace", fontSize: 12, outline: "none" }} />
                      </div>
                      <div>
                        <Label>To</Label>
                        <input type="datetime-local" step="1" value={rangeTo} onChange={(e) => setRangeTo(e.target.value)}
                          style={{ background: "#fff", color: C.ink, border: `1px solid ${C.faint}`, borderRadius: 4, padding: "6px 9px", fontFamily: "'IBM Plex Mono', monospace", fontSize: 12, outline: "none" }} />
                      </div>
                      <Btn small primary onClick={applyRange}>Apply</Btn>
                      <div style={{ display: "flex", gap: 6, alignItems: "center", paddingBottom: 2 }}>
                        <span style={{ fontSize: 10, color: C.dim, letterSpacing: "0.1em", textTransform: "uppercase", fontFamily: "'Barlow Condensed', sans-serif", fontWeight: 600 }}>Quick:</span>
                        {[["1 min", 60000], ["5 min", 300000], ["15 min", 900000], ["1 hr", 3600000], ["6 hr", 21600000], ["24 hr", 86400000]].map(([lbl, ms]) => (
                          <Chip key={lbl} onClick={() => applyDuration(ms)}>{lbl}</Chip>
                        ))}
                      </div>
                      <span style={{ fontSize: 10, color: C.dim, paddingBottom: 6 }}>tip: Shift+drag on the chart also selects a range</span>
                    </div>
                  )}

                  {/* ---------- edit panel ---------- */}
                  {editOpen && (
                    <div style={{ background: "#f3f3f0", border: `1px solid ${C.faint}`, borderRadius: 6, padding: 14, marginBottom: 12 }}>
                      <div style={{ display: "flex", gap: 14, flexWrap: "wrap", alignItems: "flex-end", marginBottom: 12 }}>
                        <div style={{ flex: "2 1 220px" }}>
                          <Label>Graph title</Label>
                          <input value={chartCfg.title} onChange={(e) => setChartCfg({ ...chartCfg, title: e.target.value })} placeholder="e.g. Car 224 — ATP Safety Brake events vs speed"
                            style={{ width: "100%", boxSizing: "border-box", background: "#fff", color: C.ink, border: `1px solid ${C.faint}`, borderRadius: 4, padding: "7px 10px", fontFamily: "'IBM Plex Mono', monospace", fontSize: 12, outline: "none" }} />
                        </div>
                        <div style={{ flex: "0 0 90px" }}>
                          <Label>Y min</Label>
                          <input value={chartCfg.yMin} disabled={normalize} onChange={(e) => setChartCfg({ ...chartCfg, yMin: e.target.value })} placeholder="auto"
                            style={{ width: "100%", boxSizing: "border-box", background: normalize ? "#e8e8e3" : "#fff", color: C.ink, border: `1px solid ${C.faint}`, borderRadius: 4, padding: "7px 10px", fontFamily: "'IBM Plex Mono', monospace", fontSize: 12, outline: "none" }} />
                        </div>
                        <div style={{ flex: "0 0 90px" }}>
                          <Label>Y max</Label>
                          <input value={chartCfg.yMax} disabled={normalize} onChange={(e) => setChartCfg({ ...chartCfg, yMax: e.target.value })} placeholder="auto"
                            style={{ width: "100%", boxSizing: "border-box", background: normalize ? "#e8e8e3" : "#fff", color: C.ink, border: `1px solid ${C.faint}`, borderRadius: 4, padding: "7px 10px", fontFamily: "'IBM Plex Mono', monospace", fontSize: 12, outline: "none" }} />
                        </div>
                        <div style={{ flex: "0 0 150px" }}>
                          <Label>Fault display</Label>
                          <select value={chartCfg.faultStyle} onChange={(e) => setChartCfg({ ...chartCfg, faultStyle: e.target.value })}
                            style={{ width: "100%", background: "#fff", color: C.ink, border: `1px solid ${C.faint}`, borderRadius: 4, padding: "7px 8px", fontFamily: "'IBM Plex Mono', monospace", fontSize: 12, outline: "none" }}>
                            <option value="overlay">overlay signal</option>
                            <option value="markers">marker lines</option>
                            <option value="both">both</option>
                          </select>
                        </div>
                        <div style={{ display: "flex", gap: 12, alignItems: "center", paddingBottom: 8 }}>
                          {[["showLegend", "Legend"], ["showGrid", "Grid"], ["showFaults", "Faults"]].map(([k, lbl]) => (
                            <label key={k} style={{ display: "flex", alignItems: "center", gap: 5, fontSize: 11, color: C.ink, cursor: "pointer", fontFamily: "'IBM Plex Mono', monospace" }}>
                              <input type="checkbox" checked={chartCfg[k]} onChange={(e) => setChartCfg({ ...chartCfg, [k]: e.target.checked })} style={{ accentColor: C.amber }} />{lbl}
                            </label>
                          ))}
                          <Btn small onClick={() => { setSeriesCfg({}); setChartCfg({ title: "", showLegend: true, showGrid: true, showFaults: true, yMin: "", yMax: "", faultStyle: "overlay" }); }}>Reset all</Btn>
                        </div>
                      </div>
                      {normalize && (chartCfg.yMin !== "" || chartCfg.yMax !== "") && <div style={{ fontSize: 10, color: C.dim, marginBottom: 8 }}>Y min/max apply in Raw values mode (normalized view is fixed 0–100%).</div>}
                      <div style={{ display: "grid", gridTemplateColumns: "minmax(160px,2fr) 60px 100px 110px 70px 60px", gap: 8, alignItems: "center", fontFamily: "'IBM Plex Mono', monospace", fontSize: 11 }}>
                        <div style={{ color: C.dim, textTransform: "uppercase", fontSize: 9, letterSpacing: "0.1em" }}>Series label</div>
                        <div style={{ color: C.dim, textTransform: "uppercase", fontSize: 9, letterSpacing: "0.1em" }}>Color</div>
                        <div style={{ color: C.dim, textTransform: "uppercase", fontSize: 9, letterSpacing: "0.1em" }}>Line style</div>
                        <div style={{ color: C.dim, textTransform: "uppercase", fontSize: 9, letterSpacing: "0.1em" }}>Shape</div>
                        <div style={{ color: C.dim, textTransform: "uppercase", fontSize: 9, letterSpacing: "0.1em" }}>Width</div>
                        <div style={{ color: C.dim, textTransform: "uppercase", fontSize: 9, letterSpacing: "0.1em" }}>Points</div>
                        {editSeries.map((d) => {
                          const o = seriesCfg[d.cfgKey] || {};
                          const cur = resolveSeries(d);
                          return (
                            <React.Fragment key={d.cfgKey}>
                              <input value={o.label ?? d.cfgKey} onChange={(e) => updateSeries(d.cfgKey, { label: e.target.value })}
                                style={{ background: "#fff", color: C.ink, border: `1px solid ${C.faint}`, borderRadius: 3, padding: "5px 8px", fontFamily: "inherit", fontSize: 11, outline: "none", minWidth: 0 }} />
                              <input type="color" value={cur.color} onChange={(e) => updateSeries(d.cfgKey, { color: e.target.value })}
                                style={{ width: 36, height: 26, padding: 0, border: `1px solid ${C.faint}`, borderRadius: 3, background: "#fff", cursor: "pointer" }} />
                              <select value={o.style || (d.defDash ? "dashed" : "solid")} onChange={(e) => updateSeries(d.cfgKey, { style: e.target.value })}
                                style={{ background: "#fff", color: C.ink, border: `1px solid ${C.faint}`, borderRadius: 3, padding: "5px 6px", fontFamily: "inherit", fontSize: 11, outline: "none" }}>
                                {["solid", "dashed", "dotted"].map((s) => <option key={s} value={s}>{s}</option>)}
                              </select>
                              <select value={o.curve || "linear"} onChange={(e) => updateSeries(d.cfgKey, { curve: e.target.value })}
                                style={{ background: "#fff", color: C.ink, border: `1px solid ${C.faint}`, borderRadius: 3, padding: "5px 6px", fontFamily: "inherit", fontSize: 11, outline: "none" }}>
                                <option value="linear">line</option>
                                <option value="smooth">smooth</option>
                                <option value="step">step</option>
                                <option value="area">area</option>
                              </select>
                              <select value={String(o.width || 1.6)} onChange={(e) => updateSeries(d.cfgKey, { width: e.target.value })}
                                style={{ background: "#fff", color: C.ink, border: `1px solid ${C.faint}`, borderRadius: 3, padding: "5px 6px", fontFamily: "inherit", fontSize: 11, outline: "none" }}>
                                {["1", "1.6", "2.5", "3.5"].map((w) => <option key={w} value={w}>{w}px</option>)}
                              </select>
                              <input type="checkbox" checked={!!o.dots} onChange={(e) => updateSeries(d.cfgKey, { dots: e.target.checked })} style={{ accentColor: cur.color, justifySelf: "start", cursor: "pointer" }} />
                            </React.Fragment>
                          );
                        })}
                      </div>
                    </div>
                  )}

                  {chartCfg.title && (
                    <div style={{ textAlign: "center", fontFamily: "'Barlow Condensed', sans-serif", fontWeight: 700, fontSize: 17, letterSpacing: "0.1em", textTransform: "uppercase", color: C.ink, marginBottom: 4 }}>{chartCfg.title}</div>
                  )}

                  {/* warn when recorder data is a sliver of the current window */}
                  {(() => {
                    if (!activeDomain || !recSpan) return null;
                    const win = activeDomain[1] - activeDomain[0];
                    const ov = Math.max(0, Math.min(recSpan[1], activeDomain[1]) - Math.max(recSpan[0], activeDomain[0]));
                    const frac = ov / win;
                    if (frac >= 0.15) return null;
                    return (
                      <div style={{ display: "flex", alignItems: "center", gap: 10, background: "#fdf6e7", border: `1px solid ${C.amber}`, borderRadius: 4, padding: "7px 12px", marginBottom: 8, fontSize: 11, fontFamily: "'IBM Plex Mono', monospace", color: C.ink }}>
                        <span style={{ color: C.amber, fontWeight: 700 }}>⚠</span>
                        {ov === 0
                          ? <span>No recorder data in this window — only fault markers are available here. Recorder coverage: {fmtFull(recSpan[0])} → {fmtFull(recSpan[1])}.</span>
                          : <span>Recorder data covers only {(frac * 100).toFixed(1)}% of this window (shaded band) — signals are compressed to a few pixels.</span>}
                        <Btn small onClick={() => { setDomain([...recSpan]); setSelFault(null); }}>Fit recorder</Btn>
                      </div>
                    );
                  })()}

                  <div ref={chartWrapRef}
                    onPointerDown={chartPointerDown} onPointerMove={chartPointerMove}
                    onPointerUp={chartPointerUp} onPointerLeave={chartPointerUp}
                    style={{ cursor: selectMode ? "crosshair" : "grab", touchAction: "pan-y", userSelect: "none", position: "relative" }}>
                    {band && (
                      <div style={{ position: "absolute", top: 0, bottom: 0, left: band.x1, width: Math.max(band.x2 - band.x1, 1), background: "rgba(0,153,153,0.12)", border: `1px solid ${C.amber}`, pointerEvents: "none", zIndex: 5 }} />
                    )}
                    {(!v2Active || viewMode === "combined") ? (
                      <>
                        {renderChart(combinedRows, combinedSeries, [...visibleFaults1, ...visibleFaults2], combinedSeries.length ? ((laneDefs1.length + laneDefs2.length) ? 300 : 420) : 70, snapDots, [{ span: recSpan1, color: VEH[1].tag }, { span: recSpan2, color: VEH[2].tag }])}
                        {renderFaultLanes(1, name1)}
                        {v2Active && renderFaultLanes(2, name2)}
                        {renderLanes(laneDefs1, laneRows1, visibleFaults1, 1)}
                        {v2Active && renderLanes(laneDefs2, laneRows2, visibleFaults2, 2)}
                      </>
                    ) : (
                      <>
                        <div style={{ fontSize: 11, color: VEH[1].tag, fontFamily: "'Barlow Condensed', sans-serif", fontWeight: 700, letterSpacing: "0.12em", textTransform: "uppercase", margin: "2px 0 2px 8px" }}>{name1}</div>
                        {renderChart(sepRows1, sepSeries1, visibleFaults1, sepSeries1.length ? (laneDefs1.length ? 190 : 240) : 60, snapDots.filter((d) => d.vi === 1), [{ span: recSpan1, color: VEH[1].tag }])}
                        {renderFaultLanes(1, name1)}
                        {renderLanes(laneDefs1, laneRows1, visibleFaults1, 1)}
                        <div style={{ fontSize: 11, color: VEH[2].tag, fontFamily: "'Barlow Condensed', sans-serif", fontWeight: 700, letterSpacing: "0.12em", textTransform: "uppercase", margin: "8px 0 2px 8px" }}>{name2}</div>
                        {renderChart(sepRows2, sepSeries2, visibleFaults2, sepSeries2.length ? (laneDefs2.length ? 190 : 240) : 60, snapDots.filter((d) => d.vi === 2), [{ span: recSpan2, color: VEH[2].tag }])}
                        {renderFaultLanes(2, name2)}
                        {renderLanes(laneDefs2, laneRows2, visibleFaults2, 2)}
                      </>
                    )}
                  </div>

                  {/* minimap: spans the recorder recording */}
                  {navFull && (
                    <div onPointerDown={mmPointerDown} onPointerMove={mmPointerMove} onPointerUp={mmPointerUp} onPointerLeave={mmPointerUp}
                      style={{ position: "relative", height: 36, background: "#f3f3f0", border: `1px solid ${C.faint}`, borderRadius: 4, marginTop: 10, touchAction: "none", cursor: "pointer", overflow: "hidden", userSelect: "none" }}>
                      {[{ span: recSpan1, col: VEH[1].tag }, { span: recSpan2, col: VEH[2].tag }].map((b, i) => b.span && (
                        <div key={i} style={{
                          position: "absolute", top: i === 0 ? 0 : "50%", height: v2Active ? "50%" : "100%", bottom: 0,
                          left: `${((b.span[0] - navFull[0]) / (navFull[1] - navFull[0])) * 100}%`,
                          width: `${Math.max(((b.span[1] - b.span[0]) / (navFull[1] - navFull[0])) * 100, 0.4)}%`,
                          background: `${b.col}2e`, borderLeft: `1px solid ${b.col}`, borderRight: `1px solid ${b.col}`,
                        }} />
                      ))}
                      {mmTicks.map((f, i) => (
                        <div key={i} style={{ position: "absolute", top: 6, bottom: 6, left: `${f.frac * 100}%`, width: 1.5, background: VEH[f.vi].fault, opacity: 0.55 }} />
                      ))}
                      <div style={{
                        position: "absolute", top: 0, bottom: 0,
                        left: `${((activeDomain[0] - navFull[0]) / (navFull[1] - navFull[0])) * 100}%`,
                        width: `${Math.max(((activeDomain[1] - activeDomain[0]) / (navFull[1] - navFull[0])) * 100, 0.6)}%`,
                        background: "rgba(0,153,153,0.10)", border: `1px solid ${C.amber}`, borderRadius: 3, cursor: "grab",
                      }} />
                      <div style={{ position: "absolute", left: 6, bottom: 2, fontSize: 9, color: C.dim, fontFamily: "'IBM Plex Mono', monospace", pointerEvents: "none" }}>{fmtFull(navFull[0])}</div>
                      <div style={{ position: "absolute", right: 6, bottom: 2, fontSize: 9, color: C.dim, fontFamily: "'IBM Plex Mono', monospace", pointerEvents: "none" }}>{fmtFull(navFull[1])}</div>
                    </div>
                  )}

                  {/* ---------- signal snapshot at selected fault ---------- */}
                  {selSnapshot && (
                    <div style={{ background: "#f3f3f0", border: `1px solid ${C.amber}`, borderRadius: 6, padding: 12, marginTop: 10 }}>
                      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8, marginBottom: 8 }}>
                        <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 12 }}>
                          <span style={{ color: VEH[selSnapshot.vi].fault, fontWeight: 600 }}>{selSnapshot.f.code}</span>
                          {selSnapshot.f.desc && <span style={{ color: C.ink }}> — {selSnapshot.f.desc}</span>}
                          <span style={{ color: VEH[selSnapshot.vi].tag }}> · {selSnapshot.vi === 1 ? name1 : name2}</span>
                          <span style={{ color: C.dim }}> · {fmtFull(selSnapshot.f.t)}</span>
                        </div>
                        <div style={{ display: "flex", gap: 6, alignItems: "center" }}>
                          <Btn small onClick={() => navFault(-1)}>◀ prev fault</Btn>
                          <Btn small onClick={() => navFault(1)}>next fault ▶</Btn>
                          <span onClick={() => setSelFault(null)} style={{ color: C.dim, cursor: "pointer", padding: "0 4px", fontSize: 14 }}>✕</span>
                        </div>
                      </div>
                      {selSnapshot.outside || !selSnapshot.row ? (
                        <div style={{ fontSize: 11, color: C.dim, fontFamily: "'IBM Plex Mono', monospace" }}>This fault is outside the recorder's coverage window — no signal data available at this instant.</div>
                      ) : (
                        <>
                          <div style={{ fontSize: 10, color: C.dim, fontFamily: "'IBM Plex Mono', monospace", marginBottom: 8 }}>
                            Signal values at nearest recorder sample ({selSnapshot.dt >= 0 ? "+" : ""}{(selSnapshot.dt / 1000).toFixed(3)} s from fault timestamp){offset1 !== 0 || offset2 !== 0 ? " · time-sync offset applied" : ""} — shown as dots on the graph
                          </div>
                          <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
                            {selSnapshot.rows.map(({ s, v, binary }) => (
                              <div key={s} style={{ display: "flex", alignItems: "center", gap: 6, background: "#fff", border: `1px solid ${C.panelEdge}`, borderLeft: `3px solid ${seriesColorOf(selSnapshot.vi, s)}`, borderRadius: 4, padding: "5px 9px", fontFamily: "'IBM Plex Mono', monospace", fontSize: 11 }}>
                                <span style={{ color: C.dim, maxWidth: 190, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }} title={s}>{s}</span>
                                <span style={{ color: binary ? (v >= 0.5 ? C.green : C.red) : C.ink, fontWeight: 600 }}>
                                  {v == null ? "—" : binary ? (v >= 0.5 ? "HIGH (1)" : "LOW (0)") : v.toLocaleString(undefined, { maximumFractionDigits: 3 })}
                                </span>
                              </div>
                            ))}
                            {!selSnapshot.rows.length && <span style={{ fontSize: 11, color: C.dim }}>No signals selected for this vehicle.</span>}
                          </div>
                        </>
                      )}
                    </div>
                  )}

                  <div style={{ fontSize: 10, color: C.dim, marginTop: 6 }}>
                    <span style={{ color: VEH[1].fault }}>┆</span> {name1} fault{v2Active && <> · <span style={{ color: VEH[2].fault }}>┆</span> {name2} fault (dashed lines = {name2} signals in combined view)</>} · click a fault (list) to inspect signal values at that instant · drag chart to pan · wheel to zoom
                  </div>
                </div>
              </div>

              {/* fault lists: one per source, side by side underneath the graph */}
              <div style={{ width: "100%", display: "flex", gap: 14, flexWrap: "wrap" }}>
                {[
                  { srcKey: "expert", title: "Expert Fault Log", search: faultSearchE, setSearch: setFaultSearchE, sel: visExpert, clearSel: () => setVisExpert([]), accent: C.red },
                  { srcKey: "vcu", title: "VCU Fault Log", search: faultSearchV, setSearch: setFaultSearchV, sel: visVcu, clearSel: () => setVisVcu([]), accent: C.violet },
                ].map(({ srcKey, title, search, setSearch, sel, clearSel, accent }) => {
                  const all = mergedFaults.filter((f) => (f.src || "expert") === srcKey);
                  if (!all.length) return null;
                  const q = search.trim().toLowerCase();
                  const listFaults = all.filter((f) =>
                    srcShown(f) &&
                    (!q || f.code.toLowerCase().includes(q) || f.desc.toLowerCase().includes(q) || (f.veh || "").toLowerCase().includes(q) || fmtFull(f.t).includes(q) || (f.vi === 1 ? name1 : name2).toLowerCase().includes(q))
                  );
                  return (
                    <div key={srcKey} style={{ flex: "1 1 250px", minWidth: 250, background: C.panel, border: `1px solid ${C.panelEdge}`, borderTop: `3px solid ${accent}`, borderRadius: 6, padding: 14, maxHeight: 420, display: "flex", flexDirection: "column" }}>
                      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
                        <Label>{title} ({listFaults.length}{listFaults.length !== all.length ? ` of ${all.length}` : ""})</Label>
                        {(sel.length > 0 || q) && <span onClick={() => { clearSel(); setSearch(""); }} style={{ color: C.amber, fontSize: 11, cursor: "pointer" }}>clear ✕</span>}
                      </div>
                      <input value={search} onChange={(e) => setSearch(e.target.value)} placeholder={`Search ${srcKey === "vcu" ? "VCU" : "Expert"} faults… code, description, time`}
                        style={{ width: "100%", boxSizing: "border-box", background: "#f3f3f0", color: C.ink, border: `1px solid ${C.faint}`, borderRadius: 4, padding: "7px 10px", fontFamily: "'IBM Plex Mono', monospace", fontSize: 11, outline: "none", marginBottom: 8 }} />
                      <div style={{ overflowY: "auto", display: "flex", flexDirection: "column", gap: 6 }}>
                        {listFaults.slice(0, 800).map((f) => {
                          const key = `${f.vi}-${f.id}`;
                          const span = f.vi === 1 ? recSpan1 : recSpan2;
                          const inRec = span && f.t >= span[0] && f.t <= span[1];
                          return (
                            <div key={key} onClick={() => zoomToFault(f)}
                              style={{
                                padding: "8px 10px", borderRadius: 4, cursor: "pointer",
                                background: key === selFault ? "#e3f3f3" : "#f3f3f0",
                                border: `1px solid ${key === selFault ? C.amber : C.faint}`,
                                borderLeft: `3px solid ${VEH[f.vi].tag}`,
                                opacity: inRec ? 1 : 0.55,
                              }}>
                              <div style={{ display: "flex", justifyContent: "space-between", gap: 8 }}>
                                <span style={{ color: VEH[f.vi].fault, fontWeight: 600, fontSize: 12 }}>
                                  {f.code}
                                  <span style={{ color: VEH[f.vi].tag, fontWeight: 400 }}> · {f.vi === 1 ? name1 : name2}{f.veh ? ` (${f.veh})` : ""}</span>
                                  {f.open && <span style={{ color: "#fff", background: C.red, borderRadius: 2, fontSize: 8, fontWeight: 700, padding: "1px 4px", marginLeft: 4, verticalAlign: "1px", letterSpacing: "0.06em" }}>ACTIVE</span>}
                                </span>
                                <span style={{ color: C.dim, fontSize: 10 }}>{fmtFull(f.t)}{f.tEnd != null ? ` · ${fmtDur(f.tEnd - f.t)}` : ""}</span>
                              </div>
                              {f.desc && <div style={{ color: C.ink, fontSize: 11, marginTop: 2 }}>{f.desc}</div>}
                              {!inRec && <div style={{ color: C.dim, fontSize: 9, marginTop: 2 }}>outside recorder window</div>}
                            </div>
                          );
                        })}
                        {!listFaults.length && <div style={{ color: C.dim, fontSize: 11, padding: 8 }}>No {srcKey === "vcu" ? "VCU" : "Expert"} faults match the current search/code selection.</div>}
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>
          )}

          {/* ---------- STATS ---------- */}
          {tab === "stats" && (
            <div style={{ display: "flex", flexDirection: "column", gap: 18, marginTop: 14 }}>
              {v2Active && (
                <div style={{ display: "flex", gap: 6, alignItems: "center" }}>
                  <span style={{ fontSize: 10, color: C.dim, letterSpacing: "0.1em", textTransform: "uppercase", fontFamily: "'Barlow Condensed', sans-serif", fontWeight: 600 }}>Scope:</span>
                  <Chip active={statsScope === "all"} onClick={() => setStatsScope("all")}>Both vehicles</Chip>
                  <Chip active={statsScope === "v1"} onClick={() => setStatsScope("v1")} color={VEH[1].tag}>{name1}</Chip>
                  <Chip active={statsScope === "v2"} onClick={() => setStatsScope("v2")} color={VEH[2].tag}>{name2}</Chip>
                </div>
              )}
              {/* ---------- AI analysis ---------- */}
              {mergedFaults.length > 0 && (
                <div style={{ background: C.panel, border: `1px solid ${C.panelEdge}`, borderTop: `3px solid ${C.amber}`, borderRadius: 6, padding: 18 }}>
                  <Label>AI analysis — ask about the events</Label>
                  <div style={{ display: "flex", gap: 6, alignItems: "center", marginBottom: 10, flexWrap: "wrap" }}>
                    <span style={{ fontSize: 10, color: C.dim, letterSpacing: "0.1em", textTransform: "uppercase", fontFamily: "'Barlow Condensed', sans-serif", fontWeight: 600 }}>Provider:</span>
                    <Chip active={aiProvider === "claude"} onClick={() => { setAiProvider("claude"); setAiError(null); }}>Claude API</Chip>
                    <Chip active={aiProvider === "ollama"} onClick={() => { setAiProvider("ollama"); setAiError(null); }} color={C.green}>Ollama (local, free)</Chip>
                    {aiProvider === "ollama" && (
                      <>
                        <input value={ollamaUrl} onChange={(e) => setOllamaUrl(e.target.value)} placeholder="http://localhost:11434"
                          style={{ width: 210, background: "#f3f3f0", color: C.ink, border: `1px solid ${C.faint}`, borderRadius: 4, padding: "6px 9px", fontFamily: "'IBM Plex Mono', monospace", fontSize: 11, outline: "none" }} />
                        <input value={ollamaModel} onChange={(e) => setOllamaModel(e.target.value)} placeholder="model e.g. llama3.1" list="ollama-models"
                          style={{ width: 150, background: "#f3f3f0", color: C.ink, border: `1px solid ${C.faint}`, borderRadius: 4, padding: "6px 9px", fontFamily: "'IBM Plex Mono', monospace", fontSize: 11, outline: "none" }} />
                        <datalist id="ollama-models">{ollamaModels.map((m) => <option key={m} value={m} />)}</datalist>
                        <Btn small onClick={listOllamaModels}>List models</Btn>
                      </>
                    )}
                  </div>
                  {aiMessages.length > 0 && (
                    <div style={{ display: "flex", flexDirection: "column", gap: 8, maxHeight: 340, overflowY: "auto", marginBottom: 10 }}>
                      {aiMessages.map((m, i) => (
                        <div key={i} style={{
                          alignSelf: m.role === "user" ? "flex-end" : "flex-start", maxWidth: "85%",
                          background: m.role === "user" ? "#e3f3f3" : "#f3f3f0",
                          border: `1px solid ${m.role === "user" ? C.amber : C.panelEdge}`,
                          borderRadius: 6, padding: "8px 12px", fontFamily: "'IBM Plex Mono', monospace", fontSize: 12, color: C.ink, whiteSpace: "pre-wrap",
                        }}>{m.content}</div>
                      ))}
                      {aiBusy && <div style={{ alignSelf: "flex-start", color: C.dim, fontFamily: "'IBM Plex Mono', monospace", fontSize: 12 }}>Analyzing…</div>}
                    </div>
                  )}
                  {aiError && <div style={{ color: C.red, fontFamily: "'IBM Plex Mono', monospace", fontSize: 11, marginBottom: 8 }}>Error: {aiError}</div>}
                  <div style={{ display: "flex", gap: 6, flexWrap: "wrap", marginBottom: 8 }}>
                    <Chip onClick={() => askAi("Summarize the fault history: dominant failure modes, notable sequences, and anything that looks systemic.")}>Summarize all faults</Chip>
                    <Chip onClick={() => askAi("What happened during the selected event? Walk through the fault, the signal states at that instant, and likely related events around it.")} color={selFault ? C.amber : undefined}>What happened during this event?</Chip>
                    <Chip onClick={() => askAi("Are there patterns by time of day, repeated sequences of codes, or faults that tend to occur together?")}>Find patterns</Chip>
                    {aiMessages.length > 0 && <Chip onClick={() => { setAiMessages([]); setAiError(null); }}>Clear chat</Chip>}
                  </div>
                  <div style={{ display: "flex", gap: 8 }}>
                    <input value={aiInput} onChange={(e) => setAiInput(e.target.value)}
                      onKeyDown={(e) => { if (e.key === "Enter") askAi(); }}
                      placeholder={selFault ? "Ask about the selected event or the fault history…" : "Ask anything about the fault data… (select an event on the Timeline for event-specific answers)"}
                      style={{ flex: 1, background: "#f3f3f0", color: C.ink, border: `1px solid ${C.faint}`, borderRadius: 4, padding: "9px 12px", fontFamily: "'IBM Plex Mono', monospace", fontSize: 12, outline: "none" }} />
                    <Btn small primary disabled={aiBusy || !aiInput.trim()} onClick={() => askAi()}>{aiBusy ? "…" : "Ask"}</Btn>
                  </div>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginTop: 6, gap: 10, flexWrap: "wrap" }}>
                    <div style={{ fontSize: 10, color: C.dim, fontFamily: "'IBM Plex Mono', monospace" }}>
                      The AI sees: per-source fault summaries, events in the current timeline window, and the selected event's signal snapshot.
                    </div>
                    {aiProvider === "claude" ? (
                      <span onClick={() => setAiKeyOpen(!aiKeyOpen)} style={{ fontSize: 10, color: aiKey ? C.green : C.dim, cursor: "pointer", fontFamily: "'IBM Plex Mono', monospace", textDecoration: "underline", whiteSpace: "nowrap" }}>
                        {aiKey ? "API key set ✓" : "API key (self-hosted)"}
                      </span>
                    ) : (
                      <span style={{ fontSize: 10, color: C.dim, fontFamily: "'IBM Plex Mono', monospace", whiteSpace: "nowrap" }}>
                        data stays on your machine 🔒
                      </span>
                    )}
                  </div>
                  {aiProvider === "ollama" && (
                    <div style={{ marginTop: 8, background: "#f3f3f0", border: `1px solid ${C.faint}`, borderRadius: 4, padding: 10, fontSize: 10, color: C.dim, fontFamily: "'IBM Plex Mono', monospace" }}>
                      Setup: install Ollama (ollama.com), pull a model (`ollama pull llama3.1` — for this analysis a larger model like qwen2.5:14b or llama3.1:70b answers noticeably better), and allow this page's origin before starting it: `OLLAMA_ORIGINS="https://YOURNAME.github.io" ollama serve` (Windows: set it as an environment variable). Requests go from your browser straight to localhost — nothing leaves your machine.
                    </div>
                  )}
                  {aiProvider === "claude" && aiKeyOpen && (
                    <div style={{ marginTop: 8, background: "#f3f3f0", border: `1px solid ${C.faint}`, borderRadius: 4, padding: 10 }}>
                      <div style={{ fontSize: 10, color: C.dim, fontFamily: "'IBM Plex Mono', monospace", marginBottom: 6 }}>
                        Hosting this outside Claude.ai (e.g. GitHub Pages)? Enter an Anthropic API key (console.anthropic.com). It is kept in memory only — never saved, never sent anywhere except api.anthropic.com. Do NOT hardcode it into the source or commit it to the repo.
                      </div>
                      <input type="password" value={aiKey} onChange={(e) => setAiKey(e.target.value)} placeholder="sk-ant-…"
                        autoComplete="off"
                        style={{ width: "100%", boxSizing: "border-box", background: "#fff", color: C.ink, border: `1px solid ${aiKey ? C.green : C.faint}`, borderRadius: 4, padding: "8px 10px", fontFamily: "'IBM Plex Mono', monospace", fontSize: 12, outline: "none" }} />
                    </div>
                  )}
                </div>
              )}

              {!stats ? (
                <div style={{ color: C.dim, fontSize: 12 }}>No fault data in this scope.</div>
              ) : (
                <>
                  <div style={{ display: "flex", gap: 14, flexWrap: "wrap" }}>
                    {[
                      ["Total faults", stats.total.toLocaleString(), C.red],
                      ["Observation span", stats.spanDays >= 1 ? stats.spanDays.toFixed(1) + " days" : (stats.spanDays * 24).toFixed(1) + " hrs", C.cyan],
                      ["Fault rate", stats.perDay.toFixed(1) + " / day", C.amber],
                      ["Mean time between faults", stats.mtbf ? (stats.mtbf / 60000 >= 90 ? (stats.mtbf / 3600000).toFixed(1) + " hr" : (stats.mtbf / 60000).toFixed(0) + " min") : "—", C.green],
                      ["Top code", stats.pareto[0]?.code || "—", C.violet],
                    ].map(([k, v, col]) => (
                      <div key={k} style={{ flex: "1 1 150px", background: C.panel, border: `1px solid ${C.panelEdge}`, borderLeft: `3px solid ${col}`, borderRadius: 6, padding: "14px 16px" }}>
                        <div style={{ fontSize: 10, color: C.dim, letterSpacing: "0.12em", textTransform: "uppercase", fontFamily: "'Barlow Condensed', sans-serif", fontWeight: 600 }}>{k}</div>
                        <div style={{ fontSize: 22, fontWeight: 600, color: col, marginTop: 4 }}>{v}</div>
                      </div>
                    ))}
                  </div>

                  <div style={{ display: "flex", gap: 18, flexWrap: "wrap" }}>
                    <div style={{ flex: "1 1 420px", background: C.panel, border: `1px solid ${C.panelEdge}`, borderRadius: 6, padding: 18 }}>
                      <Label>Fault frequency by code — top 12 (click to filter timeline)</Label>
                      <ResponsiveContainer width="100%" height={300}>
                        <BarChart data={stats.pareto.slice(0, 12)} margin={{ top: 10, right: 10, bottom: 5, left: 0 }}>
                          <CartesianGrid stroke={C.faint} strokeDasharray="2 6" vertical={false} />
                          <XAxis dataKey="code" stroke={C.dim} tick={{ fontSize: 10, fontFamily: "'IBM Plex Mono', monospace" }} interval={0} angle={-30} textAnchor="end" height={55} />
                          <YAxis stroke={C.dim} tick={{ fontSize: 10, fontFamily: "'IBM Plex Mono', monospace" }} allowDecimals={false} width={40} />
                          <Tooltip content={<DarkTooltip />} cursor={{ fill: "rgba(0,0,0,0.05)" }} />
                          <Bar dataKey="count" onClick={(d) => { const code = d?.code ?? d?.payload?.code; if (code != null) { toggleCode(code); setTab("timeline"); } }} cursor="pointer" isAnimationActive={false}>
                            {stats.pareto.slice(0, 12).map((e, i) => (
                              <Cell key={i} fill={visExpert.includes(e.code) || visVcu.includes(e.code) ? C.amber : C.red} fillOpacity={0.85} />
                            ))}
                          </Bar>
                        </BarChart>
                      </ResponsiveContainer>
                    </div>

                    <div style={{ flex: "1 1 420px", background: C.panel, border: `1px solid ${C.panelEdge}`, borderRadius: 6, padding: 18 }}>
                      <Label>Fault trend ({stats.hourly ? "per hour" : "per day"})</Label>
                      <ResponsiveContainer width="100%" height={300}>
                        <BarChart data={stats.trend} margin={{ top: 10, right: 10, bottom: 5, left: 0 }}>
                          <CartesianGrid stroke={C.faint} strokeDasharray="2 6" vertical={false} />
                          <XAxis dataKey="t" stroke={C.dim} tickFormatter={(t) => (stats.hourly ? fmtTime(t, 3600000) : fmtDay(t))} tick={{ fontSize: 10, fontFamily: "'IBM Plex Mono', monospace" }} />
                          <YAxis stroke={C.dim} tick={{ fontSize: 10, fontFamily: "'IBM Plex Mono', monospace" }} allowDecimals={false} width={40} />
                          <Tooltip content={<DarkTooltip />} cursor={{ fill: "rgba(0,0,0,0.05)" }} />
                          <Bar dataKey="count" name="faults" fill={C.cyan} fillOpacity={0.85} isAnimationActive={false} />
                        </BarChart>
                      </ResponsiveContainer>
                    </div>
                  </div>

                  <div style={{ background: C.panel, border: `1px solid ${C.panelEdge}`, borderRadius: 6, padding: 18, overflowX: "auto" }}>
                    <Label>Fault code summary ({stats.pareto.length} distinct codes)</Label>
                    <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                      <thead>
                        <tr style={{ color: C.dim, textAlign: "left" }}>
                          {["Code", "Description", "Count", "% of total"].map((h) => (
                            <th key={h} style={{ padding: "8px 10px", borderBottom: `1px solid ${C.faint}`, fontFamily: "'Barlow Condensed', sans-serif", letterSpacing: "0.1em", textTransform: "uppercase", fontSize: 11 }}>{h}</th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {stats.pareto.slice(0, 60).map((p) => (
                          <tr key={p.code} style={{ borderBottom: `1px solid ${C.panelEdge}`, cursor: "pointer", background: visExpert.includes(p.code) || visVcu.includes(p.code) ? "#e3f3f3" : "transparent" }}
                            onClick={() => toggleCode(p.code)}>
                            <td style={{ padding: "7px 10px", color: C.red, fontWeight: 600 }}>{p.code}</td>
                            <td style={{ padding: "7px 10px", color: C.ink }}>{p.desc || "—"}</td>
                            <td style={{ padding: "7px 10px", color: C.amber }}>{p.count}</td>
                            <td style={{ padding: "7px 10px", color: C.dim }}>{((p.count / stats.total) * 100).toFixed(1)}%</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </>
              )}
            </div>
          )}
        </div>
      )}
    </div>
  );
}
