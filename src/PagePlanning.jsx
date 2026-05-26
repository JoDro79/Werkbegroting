import { useState, useCallback } from "react";
import * as XLSX from "xlsx";

const HOOFDSTUKKEN = [
  "VOORBEREIDENDE WERKZAAMHEDEN",
  "VERHARDINGEN",
  "GRONDWERK SLEUVEN",
  "GROENVOORZIENINGEN",
  "KRUISINGEN",
  "KABEL/LEIDINGWERK ELEKTRA",
  "MONTAGEWERKZAAMHEDEN ELEKTRA",
  "AARDING CS - OS (op het veld)",
  "KABEL/LEIDINGWERK INFORMATIE DISTRIBUTIE MS",
  "STATIONS",
  "BIJKOMENDE WERKZAAMHEDEN",
  "DIRECTIEBEHOEFTEN",
  "OPLEVERING EN REVISIE",
  "TER BESCHIKKING STELLEN",
];

const CHAP_SHORT = {
  "VOORBEREIDENDE WERKZAAMHEDEN": "Voorbereiding",
  "VERHARDINGEN": "Verhardingen",
  "GRONDWERK SLEUVEN": "Grondwerk",
  "GROENVOORZIENINGEN": "Groenwerk",
  "KRUISINGEN": "Kruisingen",
  "KABEL/LEIDINGWERK ELEKTRA": "Kabelwerk LS/MS",
  "MONTAGEWERKZAAMHEDEN ELEKTRA": "Montage",
  "AARDING CS - OS (op het veld)": "Aarding",
  "KABEL/LEIDINGWERK INFORMATIE DISTRIBUTIE MS": "Glasvezel",
  "STATIONS": "Stations",
  "BIJKOMENDE WERKZAAMHEDEN": "Bijkomend",
  "DIRECTIEBEHOEFTEN": "Directie",
  "OPLEVERING EN REVISIE": "Oplevering",
  "TER BESCHIKKING STELLEN": "T.b.s.",
};

function parseDate(v) {
  if (!v) return null;
  if (v instanceof Date) return v;
  // Excel serial number
  if (typeof v === "number") {
    const d = new Date((v - 25569) * 86400 * 1000);
    return isNaN(d) ? null : d;
  }
  // String formats: dd-mm-yyyy, dd/mm/yyyy, yyyy-mm-dd
  const s = String(v).trim();
  const m1 = s.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$/);
  if (m1) return new Date(+m1[3], +m1[2] - 1, +m1[1]);
  const m2 = s.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/);
  if (m2) return new Date(+m2[1], +m2[2] - 1, +m2[3]);
  const d = new Date(s);
  return isNaN(d) ? null : d;
}

function parseMSProject(arrayBuffer) {
  const wb = XLSX.read(new Uint8Array(arrayBuffer), { type: "array", cellDates: true });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: false });

  // Find header row (contains "Taaknaam" or "Task Name" or "Name")
  let headerRow = -1;
  let colMap = {};
  for (let i = 0; i < Math.min(10, data.length); i++) {
    const row = data[i].map(c => c ? String(c).toLowerCase().trim() : "");
    const nameIdx = row.findIndex(c =>
      c.includes("taaknaam") || c === "name" || c === "task name" || c === "naam"
    );
    if (nameIdx >= 0) {
      headerRow = i;
      row.forEach((c, j) => {
        if (c.includes("taaknaam") || c === "name" || c === "task name" || c === "naam") colMap.name = j;
        if (c.includes("begin") || c === "start") colMap.start = j;
        if (c.includes("voltooiing") || c === "finish" || c === "einde" || c === "end") colMap.finish = j;
        if (c.includes("duur") || c === "duration") colMap.duration = j;
        if (c.includes("overzicht") || c.includes("outline") || c.includes("level") || c.includes("niveau")) colMap.level = j;
        if (c.includes("% vol") || c.includes("complete") || c.includes("voortgang")) colMap.pct = j;
        if (c.includes("id") && j < 3) colMap.id = j;
      });
      break;
    }
  }

  if (headerRow < 0 || colMap.name === undefined) {
    // Fallback: assume first row is header, col 1 = name, col 3 = start, col 4 = finish
    headerRow = 0;
    colMap = { id: 0, name: 1, duration: 2, start: 3, finish: 4, pct: 5, level: 8 };
  }

  const tasks = [];
  for (let i = headerRow + 1; i < data.length; i++) {
    const row = data[i];
    const name = row[colMap.name] ? String(row[colMap.name]).trim() : "";
    if (!name || name === "") continue;

    const start = parseDate(colMap.start !== undefined ? row[colMap.start] : null);
    const finish = parseDate(colMap.finish !== undefined ? row[colMap.finish] : null);
    const level = colMap.level !== undefined ? (parseInt(row[colMap.level]) || 0) : name.match(/^(\s+)/)?.[1]?.length / 2 || 0;
    const pct = colMap.pct !== undefined ? (parseFloat(String(row[colMap.pct]).replace("%","")) || 0) : 0;
    const id = colMap.id !== undefined ? (row[colMap.id] || i) : i;

    // Determine indent level from leading spaces if no outline level col
    const indent = name.match(/^(\s+)/)?.[1]?.length || 0;
    const cleanName = name.trim();
    const effectiveLevel = level || Math.floor(indent / 2);

    if (start && finish && cleanName) {
      tasks.push({ id: String(id), name: cleanName, start, finish, level: effectiveLevel, pct, indent });
    }
  }

  return tasks;
}

// ── Component ─────────────────────────────────────────────────────────────────
export default function PagePlanning({ tasks, setTasks, koppelingen, setKoppelingen, posts }) {
  const [dragging, setDragging] = useState(false);
  const fileRef = useState(null);

  const handleFile = useCallback(async (file) => {
    const buf = await file.arrayBuffer();
    const parsed = parseMSProject(buf);
    setTasks(parsed);
    // Auto-suggest koppelingen based on name similarity
    const suggested = {};
    parsed.forEach(t => {
      if (!suggested[t.id]) {
        const nl = t.name.toLowerCase();
        const matches = [];
        if (nl.includes("grond") || nl.includes("sleuf") || nl.includes("graaf")) matches.push("GRONDWERK SLEUVEN");
        if (nl.includes("kabel") && !nl.includes("montage") && !nl.includes("glasvezel")) matches.push("KABEL/LEIDINGWERK ELEKTRA");
        if (nl.includes("montage") || nl.includes("mof") || nl.includes("eindsluiting")) matches.push("MONTAGEWERKZAAMHEDEN ELEKTRA");
        if (nl.includes("hdd") || nl.includes("boring") || nl.includes("kruising") || nl.includes("persing") || nl.includes("bemaling")) matches.push("KRUISINGEN");
        if (nl.includes("verharding") || nl.includes("asfalt") || nl.includes("straat") || nl.includes("tegel")) matches.push("VERHARDINGEN");
        if (nl.includes("voorber") || nl.includes("inrichten") || nl.includes("werkterrein")) matches.push("VOORBEREIDENDE WERKZAAMHEDEN");
        if (nl.includes("groen") || nl.includes("inzaai") || nl.includes("berm")) matches.push("GROENVOORZIENINGEN");
        if (nl.includes("oplevert") || nl.includes("revisie") || nl.includes("eindmeting") || nl.includes("dossier")) matches.push("OPLEVERING EN REVISIE");
        if (nl.includes("station")) matches.push("STATIONS");
        if (nl.includes("glasvezel") || nl.includes("fiber")) matches.push("KABEL/LEIDINGWERK INFORMATIE DISTRIBUTIE MS");
        if (nl.includes("aarding")) matches.push("AARDING CS - OS (op het veld)");
        if (matches.length > 0) {
          const pctEach = Math.round(100 / matches.length);
          suggested[t.id] = matches.map((h, i) => ({
            hoofdstuk: h,
            percentage: i === matches.length - 1 ? 100 - pctEach * (matches.length - 1) : pctEach,
          }));
        }
      }
    });
    setKoppelingen(suggested);
  }, [setTasks, setKoppelingen]);

  const onDrop = useCallback(e => {
    e.preventDefault();
    setDragging(false);
    const f = e.dataTransfer.files[0];
    if (f) handleFile(f);
  }, [handleFile]);

  const addKoppeling = (taskId) => {
    setKoppelingen(prev => ({
      ...prev,
      [taskId]: [...(prev[taskId] || []), { hoofdstuk: HOOFDSTUKKEN[0], percentage: 100 }],
    }));
  };

  const updateKoppeling = (taskId, idx, field, value) => {
    setKoppelingen(prev => ({
      ...prev,
      [taskId]: prev[taskId].map((k, i) => i === idx ? { ...k, [field]: value } : k),
    }));
  };

  const removeKoppeling = (taskId, idx) => {
    setKoppelingen(prev => ({
      ...prev,
      [taskId]: prev[taskId].filter((_, i) => i !== idx),
    }));
  };

  const totalPct = (taskId) => (koppelingen[taskId] || []).reduce((s, k) => s + (k.percentage || 0), 0);

  const fmtDate = d => d ? d.toLocaleDateString("nl-NL", { day: "2-digit", month: "2-digit", year: "numeric" }) : "—";

  // Costs per hoofdstuk
  const kostenPerHoofdstuk = {};
  if (posts) {
    posts.forEach(p => {
      const k = p.hoeveelheid * p.kostprijs;
      kostenPerHoofdstuk[p.hoofdstuk] = (kostenPerHoofdstuk[p.hoofdstuk] || 0) + k;
    });
  }
  const euro = v => new Intl.NumberFormat("nl-NL", { style: "currency", currency: "EUR", maximumFractionDigits: 0 }).format(v || 0);

  return (
    <div style={{ maxWidth: 1100, margin: "0 auto", padding: "20px 20px 60px" }}>
      {tasks.length === 0 ? (
        // ── Upload screen ──
        <div style={{ maxWidth: 600, margin: "40px auto" }}>
          <div style={{ marginBottom: 28 }}>
            <h2 style={{ fontSize: 18, fontWeight: 900, color: "#90caf9", letterSpacing: 2, marginBottom: 6 }}>PLANNING LADEN</h2>
            <p style={{ fontSize: 12, color: "#546e7a", lineHeight: 1.8 }}>
              Exporteer je MS Project planning naar Excel:<br />
              <span style={{ color: "#78909c" }}>Bestand → Opslaan als → Excel-werkmap (.xlsx)</span><br />
              Of: Rapport → Visuele rapporten → kies taakoverzicht
            </p>
          </div>
          <div
            onDrop={onDrop}
            onDragOver={e => { e.preventDefault(); setDragging(true); }}
            onDragLeave={() => setDragging(false)}
            onClick={() => document.getElementById("planningFile").click()}
            style={{
              border: `2px dashed ${dragging ? "#42a5f5" : "#1e4976"}`,
              borderRadius: 14, padding: "50px 28px", textAlign: "center",
              cursor: "pointer", background: dragging ? "rgba(25,118,210,.1)" : "rgba(25,118,210,.03)",
              transition: "all .2s",
            }}
          >
            <div style={{ fontSize: 40, marginBottom: 12 }}>📅</div>
            <div style={{ fontSize: 14, color: "#90caf9", fontWeight: 700, marginBottom: 6 }}>Sleep MS Project export hierheen</div>
            <div style={{ fontSize: 11, color: "#37474f" }}>Excel export (.xlsx) van MS Project taakoverzicht</div>
          </div>
          <input id="planningFile" type="file" accept=".xlsx" style={{ display: "none" }}
            onChange={e => e.target.files[0] && handleFile(e.target.files[0])} />

          <div style={{ marginTop: 24, padding: 18, background: "rgba(255,255,255,.03)", borderRadius: 10, border: "1px solid #1e4976" }}>
            <div style={{ fontSize: 10, color: "#546e7a", letterSpacing: 1, marginBottom: 10 }}>HOE TE EXPORTEREN UIT MS PROJECT</div>
            {[
              "Open je planning in MS Project",
              "Klik op Bestand → Opslaan als",
              "Kies als bestandstype: Excel-werkmap (*.xlsx)",
              "Sla op en upload hier",
            ].map((s, i) => (
              <div key={i} style={{ display: "flex", gap: 10, marginBottom: 7, fontSize: 12, color: "#78909c" }}>
                <span style={{ color: "#1976d2", fontWeight: 700, minWidth: 18 }}>{i + 1}.</span> {s}
              </div>
            ))}
          </div>
        </div>
      ) : (
        // ── Koppeling screen ──
        <div>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 20 }}>
            <div>
              <h2 style={{ fontSize: 16, fontWeight: 900, color: "#90caf9", letterSpacing: 2, marginBottom: 4 }}>
                KOPPELING PLANNING ↔ BEGROTING
              </h2>
              <p style={{ fontSize: 11, color: "#546e7a" }}>
                {tasks.length} taken geladen · Wijs per werkpakket de bijbehorende kostenposten toe
              </p>
            </div>
            <button
              onClick={() => { setTasks([]); setKoppelingen({}); }}
              style={{ background: "transparent", border: "1px solid #37474f", color: "#78909c", padding: "6px 14px", borderRadius: 6, cursor: "pointer", fontSize: 10, fontFamily: "inherit", letterSpacing: 1 }}
            >↩ NIEUW BESTAND</button>
          </div>

          {/* Legend */}
          <div style={{ display: "flex", gap: 16, marginBottom: 16, fontSize: 10, color: "#546e7a" }}>
            <span>🟢 = koppeling compleet (100%)</span>
            <span>🟡 = gedeeltelijk (&lt;100%)</span>
            <span>⚪ = geen koppeling</span>
          </div>

          <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
            {tasks.map(task => {
              const kops = koppelingen[task.id] || [];
              const tot = totalPct(task.id);
              const isOk = tot === 100;
              const hasAny = kops.length > 0;
              const indent = Math.min(task.level, 3) * 20;

              return (
                <div key={task.id} style={{
                  borderRadius: 8, overflow: "hidden",
                  border: `1px solid ${isOk ? "#1b5e20" : hasAny ? "#e65100" : "#1e4976"}`,
                  marginLeft: indent,
                  opacity: task.level > 2 ? 0.8 : 1,
                }}>
                  {/* Task header */}
                  <div style={{
                    display: "flex", alignItems: "center", gap: 12,
                    padding: "8px 14px",
                    background: isOk ? "rgba(46,125,50,.12)" : hasAny ? "rgba(230,81,0,.08)" : "rgba(255,255,255,.03)",
                  }}>
                    <span style={{ fontSize: 12 }}>{isOk ? "🟢" : hasAny ? "🟡" : "⚪"}</span>
                    <div style={{ flex: 1 }}>
                      <span style={{ fontSize: 12, fontWeight: task.level <= 1 ? 700 : 400, color: task.level <= 1 ? "#cfd8dc" : "#90a4ae" }}>
                        {task.name}
                      </span>
                      <span style={{ fontSize: 10, color: "#37474f", marginLeft: 10 }}>
                        {fmtDate(task.start)} → {fmtDate(task.finish)}
                      </span>
                    </div>
                    {tot > 0 && (
                      <span style={{ fontSize: 10, color: isOk ? "#4caf50" : "#ff8a65", fontWeight: 700 }}>{tot}%</span>
                    )}
                    <button
                      onClick={() => addKoppeling(task.id)}
                      style={{
                        background: "rgba(25,118,210,.15)", border: "1px solid #1e4976",
                        color: "#90caf9", padding: "3px 10px", borderRadius: 4,
                        cursor: "pointer", fontSize: 10, fontFamily: "inherit",
                      }}
                    >+ Koppel</button>
                  </div>

                  {/* Koppelingen */}
                  {kops.length > 0 && (
                    <div style={{ background: "#0a1520", padding: "8px 14px", display: "flex", flexDirection: "column", gap: 6 }}>
                      {kops.map((k, idx) => {
                        const kKosten = kostenPerHoofdstuk[k.hoofdstuk] || 0;
                        const toegewezen = kKosten * k.percentage / 100;
                        return (
                          <div key={idx} style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
                            <select
                              value={k.hoofdstuk}
                              onChange={e => updateKoppeling(task.id, idx, "hoofdstuk", e.target.value)}
                              style={{
                                background: "#0d1520", border: "1px solid #1e4976", color: "#90caf9",
                                padding: "4px 8px", borderRadius: 4, fontSize: 11, fontFamily: "inherit",
                                flex: "1 1 200px",
                              }}
                            >
                              {HOOFDSTUKKEN.map(h => (
                                <option key={h} value={h}>{CHAP_SHORT[h] || h}</option>
                              ))}
                            </select>
                            <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                              <input
                                type="number" min={0} max={100} step={5}
                                value={k.percentage}
                                onChange={e => updateKoppeling(task.id, idx, "percentage", Math.min(100, parseInt(e.target.value) || 0))}
                                style={{
                                  background: "#0d1520", border: "1px solid #1e4976", color: "#90caf9",
                                  padding: "4px 6px", borderRadius: 4, fontSize: 11, width: 56,
                                  textAlign: "right", fontFamily: "inherit",
                                }}
                              />
                              <span style={{ fontSize: 10, color: "#37474f" }}>%</span>
                            </div>
                            {kKosten > 0 && (
                              <span style={{ fontSize: 10, color: "#546e7a" }}>
                                → {euro(toegewezen)}
                              </span>
                            )}
                            <button
                              onClick={() => removeKoppeling(task.id, idx)}
                              style={{ background: "transparent", border: "none", color: "#37474f", cursor: "pointer", fontSize: 14, padding: "0 4px" }}
                            >×</button>
                          </div>
                        );
                      })}
                      {tot !== 100 && (
                        <div style={{ fontSize: 10, color: "#ff8a65" }}>
                          ⚠ Totaal is {tot}% — moet 100% zijn voor correcte berekening
                        </div>
                      )}
                    </div>
                  )}
                </div>
              );
            })}
          </div>
        </div>
      )}
    </div>
  );
}
