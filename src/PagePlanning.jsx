import { useState, useCallback } from "react";

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
  "KABEL/LEIDINGWERK ELEKTRA": "Kabelwerk",
  "MONTAGEWERKZAAMHEDEN ELEKTRA": "Montage",
  "AARDING CS - OS (op het veld)": "Aarding",
  "KABEL/LEIDINGWERK INFORMATIE DISTRIBUTIE MS": "Glasvezel",
  "STATIONS": "Stations",
  "BIJKOMENDE WERKZAAMHEDEN": "Bijkomend",
  "DIRECTIEBEHOEFTEN": "Directie",
  "OPLEVERING EN REVISIE": "Oplevering",
  "TER BESCHIKKING STELLEN": "T.b.s.",
};

// Default werkpakketten op basis van typische kabelwerkprojecten
const DEFAULT_TASKS = [
  { name: "Voorbereiding",     defaultHoofdstuk: "VOORBEREIDENDE WERKZAAMHEDEN" },
  { name: "Grondwerk",         defaultHoofdstuk: "GRONDWERK SLEUVEN" },
  { name: "Kruisingen / HDD",  defaultHoofdstuk: "KRUISINGEN" },
  { name: "Kabelwerk",         defaultHoofdstuk: "KABEL/LEIDINGWERK ELEKTRA" },
  { name: "Montagewerk",       defaultHoofdstuk: "MONTAGEWERKZAAMHEDEN ELEKTRA" },
  { name: "Verhardingsherstel",defaultHoofdstuk: "VERHARDINGEN" },
  { name: "Oplevering",        defaultHoofdstuk: "OPLEVERING EN REVISIE" },
];

const euro = v => new Intl.NumberFormat("nl-NL", { style: "currency", currency: "EUR", maximumFractionDigits: 0 }).format(v || 0);

function uid() { return Math.random().toString(36).slice(2, 8); }

function dateToInput(d) {
  if (!d) return "";
  const dd = String(d.getDate()).padStart(2, "0");
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  return `${d.getFullYear()}-${mm}-${dd}`;
}

function inputToDate(s) {
  if (!s) return null;
  const d = new Date(s + "T00:00:00");
  return isNaN(d) ? null : d;
}

// ── Stijlen ────────────────────────────────────────────────────────────────────
const S = {
  input: {
    background: "#e8eef5", border: "1px solid #1e4976", color: "#2c4a6e",
    padding: "5px 8px", borderRadius: 5, fontSize: 11, fontFamily: "inherit",
    outline: "none",
  },
  dateInput: {
    background: "#e8eef5", border: "1px solid #1e4976", color: "#1565c0",
    padding: "5px 8px", borderRadius: 5, fontSize: 11, fontFamily: "inherit",
    outline: "none", width: 130,
  },
  select: {
    background: "#e8eef5", border: "1px solid #1e4976", color: "#1565c0",
    padding: "5px 8px", borderRadius: 5, fontSize: 11, fontFamily: "inherit",
    outline: "none",
  },
  btnSmall: {
    background: "#cce0f5", border: "1px solid #1e4976",
    color: "#1565c0", padding: "4px 10px", borderRadius: 4,
    cursor: "pointer", fontSize: 10, fontFamily: "inherit", letterSpacing: 1,
  },
  btnDanger: {
    background: "transparent", border: "none",
    color: "#8aabca", padding: "4px 6px",
    cursor: "pointer", fontSize: 14, fontFamily: "inherit",
  },
  btnAdd: {
    background: "#d6e6f8", border: "1px dashed #1e4976",
    color: "#6b8caa", padding: "7px 14px", borderRadius: 6,
    cursor: "pointer", fontSize: 11, fontFamily: "inherit",
    width: "100%", textAlign: "left", marginTop: 4,
  },
};

// ── Hoofd component ────────────────────────────────────────────────────────────
export default function PagePlanning({ tasks, setTasks, koppelingen, setKoppelingen, posts }) {

  const [uploadError, setUploadError] = useState("");

  // Parse Excel planning file (our own format or MS Project export)
  const parseExcelPlanning = useCallback(async (file) => {
    setUploadError("");
    try {
      const XLSX = await import("xlsx");
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(new Uint8Array(buf), { type: "array", cellDates: true, raw: true });

      // Prefer "App Import" sheet, then "Taakoverzicht", then first sheet
      const preferred = ["App Import", "Taakoverzicht", "Taakoverzicht (MS Project)"];
      const sheetName = preferred.find(n => wb.SheetNames.includes(n)) || wb.SheetNames[0];
      const ws = wb.Sheets[sheetName];
      const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });

      // Find header row — look for "taaknaam" or "name"
      let headerRow = -1, colMap = {};
      const KW = {
        name:   ["taaknaam","task name","name","naam"],
        start:  ["begin","start","startdatum"],
        finish: ["voltooiing","finish","einde","einddatum","end"],
        level:  ["overzichtsniveau","outline level","niveau","level","l"],
        id:     ["id","nr","nummer"],
      };
      for (let i = 0; i < Math.min(10, data.length); i++) {
        const row = (data[i] || []).map(c => c ? String(c).toLowerCase().trim() : "");
        const nameCol = row.findIndex(c => KW.name.some(k => c === k || c.startsWith(k)));
        if (nameCol >= 0) {
          headerRow = i;
          row.forEach((c, j) => {
            if (KW.name.some(k => c === k || c.startsWith(k)) && colMap.name === undefined) colMap.name = j;
            if (KW.start.some(k => c === k || c.startsWith(k)) && colMap.start === undefined) colMap.start = j;
            if (KW.finish.some(k => c === k || c.startsWith(k)) && colMap.finish === undefined) colMap.finish = j;
            if (KW.level.some(k => c === k) && colMap.level === undefined) colMap.level = j;
            if (KW.id.some(k => c === k) && colMap.id === undefined) colMap.id = j;
          });
          break;
        }
      }

      if (headerRow < 0) {
        setUploadError("Geen herkende kolomheaders gevonden. Gebruik het tabblad 'App Import' uit het planningsbestand.");
        return;
      }

      function parseDate(v) {
        if (!v) return null;
        if (v instanceof Date) return isNaN(v) ? null : v;
        if (typeof v === "number" && v > 40000 && v < 60000)
          return new Date(Math.round((v - 25569) * 86400000));
        const s = String(v).trim();
        let m = s.match(/^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{4})$/);
        if (m) return new Date(+m[3], +m[2]-1, +m[1]);
        m = s.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})$/);
        if (m) return new Date(+m[1], +m[2]-1, +m[3]);
        const d = new Date(s); return isNaN(d) ? null : d;
      }

      const loaded = [];
      for (let i = headerRow + 1; i < data.length; i++) {
        const row = data[i] || [];
        const rawName = row[colMap.name];
        if (!rawName) continue;
        const name = String(rawName).trim();
        if (!name) continue;
        const start = parseDate(colMap.start !== undefined ? row[colMap.start] : null);
        const finish = parseDate(colMap.finish !== undefined ? row[colMap.finish] : null);
        if (!start || !finish) continue;
        const levelRaw = colMap.level !== undefined ? row[colMap.level] : null;
        const level = levelRaw != null ? (parseInt(levelRaw) || 0) : 0;
        const id = colMap.id !== undefined && row[colMap.id] != null
          ? String(row[colMap.id]) : String(i);
        loaded.push({ id, name, start, finish, level, pct: 0 });
      }

      if (loaded.length === 0) {
        setUploadError("Geen taken met start- en einddatum gevonden.");
        return;
      }

      // Auto-suggest koppelingen
      const suggested = {};
      loaded.forEach(task => {
        const nl = task.name.toLowerCase();
        const matches = [];
        if (nl.includes("grond") || nl.includes("sleuf") || nl.includes("ontgrav")) matches.push("GRONDWERK SLEUVEN");
        if ((nl.includes("kabel") || nl.includes("trek")) && !nl.includes("montage") && !nl.includes("glasvezel")) matches.push("KABEL/LEIDINGWERK ELEKTRA");
        if (nl.includes("montage") || nl.includes("mof") || nl.includes("eindsluiting")) matches.push("MONTAGEWERKZAAMHEDEN ELEKTRA");
        if (nl.includes("hdd") || nl.includes("boring") || nl.includes("kruising") || nl.includes("bemaling")) matches.push("KRUISINGEN");
        if (nl.includes("verharding") || nl.includes("asfalt") || nl.includes("straat") || nl.includes("uitnem")) matches.push("VERHARDINGEN");
        if (nl.includes("voorber") || nl.includes("inrichten") || nl.includes("vergunnin") || nl.includes("klic")) matches.push("VOORBEREIDENDE WERKZAAMHEDEN");
        if (nl.includes("groen") || nl.includes("inzaai") || nl.includes("berm")) matches.push("GROENVOORZIENINGEN");
        if (nl.includes("oplever") || nl.includes("revisie") || nl.includes("eindmet") || nl.includes("ibn") || nl.includes("dossier")) matches.push("OPLEVERING EN REVISIE");
        if (nl.includes("station")) matches.push("STATIONS");
        if (nl.includes("verkeersmaatregel") || nl.includes("tvm")) matches.push("VOORBEREIDENDE WERKZAAMHEDEN");
        if (matches.length > 0) {
          const pctEach = Math.round(100 / matches.length);
          suggested[task.id] = matches.map((h, i) => ({
            hoofdstuk: h,
            percentage: i === matches.length - 1 ? 100 - pctEach * (matches.length - 1) : pctEach,
          }));
        }
      });

      setTasks(loaded);
      setKoppelingen(suggested);
    } catch (err) {
      setUploadError("Fout bij inlezen: " + err.message);
    }
  }, [setTasks, setKoppelingen]);

  // Costs per hoofdstuk (for display)
  const kostenPerHoofdstuk = {};
  if (posts) {
    posts.forEach(p => {
      const k = p.hoeveelheid * p.kostprijs;
      kostenPerHoofdstuk[p.hoofdstuk] = (kostenPerHoofdstuk[p.hoofdstuk] || 0) + k;
    });
  }

  // ── Initialiseer met standaard werkpakketten als leeg ──────────────────────
  const initDefault = () => {
    const today = new Date();
    const newTasks = DEFAULT_TASKS.map((t, i) => ({
      id: uid(),
      name: t.name,
      start: new Date(today.getFullYear(), today.getMonth() + i * 1, 1),
      finish: new Date(today.getFullYear(), today.getMonth() + i * 1 + 1, 0),
      level: 0,
      pct: 0,
    }));
    setTasks(newTasks);
    // Auto-koppel op basis van defaultHoofdstuk
    const kops = {};
    newTasks.forEach((t, i) => {
      kops[t.id] = [{ hoofdstuk: DEFAULT_TASKS[i].defaultHoofdstuk, percentage: 100 }];
    });
    setKoppelingen(kops);
  };

  // ── Task CRUD ──────────────────────────────────────────────────────────────
  const addTask = (parentId) => {
    const parent = tasks.find(t => t.id === parentId);
    const level = parentId ? (parent ? parent.level + 1 : 1) : 0;
    const newTask = {
      id: uid(),
      name: level === 0 ? "Nieuw werkpakket" : "Nieuwe deeltaak",
      start: parent ? new Date(parent.start) : new Date(),
      finish: parent ? new Date(parent.finish) : new Date(Date.now() + 30 * 86400000),
      level,
      pct: 0,
      parentId: parentId || null,
    };
    if (parentId) {
      // Insert after parent and its existing children
      const parentIdx = tasks.findIndex(t => t.id === parentId);
      const children = tasks.filter(t => t.parentId === parentId);
      let insertIdx = parentIdx + 1;
      if (children.length > 0) {
        const lastChild = tasks.findLastIndex(t => t.parentId === parentId);
        insertIdx = lastChild + 1;
      }
      const newTasks = [...tasks];
      newTasks.splice(insertIdx, 0, newTask);
      setTasks(newTasks);
    } else {
      setTasks(prev => [...prev, newTask]);
    }
  };

  const updateTask = (id, field, value) => {
    setTasks(prev => prev.map(t => t.id === id ? { ...t, [field]: value } : t));
  };

  const deleteTask = (id) => {
    // Also delete children
    setTasks(prev => prev.filter(t => t.id !== id && t.parentId !== id));
    setKoppelingen(prev => {
      const next = { ...prev };
      delete next[id];
      return next;
    });
  };

  // ── Koppeling CRUD ─────────────────────────────────────────────────────────
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
      [taskId]: (prev[taskId] || []).filter((_, i) => i !== idx),
    }));
  };

  const totalPct = (taskId) => (koppelingen[taskId] || []).reduce((s, k) => s + (k.percentage || 0), 0);

  const fmtDate = d => d ? d.toLocaleDateString("nl-NL", { day: "2-digit", month: "2-digit", year: "numeric" }) : "—";

  // ── Leeg scherm ─────────────────────────────────────────────────────────────
  if (tasks.length === 0) {
    return (
      <div style={{ maxWidth: 680, margin: "50px auto", padding: "0 20px" }}>
        <div style={{ textAlign: "center", marginBottom: 28 }}>
          <div style={{ fontSize: 38, marginBottom: 12 }}>📅</div>
          <h2 style={{ fontSize: 17, fontWeight: 900, color: "#1565c0", letterSpacing: 2, marginBottom: 8 }}>UITVOERINGSPLANNING</h2>
          <p style={{ fontSize: 12, color: "#6b8caa", lineHeight: 1.9 }}>Laad een planningsbestand of maak de planning handmatig aan.</p>
        </div>
        <div
          style={{ border: "2px dashed #1e4976", borderRadius: 12, padding: "24px 20px", marginBottom: 14, background: "#f5f9fe", cursor: "pointer", transition: "all .2s" }}
          onDragOver={e => { e.preventDefault(); e.stopPropagation(); e.currentTarget.style.borderColor="#42a5f5"; }}
          onDragLeave={e => { e.currentTarget.style.borderColor="#b0cce8"; }}
          onDrop={e => { e.preventDefault(); e.stopPropagation(); e.currentTarget.style.borderColor="#b0cce8"; const f=e.dataTransfer.files[0]; if(f) parseExcelPlanning(f); }}
          onClick={() => document.getElementById("planningUploadInput").click()}
        >
          <div style={{ display: "flex", alignItems: "center", gap: 14 }}>
            <span style={{ fontSize: 26 }}>📂</span>
            <div style={{ flex: 1 }}>
              <div style={{ fontSize: 13, fontWeight: 700, color: "#1565c0", marginBottom: 3 }}>Excel planningsbestand uploaden</div>
              <div style={{ fontSize: 11, color: "#6b8caa" }}>Sleep het bestand hierheen — tabblad "App Import" wordt automatisch herkend</div>
            </div>
            <span style={{ background: "linear-gradient(135deg,#1976d2,#0d47a1)", color: "#fff", padding: "7px 16px", borderRadius: 6, fontSize: 11, fontWeight: 700, whiteSpace: "nowrap" }}>KIEZEN</span>
          </div>
        </div>
        <input id="planningUploadInput" type="file" accept=".xlsx,.xls" style={{ display: "none" }}
          onChange={e => { if (e.target.files[0]) { parseExcelPlanning(e.target.files[0]); e.target.value=""; } }} />
        {uploadError && (
          <div style={{ padding: "9px 13px", background: "rgba(230,81,0,.08)", border: "1px solid #e65100", borderRadius: 7, fontSize: 11, color: "#ff8a65", marginBottom: 12 }}>⚠ {uploadError}</div>
        )}
        <div style={{ display: "flex", alignItems: "center", gap: 10, margin: "16px 0" }}>
          <div style={{ flex: 1, height: 1, background: "#b0cce8" }} /><span style={{ fontSize: 10, color: "#8aabca" }}>OF</span><div style={{ flex: 1, height: 1, background: "#b0cce8" }} />
        </div>
        <div style={{ display: "flex", gap: 10 }}>
          <button onClick={initDefault} style={{ flex: 1, background: "#d6e6f8", border: "1px solid #1e4976", color: "#1565c0", padding: "12px 18px", borderRadius: 8, cursor: "pointer", fontSize: 11, fontWeight: 700, letterSpacing: 1, fontFamily: "inherit" }}>✦ STANDAARD WERKPAKKETTEN</button>
          <button onClick={() => addTask(null)} style={{ background: "transparent", border: "1px solid #1e4976", color: "#6b8caa", padding: "12px 18px", borderRadius: 8, cursor: "pointer", fontSize: 11, fontWeight: 700, letterSpacing: 1, fontFamily: "inherit" }}>+ LEEG BEGIN</button>
        </div>
      </div>
    );
  }

  // ── Planning editor ─────────────────────────────────────────────────────────
  const topLevelTasks = tasks.filter(t => !t.parentId && t.level === 0);

  return (
    <div style={{ maxWidth: 1100, margin: "0 auto", padding: "20px 20px 80px" }}>
      {/* Header */}
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 20 }}>
        <div>
          <h2 style={{ fontSize: 15, fontWeight: 900, color: "#1565c0", letterSpacing: 2, marginBottom: 3 }}>
            KOPPELING PLANNING ↔ BEGROTING
          </h2>
          <p style={{ fontSize: 11, color: "#6b8caa" }}>
            {tasks.length} taken · {Object.keys(koppelingen).filter(k => totalPct(k) === 100).length} volledig gekoppeld
          </p>
        </div>
        <div style={{ display: "flex", gap: 8 }}>
          <button onClick={() => addTask(null)} style={{
            ...S.btnSmall, background: "#c0d8f2", fontWeight: 700,
          }}>+ WERKPAKKET</button>
          <button onClick={() => { setTasks([]); setKoppelingen({}); }} style={S.btnSmall}>
            ↩ RESET
          </button>
        </div>
      </div>

      {/* Legend */}
      <div style={{ display: "flex", gap: 16, marginBottom: 14, fontSize: 10, color: "#6b8caa" }}>
        <span>🟢 koppeling 100%</span>
        <span>🟡 onvolledig</span>
        <span>⚪ niet gekoppeld</span>
      </div>

      {/* Task rows */}
      <div style={{ display: "flex", flexDirection: "column", gap: 5 }}>
        {tasks.map(task => {
          const kops = koppelingen[task.id] || [];
          const tot = totalPct(task.id);
          const isOk = tot === 100;
          const hasAny = kops.length > 0;
          const indent = task.level * 28;
          const isSubtask = task.level > 0;

          // Kosten toegewezen via koppelingen
          const taskKosten = kops.reduce((s, k) => {
            return s + (kostenPerHoofdstuk[k.hoofdstuk] || 0) * k.percentage / 100;
          }, 0);

          return (
            <div key={task.id} style={{
              marginLeft: indent,
              borderRadius: 8,
              border: `1px solid ${isOk ? "#1b5e20" : hasAny ? "#e65100" : "#b0cce8"}`,
              overflow: "hidden",
            }}>
              {/* Task header row */}
              <div style={{
                display: "flex", alignItems: "center", gap: 8,
                padding: "8px 12px",
                background: isOk ? "rgba(27,94,32,.08)" : hasAny ? "rgba(230,81,0,.06)" : "#f8fafc",
                flexWrap: "wrap",
              }}>
                {/* Status dot */}
                <span style={{ fontSize: 12, flexShrink: 0 }}>
                  {isOk ? "🟢" : hasAny ? "🟡" : "⚪"}
                </span>

                {/* Naam */}
                <input
                  value={task.name}
                  onChange={e => updateTask(task.id, "name", e.target.value)}
                  style={{
                    ...S.input,
                    fontWeight: isSubtask ? 400 : 700,
                    color: isSubtask ? "#4a6785" : "#2c4a6e",
                    flex: "1 1 140px",
                    minWidth: 100,
                  }}
                />

                {/* Start */}
                <input
                  type="date"
                  value={dateToInput(task.start)}
                  onChange={e => updateTask(task.id, "start", inputToDate(e.target.value))}
                  style={S.dateInput}
                />

                {/* Finish */}
                <input
                  type="date"
                  value={dateToInput(task.finish)}
                  onChange={e => updateTask(task.id, "finish", inputToDate(e.target.value))}
                  style={S.dateInput}
                />

                {/* Kosten indicator */}
                {taskKosten > 0 && (
                  <span style={{ fontSize: 10, color: "#6b8caa", whiteSpace: "nowrap" }}>
                    {euro(taskKosten)}
                  </span>
                )}

                {/* % totaal */}
                {hasAny && (
                  <span style={{ fontSize: 10, fontWeight: 700, color: isOk ? "#1b5e20" : "#ff8a65", whiteSpace: "nowrap" }}>
                    {tot}%
                  </span>
                )}

                {/* Acties */}
                <button onClick={() => addTask(task.id)} style={{ ...S.btnSmall, fontSize: 9, padding: "3px 8px" }}>
                  + subtaak
                </button>
                <button onClick={() => addKoppeling(task.id)} style={{ ...S.btnSmall, fontSize: 9, padding: "3px 8px" }}>
                  + koppel
                </button>
                <button onClick={() => deleteTask(task.id)} style={S.btnDanger} title="Verwijderen">×</button>
              </div>

              {/* Koppelingen */}
              {kops.length > 0 && (
                <div style={{ background: "#edf1f7", padding: "7px 12px 8px 44px", display: "flex", flexDirection: "column", gap: 5 }}>
                  {kops.map((k, idx) => {
                    const kKosten = kostenPerHoofdstuk[k.hoofdstuk] || 0;
                    const toegewezen = kKosten * k.percentage / 100;
                    return (
                      <div key={idx} style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap" }}>
                        <select
                          value={k.hoofdstuk}
                          onChange={e => updateKoppeling(task.id, idx, "hoofdstuk", e.target.value)}
                          style={{ ...S.select, flex: "1 1 180px" }}
                        >
                          {HOOFDSTUKKEN.map(h => (
                            <option key={h} value={h}>{CHAP_SHORT[h] || h}</option>
                          ))}
                        </select>

                        <input
                          type="number" min={0} max={100} step={5}
                          value={k.percentage}
                          onChange={e => updateKoppeling(task.id, idx, "percentage", Math.min(100, parseInt(e.target.value) || 0))}
                          style={{ ...S.input, width: 52, textAlign: "right" }}
                        />
                        <span style={{ fontSize: 10, color: "#8aabca" }}>%</span>

                        {kKosten > 0 && (
                          <span style={{ fontSize: 10, color: "#6b8caa", whiteSpace: "nowrap" }}>
                            → {euro(toegewezen)}
                          </span>
                        )}

                        <button onClick={() => removeKoppeling(task.id, idx)} style={S.btnDanger}>×</button>
                      </div>
                    );
                  })}
                  {tot !== 100 && tot > 0 && (
                    <div style={{ fontSize: 10, color: "#ff8a65" }}>
                      ⚠ Totaal is {tot}% — moet 100% zijn
                    </div>
                  )}
                </div>
              )}
            </div>
          );
        })}
      </div>

      {/* Onderaan: werkpakket toevoegen */}
      <button onClick={() => addTask(null)} style={S.btnAdd}>
        + werkpakket toevoegen
      </button>
    </div>
  );
}
