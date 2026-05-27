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
    background: "#0d1520", border: "1px solid #1e4976", color: "#cfd8dc",
    padding: "5px 8px", borderRadius: 5, fontSize: 11, fontFamily: "inherit",
    outline: "none",
  },
  dateInput: {
    background: "#0d1520", border: "1px solid #1e4976", color: "#90caf9",
    padding: "5px 8px", borderRadius: 5, fontSize: 11, fontFamily: "inherit",
    outline: "none", width: 130,
  },
  select: {
    background: "#0d1520", border: "1px solid #1e4976", color: "#90caf9",
    padding: "5px 8px", borderRadius: 5, fontSize: 11, fontFamily: "inherit",
    outline: "none",
  },
  btnSmall: {
    background: "rgba(25,118,210,.15)", border: "1px solid #1e4976",
    color: "#90caf9", padding: "4px 10px", borderRadius: 4,
    cursor: "pointer", fontSize: 10, fontFamily: "inherit", letterSpacing: 1,
  },
  btnDanger: {
    background: "transparent", border: "none",
    color: "#37474f", padding: "4px 6px",
    cursor: "pointer", fontSize: 14, fontFamily: "inherit",
  },
  btnAdd: {
    background: "rgba(25,118,210,.12)", border: "1px dashed #1e4976",
    color: "#546e7a", padding: "7px 14px", borderRadius: 6,
    cursor: "pointer", fontSize: 11, fontFamily: "inherit",
    width: "100%", textAlign: "left", marginTop: 4,
  },
};

// ── Hoofd component ────────────────────────────────────────────────────────────
export default function PagePlanning({ tasks, setTasks, koppelingen, setKoppelingen, posts }) {

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
      <div style={{ maxWidth: 680, margin: "60px auto", padding: "0 20px", textAlign: "center" }}>
        <div style={{ fontSize: 40, marginBottom: 16 }}>📅</div>
        <h2 style={{ fontSize: 18, fontWeight: 900, color: "#90caf9", letterSpacing: 2, marginBottom: 10 }}>
          UITVOERINGSPLANNING
        </h2>
        <p style={{ fontSize: 12, color: "#546e7a", lineHeight: 1.9, marginBottom: 32 }}>
          Maak de planning handmatig aan met standaard werkpakketten,<br />
          of begin met een leeg schema en voeg eigen taken toe.
        </p>
        <div style={{ display: "flex", gap: 12, justifyContent: "center" }}>
          <button onClick={initDefault} style={{
            background: "linear-gradient(135deg,#1976d2,#0d47a1)", border: "none",
            color: "#fff", padding: "12px 28px", borderRadius: 8, cursor: "pointer",
            fontSize: 12, fontWeight: 700, letterSpacing: 1, fontFamily: "inherit",
            boxShadow: "0 2px 14px rgba(25,118,210,.4)",
          }}>
            ✦ START MET STANDAARD WERKPAKKETTEN
          </button>
          <button onClick={() => addTask(null)} style={{
            background: "transparent", border: "1px solid #1e4976",
            color: "#546e7a", padding: "12px 24px", borderRadius: 8, cursor: "pointer",
            fontSize: 12, fontWeight: 700, letterSpacing: 1, fontFamily: "inherit",
          }}>
            + LEEG BEGIN
          </button>
        </div>
        <div style={{ marginTop: 32, padding: 20, background: "rgba(255,255,255,.02)", borderRadius: 10, border: "1px solid #1e4976", textAlign: "left" }}>
          <div style={{ fontSize: 10, color: "#546e7a", letterSpacing: 1, marginBottom: 12 }}>STANDAARD WERKPAKKETTEN</div>
          {DEFAULT_TASKS.map((t, i) => (
            <div key={i} style={{ display: "flex", gap: 10, marginBottom: 7, fontSize: 11, color: "#78909c", alignItems: "center" }}>
              <span style={{ color: "#1976d2", minWidth: 16 }}>▸</span>
              <span style={{ fontWeight: 700, color: "#90a4ae" }}>{t.name}</span>
              <span style={{ color: "#37474f" }}>→ {CHAP_SHORT[t.defaultHoofdstuk]}</span>
            </div>
          ))}
          <div style={{ fontSize: 10, color: "#37474f", marginTop: 10 }}>
            Datums, namen en koppelingen zijn daarna volledig aanpasbaar.
          </div>
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
          <h2 style={{ fontSize: 15, fontWeight: 900, color: "#90caf9", letterSpacing: 2, marginBottom: 3 }}>
            KOPPELING PLANNING ↔ BEGROTING
          </h2>
          <p style={{ fontSize: 11, color: "#546e7a" }}>
            {tasks.length} taken · {Object.keys(koppelingen).filter(k => totalPct(k) === 100).length} volledig gekoppeld
          </p>
        </div>
        <div style={{ display: "flex", gap: 8 }}>
          <button onClick={() => addTask(null)} style={{
            ...S.btnSmall, background: "rgba(25,118,210,.2)", fontWeight: 700,
          }}>+ WERKPAKKET</button>
          <button onClick={() => { setTasks([]); setKoppelingen({}); }} style={S.btnSmall}>
            ↩ RESET
          </button>
        </div>
      </div>

      {/* Legend */}
      <div style={{ display: "flex", gap: 16, marginBottom: 14, fontSize: 10, color: "#546e7a" }}>
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
              border: `1px solid ${isOk ? "#1b5e20" : hasAny ? "#e65100" : "#1e4976"}`,
              overflow: "hidden",
            }}>
              {/* Task header row */}
              <div style={{
                display: "flex", alignItems: "center", gap: 8,
                padding: "8px 12px",
                background: isOk ? "rgba(46,125,50,.1)" : hasAny ? "rgba(230,81,0,.07)" : "rgba(255,255,255,.02)",
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
                    color: isSubtask ? "#90a4ae" : "#cfd8dc",
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
                  <span style={{ fontSize: 10, color: "#546e7a", whiteSpace: "nowrap" }}>
                    {euro(taskKosten)}
                  </span>
                )}

                {/* % totaal */}
                {hasAny && (
                  <span style={{ fontSize: 10, fontWeight: 700, color: isOk ? "#4caf50" : "#ff8a65", whiteSpace: "nowrap" }}>
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
                <div style={{ background: "#080f18", padding: "7px 12px 8px 44px", display: "flex", flexDirection: "column", gap: 5 }}>
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
                        <span style={{ fontSize: 10, color: "#37474f" }}>%</span>

                        {kKosten > 0 && (
                          <span style={{ fontSize: 10, color: "#546e7a", whiteSpace: "nowrap" }}>
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
