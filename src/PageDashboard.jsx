import { useMemo, useState } from "react";
import * as XLSX from "xlsx";

const euro = v => new Intl.NumberFormat("nl-NL", { style: "currency", currency: "EUR", maximumFractionDigits: 0 }).format(v || 0);
const euroK = v => {
  if (Math.abs(v) >= 1000000) return "€" + (v / 1000000).toFixed(1) + "M";
  if (Math.abs(v) >= 1000) return "€" + (v / 1000).toFixed(0) + "K";
  return euro(v);
};

function monthsBetween(start, end) {
  const months = [];
  const d = new Date(start.getFullYear(), start.getMonth(), 1);
  const e = new Date(end.getFullYear(), end.getMonth(), 1);
  while (d <= e) {
    months.push(new Date(d));
    d.setMonth(d.getMonth() + 1);
  }
  return months;
}

function monthKey(d) {
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
}

function monthLabel(d) {
  return d.toLocaleDateString("nl-NL", { month: "short", year: "2-digit" });
}

function taskOverlapFraction(task, month) {
  // Returns what fraction of the task falls within this month (0-1)
  const mStart = new Date(month.getFullYear(), month.getMonth(), 1);
  const mEnd = new Date(month.getFullYear(), month.getMonth() + 1, 0);
  const tStart = task.start;
  const tEnd = task.finish;
  const overlapStart = tStart > mStart ? tStart : mStart;
  const overlapEnd = tEnd < mEnd ? tEnd : mEnd;
  if (overlapEnd < overlapStart) return 0;
  const overlapDays = (overlapEnd - overlapStart) / 86400000 + 1;
  const taskDays = Math.max(1, (tEnd - tStart) / 86400000 + 1);
  return overlapDays / taskDays;
}

const CAT_LABELS = ["Stafkosten","Projectkosten","Werkterrein","Bouwwegen","Sleufloos","Grondverzet","Kabelwerk","Bemaling","Civiel (derden)","Bemaling (derden)","Sleufloos (derden)","Kabelwerk (derden)","Leveranties"];
const CAT_COLORS = ["#4fc3f7","#81c784","#ffb74d","#f06292","#9575cd","#4db6ac","#fff176","#ff8a65","#a1887f","#90a4ae","#ce93d8","#80cbc4","#e6ee9c"];

export default function PageDashboard({ posts, tasks, koppelingen, staartkosten, projectName }) {
  const [activeView, setActiveView] = useState("scurve");
  const [showTable, setShowTable] = useState(false);

  // ── Core calculation ──────────────────────────────────────────────────────
  const { months, cashflowData, totals } = useMemo(() => {
    if (!posts?.length || !tasks?.length || !Object.keys(koppelingen || {}).length) {
      return { months: [], cashflowData: [], totals: {} };
    }

    // Cost per hoofdstuk (inschrijf + intern)
    const inschrijfPerHoofdstuk = {};
    const kostenPerHoofdstuk = {};
    const catsPerHoofdstuk = {};

    posts.forEach(p => {
      const ins = p.hoeveelheid * p.inschrijfprijs;
      const kst = p.hoeveelheid * p.kostprijs;
      inschrijfPerHoofdstuk[p.hoofdstuk] = (inschrijfPerHoofdstuk[p.hoofdstuk] || 0) + ins;
      kostenPerHoofdstuk[p.hoofdstuk] = (kostenPerHoofdstuk[p.hoofdstuk] || 0) + kst;
      if (!catsPerHoofdstuk[p.hoofdstuk]) catsPerHoofdstuk[p.hoofdstuk] = Array(13).fill(0);
      p.cats.forEach((c, i) => { catsPerHoofdstuk[p.hoofdstuk][i] += p.hoeveelheid * c; });
    });

    // Determine project time range
    const validTasks = tasks.filter(t => koppelingen[t.id]?.length > 0 && t.start && t.finish);
    if (!validTasks.length) return { months: [], cashflowData: [], totals: {} };

    const allStarts = validTasks.map(t => t.start).filter(Boolean);
    const allEnds = validTasks.map(t => t.finish).filter(Boolean);
    const projectStart = new Date(Math.min(...allStarts.map(d => d.getTime())));
    const projectEnd = new Date(Math.max(...allEnds.map(d => d.getTime())));

    const months = monthsBetween(projectStart, projectEnd);

    // Per month: calculate cost distribution
    const monthlyData = months.map(month => {
      let inschrijf = 0, kosten = 0;
      const cats = Array(13).fill(0);
      const taskBreakdown = [];

      tasks.forEach(task => {
        const kops = koppelingen[task.id];
        if (!kops?.length || !task.start || !task.finish) return;
        const frac = taskOverlapFraction(task, month);
        if (frac === 0) return;

        kops.forEach(kop => {
          const pct = kop.percentage / 100;
          const hfdst = kop.hoofdstuk;
          const ins = (inschrijfPerHoofdstuk[hfdst] || 0) * pct * frac;
          const kst = (kostenPerHoofdstuk[hfdst] || 0) * pct * frac;
          inschrijf += ins;
          kosten += kst;
          const hCats = catsPerHoofdstuk[hfdst] || Array(13).fill(0);
          hCats.forEach((c, i) => { cats[i] += c * pct * frac; });
          if (ins > 0) taskBreakdown.push({ taskName: task.name, hoofdstuk: hfdst, inschrijf: ins, kosten: kst });
        });
      });

      return { month, key: monthKey(month), label: monthLabel(month), inschrijf, kosten, cats, taskBreakdown };
    });

    // Cumulative
    let cumIns = 0, cumKst = 0;
    const cumCats = Array(13).fill(0);
    const cashflowData = monthlyData.map(m => {
      cumIns += m.inschrijf;
      cumKst += m.kosten;
      m.cats.forEach((c, i) => { cumCats[i] += c; });
      return { ...m, cumInschrijf: cumIns, cumKosten: cumKst, cumCats: [...cumCats] };
    });

    // Totals
    const totalIns = cumIns;
    const totalKst = cumKst;

    return { months, cashflowData, totals: { totalIns, totalKst } };
  }, [posts, tasks, koppelingen]);

  const hasData = cashflowData.length > 0;

  // ── S-curve SVG ───────────────────────────────────────────────────────────
  const SCurve = () => {
    if (!hasData) return null;
    const W = 700, H = 280, pad = { t: 20, r: 20, b: 50, l: 70 };
    const iw = W - pad.l - pad.r, ih = H - pad.t - pad.b;
    const maxVal = Math.max(...cashflowData.map(d => d.cumInschrijf));
    const n = cashflowData.length;

    const px = (i) => pad.l + (i / (n - 1)) * iw;
    const py = (v) => pad.t + ih - (v / maxVal) * ih;

    // Y axis ticks
    const yTicks = [0, 0.25, 0.5, 0.75, 1].map(f => f * maxVal);

    const insPath = cashflowData.map((d, i) => `${i === 0 ? "M" : "L"} ${px(i)} ${py(d.cumInschrijf)}`).join(" ");
    const kstPath = cashflowData.map((d, i) => `${i === 0 ? "M" : "L"} ${px(i)} ${py(d.cumKosten)}`).join(" ");

    // Area under inschrijf
    const insArea = `${insPath} L ${px(n - 1)} ${py(0)} L ${px(0)} ${py(0)} Z`;

    return (
      <svg viewBox={`0 0 ${W} ${H}`} style={{ width: "100%", maxWidth: W }}>
        <defs>
          <linearGradient id="insGrad" x1="0" y1="0" x2="0" y2="1">
            <stop offset="0%" stopColor="#1976d2" stopOpacity="0.3" />
            <stop offset="100%" stopColor="#1976d2" stopOpacity="0.02" />
          </linearGradient>
          <linearGradient id="kstGrad" x1="0" y1="0" x2="0" y2="1">
            <stop offset="0%" stopColor="#7b1fa2" stopOpacity="0.3" />
            <stop offset="100%" stopColor="#7b1fa2" stopOpacity="0.02" />
          </linearGradient>
        </defs>

        {/* Grid */}
        {yTicks.map((v, i) => (
          <g key={i}>
            <line x1={pad.l} y1={py(v)} x2={W - pad.r} y2={py(v)} stroke="#1e4976" strokeWidth={0.5} strokeDasharray={i > 0 ? "4,4" : ""} />
            <text x={pad.l - 6} y={py(v) + 4} textAnchor="end" fontSize={9} fill="#37474f">{euroK(v)}</text>
          </g>
        ))}

        {/* Area fills */}
        <path d={insArea} fill="url(#insGrad)" />

        {/* Lines */}
        <path d={insPath} fill="none" stroke="#1976d2" strokeWidth={2.5} strokeLinejoin="round" />
        <path d={kstPath} fill="none" stroke="#7b1fa2" strokeWidth={2} strokeLinejoin="round" strokeDasharray="6,3" />

        {/* X axis labels */}
        {cashflowData.filter((_, i) => i % Math.max(1, Math.floor(n / 8)) === 0 || i === n - 1).map((d, i) => {
          const idx = cashflowData.indexOf(d);
          return (
            <text key={i} x={px(idx)} y={H - 10} textAnchor="middle" fontSize={9} fill="#546e7a">{d.label}</text>
          );
        })}

        {/* Dots at last point */}
        <circle cx={px(n - 1)} cy={py(cashflowData[n - 1].cumInschrijf)} r={4} fill="#1976d2" />
        <circle cx={px(n - 1)} cy={py(cashflowData[n - 1].cumKosten)} r={4} fill="#7b1fa2" />

        {/* Legend */}
        <line x1={W - 160} y1={18} x2={W - 135} y2={18} stroke="#1976d2" strokeWidth={2.5} />
        <text x={W - 130} y={22} fontSize={10} fill="#90caf9">Inschrijfsom</text>
        <line x1={W - 160} y1={34} x2={W - 135} y2={34} stroke="#7b1fa2" strokeWidth={2} strokeDasharray="6,3" />
        <text x={W - 130} y={38} fontSize={10} fill="#ce93d8">Interne kosten</text>
      </svg>
    );
  };

  // ── Gantt ─────────────────────────────────────────────────────────────────
  const GanttChart = () => {
    if (!hasData) return null;
    const koppeldeTasks = tasks.filter(t => koppelingen[t.id]?.length > 0 && t.start && t.finish);
    if (!koppeldeTasks.length) return null;

    const allStarts = koppeldeTasks.map(t => t.start.getTime());
    const allEnds = koppeldeTasks.map(t => t.finish.getTime());
    const minT = Math.min(...allStarts);
    const maxT = Math.max(...allEnds);
    const span = maxT - minT;

    const kostenPerHoofdstuk = {};
    posts?.forEach(p => { kostenPerHoofdstuk[p.hoofdstuk] = (kostenPerHoofdstuk[p.hoofdstuk] || 0) + p.hoeveelheid * p.kostprijs; });
    const maxTaskKosten = Math.max(...koppeldeTasks.map(t => {
      return (koppelingen[t.id] || []).reduce((s, k) => s + (kostenPerHoofdstuk[k.hoofdstuk] || 0) * k.percentage / 100, 0);
    }));

    return (
      <div style={{ overflowX: "auto" }}>
        {koppeldeTasks.map((task, idx) => {
          const left = ((task.start.getTime() - minT) / span) * 100;
          const width = Math.max(0.5, ((task.finish.getTime() - task.start.getTime()) / span) * 100);
          const taskKosten = (koppelingen[task.id] || []).reduce((s, k) => s + (kostenPerHoofdstuk[k.hoofdstuk] || 0) * k.percentage / 100, 0);
          const intensity = maxTaskKosten > 0 ? taskKosten / maxTaskKosten : 0;
          const bg = `rgba(25, ${Math.round(118 + intensity * 60)}, ${Math.round(210 - intensity * 80)}, ${0.4 + intensity * 0.5})`;

          return (
            <div key={task.id} style={{ marginBottom: 4, display: "flex", alignItems: "center", gap: 10 }}>
              <div style={{ width: 180, fontSize: 10, color: "#90a4ae", textAlign: "right", flexShrink: 0, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
                {task.name}
              </div>
              <div style={{ flex: 1, height: 22, background: "#0a1520", borderRadius: 4, position: "relative" }}>
                <div style={{
                  position: "absolute", left: `${left}%`, width: `${width}%`,
                  height: "100%", background: bg,
                  borderRadius: 3, display: "flex", alignItems: "center",
                  paddingLeft: 6, overflow: "hidden",
                  border: `1px solid rgba(255,255,255,0.1)`,
                }}>
                  {width > 5 && (
                    <span style={{ fontSize: 9, color: "#fff", whiteSpace: "nowrap" }}>{euroK(taskKosten)}</span>
                  )}
                </div>
              </div>
            </div>
          );
        })}
        <div style={{ display: "flex", alignItems: "center", gap: 10, marginTop: 8 }}>
          <div style={{ width: 180 }} />
          <div style={{ flex: 1, display: "flex", justifyContent: "space-between", fontSize: 9, color: "#37474f" }}>
            <span>{new Date(minT).toLocaleDateString("nl-NL", { month: "short", year: "numeric" })}</span>
            <span>{new Date((minT + maxT) / 2).toLocaleDateString("nl-NL", { month: "short", year: "numeric" })}</span>
            <span>{new Date(maxT).toLocaleDateString("nl-NL", { month: "short", year: "numeric" })}</span>
          </div>
        </div>
      </div>
    );
  };

  // ── Kostensoorten per maand stacked bar ───────────────────────────────────
  const KostenBar = () => {
    if (!hasData) return null;
    const maxMonth = Math.max(...cashflowData.map(d => d.inschrijf));
    const barH = 120;

    return (
      <div style={{ overflowX: "auto", paddingBottom: 8 }}>
        <div style={{ display: "flex", gap: 2, alignItems: "flex-end", minWidth: cashflowData.length * 36 }}>
          {cashflowData.map((d, i) => (
            <div key={i} style={{ display: "flex", flexDirection: "column", alignItems: "center", gap: 2, flex: 1, minWidth: 32 }}>
              <div style={{ width: "100%", height: barH, display: "flex", flexDirection: "column-reverse", gap: 0, position: "relative" }}>
                {d.cats.map((c, ci) => {
                  const h = maxMonth > 0 ? (c / maxMonth) * barH : 0;
                  return h > 0.5 ? (
                    <div key={ci} title={CAT_LABELS[ci]} style={{
                      width: "100%", height: h, background: CAT_COLORS[ci],
                      opacity: 0.85, flexShrink: 0,
                    }} />
                  ) : null;
                })}
              </div>
              <div style={{ fontSize: 8, color: "#37474f", transform: "rotate(-45deg)", transformOrigin: "top right", whiteSpace: "nowrap", marginTop: 4 }}>
                {d.label}
              </div>
            </div>
          ))}
        </div>
        {/* Legend */}
        <div style={{ display: "flex", flexWrap: "wrap", gap: "6px 14px", marginTop: 24 }}>
          {CAT_LABELS.map((l, i) => (
            <div key={i} style={{ display: "flex", alignItems: "center", gap: 5, fontSize: 9, color: "#78909c" }}>
              <div style={{ width: 10, height: 10, background: CAT_COLORS[i], borderRadius: 2 }} />
              {l}
            </div>
          ))}
        </div>
      </div>
    );
  };

  // ── Export ────────────────────────────────────────────────────────────────
  const exportXlsx = () => {
    const wb = XLSX.utils.book_new();

    // Cashflow sheet
    const rows = [
      ["CASHFLOW — " + (projectName || "Project")],
      [],
      ["Periode", "Inschrijfsom", "Interne kosten", "Marge", "Cum. Inschrijfsom", "Cum. Kosten", ...CAT_LABELS],
    ];
    cashflowData.forEach(d => {
      rows.push([
        d.label, d.inschrijf, d.kosten, d.inschrijf - d.kosten,
        d.cumInschrijf, d.cumKosten,
        ...d.cats,
      ]);
    });
    rows.push([]);
    rows.push(["TOTAAL", totals.totalIns, totals.totalKst, totals.totalIns - totals.totalKst]);

    const ws = XLSX.utils.aoa_to_sheet(rows);
    ws["!cols"] = [{ wch: 12 }, { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 16 }, { wch: 14 }, ...Array(13).fill({ wch: 13 })];
    XLSX.utils.book_append_sheet(wb, ws, "Cashflow");

    XLSX.writeFile(wb, `Cashflow_${(projectName || "project").replace(/\s+/g, "_")}.xlsx`);
  };

  const noDataMsg = (
    <div style={{ padding: 40, textAlign: "center", color: "#37474f", fontSize: 12 }}>
      <div style={{ fontSize: 32, marginBottom: 12 }}>📊</div>
      {!posts?.length ? "Laad eerst een inschrijfstaat op tab Begroting" :
        !tasks?.length ? "Laad een MS Project export op tab Planning" :
          "Maak koppelingen op tab Planning om het dashboard te vullen"}
    </div>
  );

  return (
    <div style={{ maxWidth: 1100, margin: "0 auto", padding: "20px 20px 60px" }}>
      {/* Header */}
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 20 }}>
        <div>
          <h2 style={{ fontSize: 16, fontWeight: 900, color: "#90caf9", letterSpacing: 2, marginBottom: 4 }}>DASHBOARD</h2>
          {hasData && (
            <p style={{ fontSize: 11, color: "#546e7a" }}>
              {cashflowData.length} periodes · {cashflowData[0]?.label} → {cashflowData[cashflowData.length - 1]?.label}
            </p>
          )}
        </div>
        {hasData && (
          <button onClick={exportXlsx} style={{
            background: "linear-gradient(135deg,#1976d2,#0d47a1)", border: "none",
            color: "#fff", padding: "7px 16px", borderRadius: 6, cursor: "pointer",
            fontSize: 10, fontWeight: 700, letterSpacing: 1, fontFamily: "inherit",
          }}>↓ EXPORTEER CASHFLOW</button>
        )}
      </div>

      {!hasData ? noDataMsg : (
        <>
          {/* KPIs */}
          <div style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 10, marginBottom: 20 }}>
            {[
              { lbl: "TOTALE INSCHRIJFSOM", val: euro(totals.totalIns), c: "#1976d2" },
              { lbl: "TOTALE INTERNE KOSTEN", val: euro(totals.totalKst), c: "#7b1fa2" },
              { lbl: "PROJECTDUUR", val: `${cashflowData.length} maanden`, c: "#00897b" },
            ].map((k, i) => (
              <div key={i} style={{
                background: "linear-gradient(135deg,#0d1f2d,#132436)",
                border: `1px solid ${k.c}40`, borderLeft: `3px solid ${k.c}`,
                borderRadius: 9, padding: "14px 16px",
              }}>
                <div style={{ fontSize: 8, color: "#546e7a", letterSpacing: 2, marginBottom: 5 }}>{k.lbl}</div>
                <div style={{ fontSize: 18, fontWeight: 900, color: k.c }}>{k.val}</div>
              </div>
            ))}
          </div>

          {/* View tabs */}
          <div style={{ display: "flex", gap: 3, borderBottom: "1px solid #1e4976", marginBottom: 16 }}>
            {[["scurve", "S-CURVE"], ["gantt", "GANTT + KOSTEN"], ["kostensoorten", "KOSTEN PER PERIODE"], ["tabel", "CASHFLOW TABEL"]].map(([id, lbl]) => (
              <button key={id} onClick={() => setActiveView(id)} style={{
                background: activeView === id ? "rgba(25,118,210,.15)" : "transparent",
                border: "none", borderBottom: `2px solid ${activeView === id ? "#1976d2" : "transparent"}`,
                color: activeView === id ? "#90caf9" : "#37474f",
                padding: "7px 14px", cursor: "pointer", fontSize: 10,
                fontWeight: 700, letterSpacing: 1, fontFamily: "inherit", transition: "all .15s",
              }}>{lbl}</button>
            ))}
          </div>

          {/* S-curve */}
          {activeView === "scurve" && (
            <div style={{ background: "#0a1520", borderRadius: 10, padding: 20, border: "1px solid #1e4976" }}>
              <div style={{ fontSize: 11, color: "#546e7a", marginBottom: 14, letterSpacing: 1 }}>CUMULATIEVE INSCHRIJFSOM VS. INTERNE KOSTEN</div>
              <SCurve />
            </div>
          )}

          {/* Gantt */}
          {activeView === "gantt" && (
            <div style={{ background: "#0a1520", borderRadius: 10, padding: 20, border: "1px solid #1e4976" }}>
              <div style={{ fontSize: 11, color: "#546e7a", marginBottom: 14, letterSpacing: 1 }}>
                GANTT — BALKBREEDTE = LOOPTIJD · KLEURINTENSITEIT = KOSTENNIVEAU
              </div>
              <GanttChart />
            </div>
          )}

          {/* Kostensoorten per maand */}
          {activeView === "kostensoorten" && (
            <div style={{ background: "#0a1520", borderRadius: 10, padding: 20, border: "1px solid #1e4976" }}>
              <div style={{ fontSize: 11, color: "#546e7a", marginBottom: 14, letterSpacing: 1 }}>INTERNE KOSTEN PER KOSTENSOORT PER MAAND</div>
              <KostenBar />
            </div>
          )}

          {/* Cashflow tabel */}
          {activeView === "tabel" && (
            <div style={{ overflowX: "auto" }}>
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 11 }}>
                <thead>
                  <tr style={{ background: "#0d1520" }}>
                    {["Periode", "Inschrijfsom", "Int. kosten", "Marge", "Cum. inschr.", "Cum. kosten"].map(h => (
                      <th key={h} style={{ padding: "7px 10px", color: "#37474f", fontWeight: 700, textAlign: "right", borderBottom: "1px solid #1e4976", whiteSpace: "nowrap" }}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {cashflowData.map((d, i) => (
                    <tr key={i} style={{ background: i % 2 ? "rgba(255,255,255,.02)" : "transparent", borderBottom: "1px solid #0d1520" }}>
                      <td style={{ padding: "5px 10px", color: "#90a4ae", fontWeight: 700 }}>{d.label}</td>
                      <td style={{ padding: "5px 10px", textAlign: "right", color: "#90caf9" }}>{euro(d.inschrijf)}</td>
                      <td style={{ padding: "5px 10px", textAlign: "right", color: "#ce93d8" }}>{euro(d.kosten)}</td>
                      <td style={{ padding: "5px 10px", textAlign: "right", color: d.inschrijf - d.kosten >= 0 ? "#4caf50" : "#ef5350", fontWeight: 700 }}>{euro(d.inschrijf - d.kosten)}</td>
                      <td style={{ padding: "5px 10px", textAlign: "right", color: "#546e7a" }}>{euro(d.cumInschrijf)}</td>
                      <td style={{ padding: "5px 10px", textAlign: "right", color: "#546e7a" }}>{euro(d.cumKosten)}</td>
                    </tr>
                  ))}
                  {/* Totaal */}
                  <tr style={{ background: "rgba(25,118,210,.12)", borderTop: "2px solid #1976d2" }}>
                    <td style={{ padding: "8px 10px", fontWeight: 900, color: "#90caf9" }}>TOTAAL</td>
                    <td style={{ padding: "8px 10px", textAlign: "right", fontWeight: 900, color: "#90caf9" }}>{euro(totals.totalIns)}</td>
                    <td style={{ padding: "8px 10px", textAlign: "right", fontWeight: 900, color: "#ce93d8" }}>{euro(totals.totalKst)}</td>
                    <td style={{ padding: "8px 10px", textAlign: "right", fontWeight: 900, color: totals.totalIns - totals.totalKst >= 0 ? "#4caf50" : "#ef5350" }}>{euro(totals.totalIns - totals.totalKst)}</td>
                    <td /><td />
                  </tr>
                </tbody>
              </table>
            </div>
          )}
        </>
      )}
    </div>
  );
}
