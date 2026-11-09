// Model Hub mockups — shared sample data and section renderers. Static, sample data only.
// Group tree = model_groups of E:\db\CompanyGlossary\company-glossary.db (2026-11-09 snapshot).
const GROUPS = [
  { name: "MODULE", children: [
    { name: "MSM", models: ["MSM-S711B", "MSM-S926U", "MSM-S921B", "MSM-S936U", "MSM-S931B", "MSM-G778B", "MSM-X316B"] },
    { name: "SI-SM", models: ["SI-SM-A546B", "SI-SM-M546B", "SI-SM-S908U"] },
  ] },
  { name: "SPEAKER", children: [
    { name: "BRS", models: ["BRS-161011S08ZZ", "BRS-161016S08ZZ", "BRS-181389S08ZZ", "BRS-161296S08", "BRS-161298S08ZZ", "BRS-181390S04ZZ", "BRS-201506S08ZZ"] },
    { name: "MSU", models: ["MSU-L16S10-24-ZZ", "MSU-L20S15-07-ZZ", "MSU-L15S10-20-ZZ"] },
  ] },
  { name: "UNIT", models: ["TIU-C11-08", "TIU-C11-18-ZZ", "TIU-L5S3-01", "TIU-L7S6-01-ZZ", "TIU-C11-20-ZZ"] },
  { name: "TWS", models: ["TIM-318", "TIM-338"] },
];

const PHASES = ["Proto", "EVT", "DVT", "PVT", "MP"];
const CUSTOMERS = ["Samsung MX", "Samsung VD", "Harman", "LG Innotek", "Bose"];
const OWNERS = ["J. Byun", "H. Kim", "S. Lee", "T. Nguyen", "M. Park", "D. Tran"];
const LINES = ["Line 1 (Day)", "Line 2 (Day)", "Line 3 (Night)", "Line 5 (Day)"];
const TODAY = "2026-11-09";

function hash(s) { let h = 2166136261; for (const c of s) { h ^= c.charCodeAt(0); h = Math.imul(h, 16777619); } return h >>> 0; }
function rng(seed) { let x = seed || 1; return () => ((x = Math.imul(x ^ (x >>> 15), 2246822519) + 0x6d2b79f5 >>> 0) % 10000) / 10000; }
const pick = (r, arr) => arr[Math.floor(r() * arr.length)];
function addDays(d, n) { const t = new Date(d + "T00:00:00Z"); t.setUTCDate(t.getUTCDate() + n); return t.toISOString().slice(0, 10); }

function allModels() {
  const out = [];
  const walk = (g, path) => {
    for (const m of g.models || []) out.push({ name: m, path: [...path, g.name] });
    for (const c of g.children || []) walk(c, [...path, g.name]);
  };
  GROUPS.forEach((g) => walk(g, []));
  return out;
}

/** Deterministic sample record for one model. */
function modelRecord(name) {
  const r = rng(hash(name));
  const entry = allModels().find((m) => m.name === name) || { path: [] };
  const phaseIdx = Math.floor(r() * PHASES.length);
  // Phases are 38 days apart; today falls inside the current phase.
  const start = addDays(TODAY, -(phaseIdx * 38 + 3 + Math.floor(r() * 25)));
  const ms = PHASES.map((p, i) => {
    const from = addDays(start, i * 38), to = addDays(from, 30);
    return { phase: p, from, to, state: i < phaseIdx ? "done" : i === phaseIdx ? "now" : "next" };
  });
  const issues = Array.from({ length: 4 + Math.floor(r() * 6) }, (_, i) => {
    const sev = pick(r, ["Critical", "Major", "Major", "Minor", "Minor"]);
    const open = r() < 0.4;
    return {
      id: `ISS-${(hash(name) % 900) + 100}-${String(i + 1).padStart(2, "0")}`,
      date: addDays(TODAY, -Math.floor(r() * 120)),
      title: pick(r, ["Buzz noise at 1 kHz on function test", "Mesh adhesive overflow", "Acoustic SPL low −2 dB", "FPCB solder crack after drop", "Frame warpage over 0.15 mm", "Magnet misalignment", "Gasket missing in visual", "Rub & buzz NG rate over 3,000 PPM"]),
      process: pick(r, ["SUB", "MAIN", "FUNCTION", "VISUAL"]),
      severity: sev, status: open ? "Open" : "Closed",
      owner: pick(r, OWNERS), due: addDays(TODAY, Math.floor(r() * 30) - 10),
    };
  }).sort((a, b) => b.date.localeCompare(a.date));
  const history = Array.from({ length: 6 + Math.floor(r() * 5) }, () => ({
    date: addDays(TODAY, -Math.floor(r() * 220)),
    phase: pick(r, PHASES.slice(0, phaseIdx + 1)),
    title: pick(r, ["Sample build 300 pcs shipped to customer", "Tooling T1 → T2 change", "Adhesive changed to 3M 9448", "Process flow updated (add plasma)", "Line setup moved to Line 3", "Jig revision B approved", "BOM rev C released", "Golden sample registered"]),
    by: pick(r, OWNERS), kind: pick(r, ["Build", "Change", "Tooling", "BOM", "Process"]),
  })).sort((a, b) => b.date.localeCompare(a.date));
  const reviews = Array.from({ length: 3 + Math.floor(r() * 4) }, () => ({
    date: addDays(TODAY, -Math.floor(r() * 200)),
    kind: pick(r, ["Design Review", "PFMEA", "Customer Review", "Build Readiness", "Report Review"]),
    result: pick(r, ["Pass", "Pass", "Conditional", "Fail"]),
    title: pick(r, ["DR2 — acoustic structure", "PFMEA update for MAIN segment", "Customer audit follow-up", "PVT build readiness check", "Weekly NG review"]),
    actions: Math.floor(r() * 5), by: pick(r, OWNERS),
  })).sort((a, b) => b.date.localeCompare(a.date));
  return {
    name, path: entry.path,
    customer: pick(r, CUSTOMERS), owner: pick(r, OWNERS), line: pick(r, LINES),
    build: pick(r, ["PV1", "PV2", "MP1", "EVT2", "DVT1"]),
    phase: PHASES[phaseIdx], phaseIdx, milestones: ms,
    sop: ms[4].from, processCount: 12 + Math.floor(r() * 18), materialCount: 8 + Math.floor(r() * 14),
    ngPpm: Math.round(800 + r() * 4200), yieldPct: (95 + r() * 4.6).toFixed(1),
    issues, history, reviews, seed: hash(name),
  };
}

// ── Production Input plan (StockPlanTab) — same stages / Plan · Actual · Gap rows ──────────────
const STAGES = ["SPK Assy (Sub 12h)", "Lower Assy (Main 12h)", "Function OK", "Visual OK", "Send OQC", "Export to KR (S-code)", "Export to KR (Free)", "Customer ship"];
const WD = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];

function planTable(rec, { days = 14, compact = false } = {}) {
  const r = rng(rec.seed + 7);
  const dates = Array.from({ length: days }, (_, i) => addDays("2026-11-02", i));
  const stages = compact ? STAGES.slice(0, 5) : STAGES;
  const head = dates.map((d) => {
    const w = new Date(d + "T00:00:00Z").getUTCDay();
    return `<th class="${d === TODAY ? "today" : w === 0 || w === 6 ? "we" : ""}">${d.slice(5)}<br>${WD[w]}</th>`;
  }).join("");
  let rows = `<tr><td class="stage">Material arrival</td><td class="kind">event</td><td class="tot"></td>${dates.map((d, i) => `<td class="ev">${i === 1 ? "FPCB 5k" : i === 8 ? "Frame 3k" : ""}</td>`).join("")}</tr>`;
  stages.forEach((st, si) => {
    const base = 600 + Math.floor(r() * 500);
    const plan = dates.map((d) => { const w = new Date(d + "T00:00:00Z").getUTCDay(); return w === 0 ? 0 : Math.round(base * (0.8 + r() * 0.4) / 10) * 10; });
    const act = dates.map((d, i) => (d <= TODAY && plan[i] ? Math.round(plan[i] * (0.85 + r() * 0.25)) : null));
    const tp = plan.reduce((a, b) => a + b, 0), ta = act.reduce((a, b) => a + (b || 0), 0);
    const ptd = plan.filter((_, i) => dates[i] <= TODAY).reduce((a, b) => a + b, 0);
    rows += `<tr class="first"><td class="stage" rowspan="3">${st}<small>${si < 5 ? "Actual = Daily NG input" : "Actual typed here"}</small></td><td class="kind">Plan</td><td class="tot">${fmt(tp)}</td>${plan.map((v) => `<td>${v ? fmt(v) : ""}</td>`).join("")}</tr>`;
    rows += `<tr><td class="kind">Actual</td><td class="tot">${fmt(ta)}</td>${act.map((v) => `<td>${v == null ? "" : fmt(v)}</td>`).join("")}</tr>`;
    rows += `<tr class="gap"><td class="kind">Gap</td><td class="tot">${signed(ta - ptd)}</td>${act.map((v, i) => (v == null ? "<td></td>" : `<td class="${v - plan[i] < 0 ? "neg" : "pos"}">${signed(v - plan[i])}</td>`)).join("")}</tr>`;
  });
  return `<div class="tbl-wrap"><table class="plan"><thead><tr><th class="stage">Stage</th><th></th><th>Total</th>${head}</tr></thead><tbody>${rows}</tbody></table></div>`;
}

/** Phase gantt: Proto → MP bars on a month axis. */
function milestoneGantt(rec) {
  const start = rec.milestones[0].from, end = addDays(rec.milestones[4].to, 10);
  const span = (new Date(end) - new Date(start)) / 864e5;
  const x = (d) => (((new Date(d) - new Date(start)) / 864e5) / span) * 100;
  const todayX = Math.max(0, Math.min(100, x(TODAY)));
  return `<div class="gantt">${rec.milestones.map((m) => `
    <div class="g-row"><span class="g-lbl">${m.phase}</span><div class="g-track">
      <div class="g-bar ${m.state}" style="left:${x(m.from)}%;width:${x(m.to) - x(m.from)}%"><span>${m.from.slice(5)} – ${m.to.slice(5)}</span></div>
      <div class="g-today" style="left:${todayX}%"></div></div></div>`).join("")}
    <div class="g-axis"><span>${start}</span><span class="g-now" style="left:${todayX}%">Today</span><span>${end}</span></div></div>`;
}

function phaseStepper(rec) {
  return `<div class="stepper">${PHASES.map((p, i) => `<div class="step ${i < rec.phaseIdx ? "done" : i === rec.phaseIdx ? "now" : ""}"><i>${i < rec.phaseIdx ? "✓" : i + 1}</i><span>${p}</span></div>`).join("<b></b>")}</div>`;
}

function infoGrid(rec) {
  const f = (k, v) => `<div class="kv"><span>${k}</span><b>${v}</b></div>`;
  return `<div class="kvs">${f("Model", rec.name)}${f("Group", rec.path.join(" › "))}${f("Customer", rec.customer)}${f("Build", rec.build)}${f("Phase", rec.phase)}${f("PM owner", rec.owner)}${f("Line", rec.line)}${f("SOP target", rec.sop)}${f("Processes", rec.processCount)}${f("Materials (BOM)", rec.materialCount)}${f("NG rate (30 d)", fmt(rec.ngPpm) + " PPM")}${f("Yield (30 d)", rec.yieldPct + "%")}</div>`;
}

function kpiRow(rec) {
  const open = rec.issues.filter((i) => i.status === "Open").length;
  const t = (l, v, s) => `<div class="card kpi"><div class="kpi-label">${l}</div><div class="kpi-value">${v}<small>${s}</small></div></div>`;
  return `<div class="grid g-4">${t("Phase", rec.phase, rec.build)}${t("Open issues", open, "/ " + rec.issues.length)}${t("NG rate (30 d)", fmt(rec.ngPpm), "PPM")}${t("SOP target", rec.sop.slice(5), rec.sop.slice(0, 4))}</div>`;
}

const sevCls = (s) => ({ Critical: "crit", Major: "maj", Minor: "min" }[s]);
const resCls = (s) => ({ Pass: "ok", Conditional: "cond", Fail: "fail" }[s]);

function issueTable(rec) {
  return `<div class="tbl-wrap"><table class="t"><thead><tr><th class="l">ID</th><th class="l">Date</th><th class="l">Title</th><th class="l">Process</th><th class="l">Severity</th><th class="l">Owner</th><th class="l">Due</th><th class="l">Status</th></tr></thead><tbody>
    ${rec.issues.map((i) => `<tr><td class="l mono">${i.id}</td><td class="l">${i.date}</td><td class="l">${i.title}</td><td class="l"><span class="ptb ${i.process === "FUNCTION" ? "func" : i.process === "VISUAL" ? "vis" : i.process.toLowerCase()}">${i.process}</span></td><td class="l"><span class="sev ${sevCls(i.severity)}">${i.severity}</span></td><td class="l">${i.owner}</td><td class="l ${i.status === "Open" && i.due < TODAY ? "late" : ""}">${i.due}</td><td class="l"><span class="st ${i.status === "Open" ? "open" : "closed"}">${i.status}</span></td></tr>`).join("")}
  </tbody></table></div>`;
}

function historyList(rec) {
  return `<ol class="tl">${rec.history.map((h) => `<li><span class="tl-dot"></span><div class="tl-date">${h.date}</div><div class="tl-body"><b>${h.title}</b><div class="muted"><span class="tag">${h.kind}</span> <span class="tag ph">${h.phase}</span> · ${h.by}</div></div></li>`).join("")}</ol>`;
}

function reviewTable(rec) {
  return `<div class="tbl-wrap"><table class="t"><thead><tr><th class="l">Date</th><th class="l">Type</th><th class="l">Subject</th><th class="l">Result</th><th>Action items</th><th class="l">Chair</th></tr></thead><tbody>
    ${rec.reviews.map((v) => `<tr><td class="l">${v.date}</td><td class="l">${v.kind}</td><td class="l">${v.title}</td><td class="l"><span class="res ${resCls(v.result)}">${v.result}</span></td><td>${v.actions}</td><td class="l">${v.by}</td></tr>`).join("")}
  </tbody></table></div>`;
}

/** All history kinds on one date line (used by the dossier and the board drawer). */
function unifiedEvents(rec) {
  return [
    ...rec.history.map((h) => ({ date: h.date, type: "dev", title: h.title, meta: `${h.kind} · ${h.phase} · ${h.by}` })),
    ...rec.issues.map((i) => ({ date: i.date, type: "issue", title: `${i.id} ${i.title}`, meta: `${i.severity} · ${i.status} · ${i.owner}` })),
    ...rec.reviews.map((v) => ({ date: v.date, type: "review", title: `${v.kind}: ${v.title}`, meta: `${v.result} · ${v.actions} action items · ${v.by}` })),
  ].sort((a, b) => b.date.localeCompare(a.date));
}

function unifiedTimeline(rec, only = null) {
  const ev = unifiedEvents(rec).filter((e) => !only || only.has(e.type));
  const label = { dev: "Dev", issue: "Issue", review: "Review" };
  let month = "";
  return `<ol class="tl uni">${ev.map((e) => {
    const m = e.date.slice(0, 7), mh = m !== month ? `<li class="tl-month">${m}</li>` : "";
    month = m;
    return `${mh}<li class="t-${e.type}"><span class="tl-dot"></span><div class="tl-date">${e.date.slice(5)}</div><div class="tl-body"><span class="kind-tag ${e.type}">${label[e.type]}</span> <b>${e.title}</b><div class="muted">${e.meta}</div></div></li>`;
  }).join("")}</ol>`;
}

/** Left group tree. onPick(modelName) is called on click; counts show model totals. */
function renderTree(el, selected, onPick, { filter = "" } = {}) {
  const q = filter.trim().toLowerCase();
  const node = (g, depth) => {
    const models = (g.models || []).filter((m) => !q || m.toLowerCase().includes(q));
    const kids = (g.children || []).map((c) => node(c, depth + 1)).filter(Boolean);
    if (q && !models.length && !kids.length) return "";
    const count = allModelsOf(g).length;
    return `<div class="tg d${depth}"><div class="tg-h"><span class="caret">▾</span>${g.name}<span class="cnt">${count}</span></div>
      ${kids.join("")}${models.map((m) => `<a class="tm ${m === selected ? "on" : ""}" data-m="${m}"><span class="ph-dot p${modelRecord(m).phaseIdx}"></span>${m}</a>`).join("")}</div>`;
  };
  el.innerHTML = GROUPS.map((g) => node(g, 0)).join("");
  el.querySelectorAll(".tm").forEach((a) => a.addEventListener("click", () => onPick(a.dataset.m)));
  el.querySelectorAll(".tg-h").forEach((h) => h.addEventListener("click", () => h.parentElement.classList.toggle("shut")));
}
function allModelsOf(g) { return [...(g.models || []), ...(g.children || []).flatMap(allModelsOf)]; }

const PHASE_LEGEND = `<div class="legend ph-legend">${PHASES.map((p, i) => `<span><i class="ph-dot p${i}"></i>${p}</span>`).join("")}</div>`;
