// Shared helpers for the static mockups: icon rail, SVG charts, heat colours. Sample data only.
const ICONS = {
  factory: "<path d='M3 21V10l5 3v-3l5 3V7l5 3h3v11z'/><path d='M7 17h2M12 17h2M17 17h2'/>",
  daily: "<rect x='3' y='4' width='18' height='17' rx='2'/><path d='M3 9h18M8 2v4M16 2v4M7 13h4v4H7z'/>",
  qr: "<rect x='3' y='3' width='7' height='7'/><rect x='14' y='3' width='7' height='7'/><rect x='3' y='14' width='7' height='7'/><path d='M14 14h3v3h-3zM21 14v.01M14 21h.01M17 21h4v-4'/>",
  translate: "<path d='M4 5h8M8 3v2c0 4-2 7-5 9M6 9c1 2.5 3 4 6 5'/><path d='M13 21l4-9 4 9M14.5 18h5'/>",
  write: "<path d='M14 3H6a1 1 0 0 0-1 1v16a1 1 0 0 0 1 1h12a1 1 0 0 0 1-1V8z'/><path d='M14 3v5h5'/><path d='M9 13h6M9 17h4'/>",
  gear: "<circle cx='12' cy='12' r='3'/><path d='M12 2v3M12 19v3M2 12h3M19 12h3M4.9 4.9L7 7M17 17l2.1 2.1M4.9 19.1L7 17M17 7l2.1-2.1'/>",
  users: "<circle cx='9' cy='8' r='3.5'/><path d='M2.5 20c0-3.5 3-6 6.5-6s6.5 2.5 6.5 6'/><path d='M16 4.5a3.5 3.5 0 0 1 0 7M18 14c2 .6 3.5 2.6 3.5 6'/>",
  trend: "<path d='M3 3v18h18'/><path d='M7 9l4 4 3-3 5 6'/>",
  report: "<path d='M5 3h14v18H5z'/><path d='M9 17v-4M12 17V9M15 17v-6'/>",
  refresh: "<path d='M20 11a8 8 0 1 0-2.3 5.7M20 5v6h-6'/>",
  download: "<path d='M12 3v12M7 10l5 5 5-5M4 21h16'/>",
  calendar: "<rect x='3' y='5' width='18' height='16' rx='2'/><path d='M3 10h18M8 3v4M16 3v4'/>",
  arrow: "<path d='M5 12h14M13 6l6 6-6 6'/>",
  filter: "<path d='M3 5h18l-7 8v6l-4 2v-8z'/>",
};

function icon(key, size = 18) {
  return `<svg width='${size}' height='${size}' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='1.7' stroke-linecap='round' stroke-linejoin='round' aria-hidden='true'>${ICONS[key] || ""}</svg>`;
}

/** Same tiles as Components/Layout/NavMenu.razor after the 2026-10-29 menu trim. */
function renderRail(active) {
  const tile = (key, ico, label, href = "#") =>
    `<a class="rail-tile ${active === key ? "on" : ""}" href="${href}"><span class="rail-ico">${icon(ico)}</span><span class="rail-lbl">${label}</span></a>`;
  document.getElementById("rail").innerHTML = `
    <div class="rail-brand"><span class="rail-mark"></span><span>GN LAB</span></div>
    <nav class="rail-list">
      ${tile("ngrate", "trend", "NG Rate", "ng-rate.html")}
      ${tile("daily", "daily", "Daily", "daily.html")}
      ${tile("report", "report", "Report", "report.html")}
      ${tile("bmes", "factory", "BMES")}
      ${tile("qr", "qr", "QR")}
      <div class="rail-sep"></div>
      ${tile("tr", "translate", "Translate")}
      ${tile("wr", "write", "Write")}
      <div class="rail-sep"></div>
      ${tile("set", "gear", "Setting")}
      ${tile("usr", "users", "Users")}
    </nav>
    <div class="rail-user"><span class="rail-avatar">JB</span><span class="rail-role">Admin</span></div>`;
}

const fmt = (v) => (v == null ? "—" : Math.round(v).toLocaleString("en-US"));
const signed = (v) => (v > 0 ? "+" : v < 0 ? "−" : "±") + Math.abs(Math.round(v)).toLocaleString("en-US");
const pct = (v) => (v > 0 ? "+" : v < 0 ? "−" : "") + Math.abs(v).toFixed(1) + "%";
const deltaClass = (v) => (v > 0 ? "up" : v < 0 ? "down" : "flat");

/** Background tint for a PPM cell: white → amber → red as the value nears max. */
function heat(v, max) {
  if (v == null) return "";
  const t = Math.max(0, Math.min(1, v / max));
  if (t < 0.5) return "background:transparent;color:#334155";
  if (t < 0.7) return "background:#fff7ed;color:#9a3412";
  if (t < 0.85) return "background:#fed7aa;color:#9a3412";
  return "background:#fecaca;color:#991b1b;font-weight:700";
}

function sparkline(values, { w = 110, h = 32, color = "#4f46e5", fill = true, dot = true } = {}) {
  const vals = values.filter((v) => v != null);
  const min = Math.min(...vals), max = Math.max(...vals);
  const span = max - min || 1;
  const pts = values.map((v, i) => [(i / (values.length - 1)) * (w - 4) + 2, h - 3 - ((v - min) / span) * (h - 8)]);
  const line = pts.map((p, i) => `${i ? "L" : "M"}${p[0].toFixed(1)},${p[1].toFixed(1)}`).join("");
  const area = `${line}L${pts.at(-1)[0]},${h}L${pts[0][0]},${h}Z`;
  const id = "g" + Math.random().toString(36).slice(2, 8);
  const last = pts.at(-1);
  return `<svg width="${w}" height="${h}" viewBox="0 0 ${w} ${h}">
    <defs><linearGradient id="${id}" x1="0" x2="0" y1="0" y2="1"><stop offset="0" stop-color="${color}" stop-opacity=".22"/><stop offset="1" stop-color="${color}" stop-opacity="0"/></linearGradient></defs>
    ${fill ? `<path d="${area}" fill="url(#${id})"/>` : ""}
    <path d="${line}" fill="none" stroke="${color}" stroke-width="1.8" stroke-linejoin="round" stroke-linecap="round"/>
    ${dot ? `<circle cx="${last[0]}" cy="${last[1]}" r="2.8" fill="#fff" stroke="${color}" stroke-width="1.8"/>` : ""}
  </svg>`;
}

/** Stacked bars (one stack per period) + total line. series: [{name,color,values}] */
function stackedChart(labels, series, { w = 860, h = 260, target = null } = {}) {
  const pad = { l: 46, r: 14, t: 14, b: 28 };
  const totals = labels.map((_, i) => series.reduce((s, x) => s + x.values[i], 0));
  const max = Math.ceil(Math.max(...totals, target || 0) * 1.15 / 500) * 500;
  const iw = w - pad.l - pad.r, ih = h - pad.t - pad.b;
  const bw = (iw / labels.length) * 0.56;
  const x = (i) => pad.l + (iw / labels.length) * (i + 0.5);
  const y = (v) => pad.t + ih - (v / max) * ih;
  let g = "";
  for (let k = 0; k <= 4; k++) {
    const v = (max / 4) * k;
    g += `<line x1="${pad.l}" x2="${w - pad.r}" y1="${y(v)}" y2="${y(v)}" stroke="#eef2f7"/>`;
    g += `<text x="${pad.l - 8}" y="${y(v) + 4}" text-anchor="end" font-size="10.5" fill="#94a3b8">${fmt(v)}</text>`;
  }
  labels.forEach((lab, i) => {
    let base = 0;
    series.forEach((s, si) => {
      const v = s.values[i];
      const top = si === series.length - 1;
      g += `<rect x="${x(i) - bw / 2}" y="${y(base + v)}" width="${bw}" height="${y(base) - y(base + v)}" fill="${s.color}" ${top ? 'rx="3"' : ""} opacity=".9"/>`;
      base += v;
    });
    g += `<text x="${x(i)}" y="${h - 9}" text-anchor="middle" font-size="10.5" fill="#64748b">${lab}</text>`;
  });
  const line = totals.map((t, i) => `${i ? "L" : "M"}${x(i)},${y(t) - 8}`).join("");
  g += `<path d="${line}" fill="none" stroke="#0f172a" stroke-width="1.6" stroke-dasharray="0" opacity=".75"/>`;
  totals.forEach((t, i) => {
    g += `<circle cx="${x(i)}" cy="${y(t) - 8}" r="3" fill="#fff" stroke="#0f172a" stroke-width="1.5"/>`;
    if (i === totals.length - 1) g += `<text x="${x(i)}" y="${y(t) - 16}" text-anchor="middle" font-size="11" font-weight="700" fill="#0f172a">${fmt(t)}</text>`;
  });
  if (target) {
    g += `<line x1="${pad.l}" x2="${w - pad.r}" y1="${y(target)}" y2="${y(target)}" stroke="#dc2626" stroke-dasharray="5 4" stroke-width="1.2"/>`;
    g += `<text x="${w - pad.r}" y="${y(target) - 5}" text-anchor="end" font-size="10.5" font-weight="600" fill="#dc2626">Target ${fmt(target)}</text>`;
  }
  return `<svg viewBox="0 0 ${w} ${h}" width="100%" preserveAspectRatio="none" style="display:block">${g}</svg>`;
}

/** Simple multi-line chart. series: [{name,color,values,dashed}] */
function lineChart(labels, series, { w = 560, h = 200, unit = "" } = {}) {
  const pad = { l: 40, r: 12, t: 12, b: 24 };
  const all = series.flatMap((s) => s.values.filter((v) => v != null));
  const raw = Math.max(...all) * 1.1;
  const step = Math.pow(10, Math.floor(Math.log10(raw)));
  const max = Math.ceil(raw / (step / 2)) * (step / 2), min = 0;
  const iw = w - pad.l - pad.r, ih = h - pad.t - pad.b;
  const x = (i) => pad.l + (iw / (labels.length - 1)) * i;
  const y = (v) => pad.t + ih - ((v - min) / (max - min)) * ih;
  let g = "";
  for (let k = 0; k <= 3; k++) {
    const v = (max / 3) * k;
    g += `<line x1="${pad.l}" x2="${w - pad.r}" y1="${y(v)}" y2="${y(v)}" stroke="#eef2f7"/>`;
    g += `<text x="${pad.l - 6}" y="${y(v) + 4}" text-anchor="end" font-size="10" fill="#94a3b8">${v < 10 ? v.toFixed(1) : fmt(v)}${unit}</text>`;
  }
  labels.forEach((l, i) => (g += `<text x="${x(i)}" y="${h - 6}" text-anchor="middle" font-size="10" fill="#64748b">${l}</text>`));
  series.forEach((s) => {
    const d = s.values.map((v, i) => (v == null ? "" : `${i && s.values[i - 1] != null ? "L" : "M"}${x(i)},${y(v)}`)).join("");
    g += `<path d="${d}" fill="none" stroke="${s.color}" stroke-width="${s.dashed ? 1.3 : 2}" ${s.dashed ? 'stroke-dasharray="5 4"' : ""} stroke-linejoin="round"/>`;
    if (!s.dashed) s.values.forEach((v, i) => v != null && (g += `<circle cx="${x(i)}" cy="${y(v)}" r="2.6" fill="#fff" stroke="${s.color}" stroke-width="1.6"/>`));
  });
  return `<svg viewBox="0 0 ${w} ${h}" width="100%" style="display:block">${g}</svg>`;
}

/** Donut. parts: [{name,color,value}] */
function donut(parts, { size = 150, thick = 20, center = "", sub = "" } = {}) {
  const total = parts.reduce((s, p) => s + p.value, 0);
  const r = (size - thick) / 2, c = 2 * Math.PI * r;
  let off = 0, g = "";
  parts.forEach((p) => {
    const len = (p.value / total) * c;
    g += `<circle cx="${size / 2}" cy="${size / 2}" r="${r}" fill="none" stroke="${p.color}" stroke-width="${thick}" stroke-dasharray="${len - 2} ${c - len + 2}" stroke-dashoffset="${-off}" transform="rotate(-90 ${size / 2} ${size / 2})"/>`;
    off += len;
  });
  g += `<text x="50%" y="48%" text-anchor="middle" font-size="20" font-weight="700" fill="#0f172a">${center}</text>`;
  g += `<text x="50%" y="62%" text-anchor="middle" font-size="10.5" fill="#64748b">${sub}</text>`;
  return `<svg width="${size}" height="${size}" viewBox="0 0 ${size} ${size}">${g}</svg>`;
}

const TYPES = [
  { name: "Sub Assy", color: "#6366f1" },
  { name: "Main Assy", color: "#0ea5e9" },
  { name: "Function", color: "#f59e0b" },
  { name: "Final Visual", color: "#10b981" },
  { name: "OQC", color: "#ec4899" },
];
const typeColor = (name) => (TYPES.find((t) => t.name === name) || { color: "#94a3b8" }).color;

// <i data-icon="name"></i> → inline SVG icon, filled once the page has loaded.
document.addEventListener("DOMContentLoaded", () => {
  document.querySelectorAll("i[data-icon]").forEach((el) => (el.outerHTML = icon(el.dataset.icon, 15)));
});
