/* ══════════════════════════════════════════════════════════════════════════════
   INSTRUMENT — shell + chart runtime for the GN LAB Supporter mockup.
   Static-only: no build step, no dependencies. Open any page with file://.

   The nav model below is the mockup's stand-in for AppMenus + MenuPermissionService;
   every entry carries the channel code shown in the rail.
   ══════════════════════════════════════════════════════════════════════════ */

const NAV = [
  { sec: 'BMES' },
  {
    id: 'bmes', label: 'BMES', glyph: 'BM', group: true, items: [
      { id: 'ng-rate',   label: 'NG Rate',        code: 'NG-01', href: 'ng-rate.html' },
      { id: 'f-cost',    label: 'F-Cost',         code: 'FC-02', href: 'f-cost.html' },
      { id: 'report',    label: 'Report',         code: 'RP-03', href: '#' },
      { id: 'worker',    label: 'Worker Status',  code: 'WS-04', href: '#' },
      { id: 'bom',       label: 'BOM & Drawing',  code: 'BD-05', href: '#' },
      { id: 'lpa',       label: 'LPA',            code: 'LP-06', href: '#' },
      {
        id: 'model', label: 'Setting Model', glyph: 'SM', group: true, items: [
          { id: 'mgroup',  label: 'Model Group',   code: 'MG-07', href: '#' },
          { id: 'routing', label: 'Routing Table', code: 'RT-08', href: '#' },
          { id: 'reason',  label: 'Reason Table',  code: 'RS-09', href: '#' }
        ]
      }
    ]
  },
  { id: 'qr',    label: 'QR BAKO Data', glyph: 'QR', code: 'QR-10', href: '#' },
  { id: 'daily', label: 'Daily Report', glyph: 'DR', code: 'DR-11', href: 'index.html' },

  { sec: 'Tools' },
  {
    id: 'tda', label: 'Test Data Analysis', glyph: 'TA', group: true, items: [
      { id: 'batch',  label: 'Input Data (Batch)', code: 'BT-12', href: '#' },
      { id: 'result', label: 'Result',             code: 'RS-13', href: '#' },
      { id: 'ask',    label: 'Ask AI',             code: 'AI-14', href: 'ask-ai.html' }
    ]
  },
  {
    id: 'dtd', label: 'Daily Test Data', glyph: 'DT', group: true, items: [
      { id: 'dtd-in', label: 'Input Data', code: 'IN-15', href: '#' }
    ]
  },
  { id: 'graph',     label: 'Graph Maker', glyph: 'GM', code: 'GM-16', href: '#' },
  { id: 'translate', label: 'Translate',   glyph: 'TR', code: 'TR-17', href: '#' },
  { id: 'pcdl',      label: 'PC Download', glyph: 'PC', code: 'PC-18', href: '#' },

  { sec: 'System' },
  { id: 'setting', label: 'Setting',   glyph: 'ST', code: 'ST-19', href: '#' },
  { id: 'users',   label: 'Users',     glyph: 'US', code: 'US-20', href: 'admin-users.html' },
  { id: 'usage',   label: 'AI Usages', glyph: 'AU', code: 'AU-21', href: '#' },
  { id: 'dbq',     label: 'DB Query',  glyph: 'DB', code: 'DB-22', href: '#' }
];

const ICON = {
  rail: '<svg viewBox="0 0 16 16"><path d="M2 3h12M2 8h12M2 13h12"/></svg>',
  sun:  '<svg viewBox="0 0 16 16"><circle cx="8" cy="8" r="3.2"/><path d="M8 1v1.6M8 13.4V15M1 8h1.6M13.4 8H15M3.1 3.1l1.1 1.1M11.8 11.8l1.1 1.1M12.9 3.1l-1.1 1.1M4.2 11.8l-1.1 1.1"/></svg>',
  moon: '<svg viewBox="0 0 16 16"><path d="M13.5 9.6A6 6 0 0 1 6.4 2.5a6 6 0 1 0 7.1 7.1z"/></svg>',
  bell: '<svg viewBox="0 0 16 16"><path d="M12 6a4 4 0 1 0-8 0c0 4-1.5 5-1.5 5h11S12 10 12 6z"/><path d="M6.6 13.5a1.6 1.6 0 0 0 2.8 0"/></svg>',
  out:  '<svg viewBox="0 0 16 16"><path d="M6 14H3.5A1.5 1.5 0 0 1 2 12.5v-9A1.5 1.5 0 0 1 3.5 2H6"/><path d="M10.5 11 14 8l-3.5-3M14 8H6.5"/></svg>'
};

/* ── Shell ────────────────────────────────────────────────────────────────── */
const Shell = {
  render(opts) {
    const { active, crumb, tabs, readouts } = opts;
    // ?theme=dark forces a mode without touching the stored preference — handy for review
    const forced = new URLSearchParams(location.search).get('theme');
    document.documentElement.dataset.theme = forced || localStorage.getItem('ins-theme') || 'light';

    const rail = document.querySelector('.rail__scroll');
    if (rail) rail.innerHTML = NAV.map(n => navHTML(n, active)).join('');

    const cr = document.querySelector('.crumb');
    if (cr) cr.innerHTML = crumb.map((c, i) =>
      i === crumb.length - 1
        ? `<span class="crumb__c">${c}</span>`
        : `<span class="crumb__b">${c}</span><span class="crumb__sep">/</span>`
    ).join('');

    const ro = document.querySelector('.readouts');
    if (ro && readouts) {
      ro.innerHTML = readouts.map(r =>
        `<div class="ro"><span class="ro__k">${r.k}</span><span class="ro__v ${r.cls || ''}">${r.v}</span></div>`
      ).join('') + `<div class="stack">
          <span class="stack__av" title="Byun Jinho">BJ</span>
          <span class="stack__av" title="Nguyen T.">NT</span>
          <span class="stack__av" title="Le M.">LM</span>
          <span class="stack__av" title="+2 more">+2</span>
        </div>`;
    }

    const ts = document.querySelector('.tabs');
    if (ts && tabs) ts.innerHTML = tabs.map(t =>
      `<a class="tab ${t.on ? 'is-active' : ''}" href="${t.href || '#'}">
         <span class="tab__code">${t.code}</span><span>${t.label}</span>
         <span class="tab__x" role="button" aria-label="Close">×</span>
       </a>`
    ).join('');

    wire();
  }
};

function navHTML(n, active) {
  if (n.sec) return `<div class="navsec">${n.sec}</div>`;

  if (n.group) {
    const open = n.items.some(i => isActive(i, active));
    return `<div class="navgroup ${open ? 'is-open' : ''}">
      <button class="navlink" data-toggle>
        <span class="navlink__glyph">${n.glyph || ''}</span>
        <span class="navlink__label">${n.label}</span>
        <span class="navlink__caret"></span>
      </button>
      <div class="navgroup__body"><div class="navgroup__inner">
        ${n.items.map(i => navHTML(i, active)).join('')}
      </div></div>
    </div>`;
  }

  const on = isActive(n, active);
  return `<a class="navlink ${on ? 'is-active' : ''}" href="${n.href}">
    <span class="navlink__glyph">${n.glyph || '·'}</span>
    <span class="navlink__label">${n.label}</span>
    <span class="navlink__code">${n.code || ''}</span>
  </a>`;
}

function isActive(n, active) {
  return n.id === active || (n.items || []).some(i => isActive(i, active));
}

function wire() {
  document.querySelectorAll('[data-toggle]').forEach(b =>
    b.addEventListener('click', () => b.closest('.navgroup').classList.toggle('is-open')));

  const shell = document.querySelector('.shell');
  if (localStorage.getItem('ins-rail') === 'min') shell.classList.add('is-collapsed');

  const rb = document.querySelector('[data-rail-toggle]');
  if (rb) {
    rb.innerHTML = ICON.rail;
    rb.addEventListener('click', () => {
      shell.classList.toggle('is-collapsed');
      localStorage.setItem('ins-rail', shell.classList.contains('is-collapsed') ? 'min' : 'full');
      window.dispatchEvent(new Event('resize'));
    });
  }

  const tb = document.querySelector('[data-theme-toggle]');
  if (tb) {
    const paint = () => tb.innerHTML = document.documentElement.dataset.theme === 'dark' ? ICON.sun : ICON.moon;
    paint();
    tb.addEventListener('click', () => {
      const next = document.documentElement.dataset.theme === 'dark' ? 'light' : 'dark';
      document.documentElement.dataset.theme = next;
      localStorage.setItem('ins-theme', next);
      paint();
      Charts.redrawAll();
    });
  }

  document.querySelectorAll('[data-icon]').forEach(e => e.innerHTML = ICON[e.dataset.icon] || '');

  document.querySelectorAll('.tab__x').forEach(x =>
    x.addEventListener('click', e => { e.preventDefault(); x.closest('.tab').remove(); }));

  // chips + segmented controls behave like the real filters
  document.querySelectorAll('.chip[data-toggleable]').forEach(c =>
    c.addEventListener('click', () => c.classList.toggle('is-on')));
  document.querySelectorAll('.seg').forEach(seg =>
    seg.querySelectorAll('.seg__b').forEach(b =>
      b.addEventListener('click', () => {
        seg.querySelectorAll('.seg__b').forEach(o => o.classList.remove('is-on'));
        b.classList.add('is-on');
      })));

  Tables.wire();
  Heat.paint();
}

/* ── Colour access (tokens are the single source of truth) ────────────────── */
const css = v => getComputedStyle(document.documentElement).getPropertyValue(v).trim();
const SERIES = () => [css('--s1'), css('--s2'), css('--s3'), css('--s4'), css('--s5'), css('--s6'), css('--s7'), css('--s8')];
const RAMP = () => ['--q1','--q2','--q3','--q4','--q5','--q6','--q7'].map(css);

/* ── Charts ───────────────────────────────────────────────────────────────── */
const Charts = {
  _reg: [],
  redrawAll() { this._reg.forEach(f => f()); Heat.paint(); },

  /* Multi-series line. 2px lines, hairline grid, ≥8px end markers with a 2px
     surface ring, direct end labels, crosshair + tooltip. */
  line(sel, cfg) {
    const host = document.querySelector(sel);
    if (!host) return;
    const draw = () => {
      const W = host.clientWidth || 720, H = cfg.height || 240;
      const m = { t: 14, r: cfg.rightPad ?? 54, b: 26, l: 44 };
      const iw = W - m.l - m.r, ih = H - m.t - m.b;
      const all = cfg.series.flatMap(s => s.data);
      const lo = cfg.min ?? 0;
      const sc = niceScale(Math.max(...all) - lo, cfg.max ? cfg.max - lo : null);
      const hi = lo + sc.max;
      const x = i => m.l + (iw * i) / (cfg.labels.length - 1);
      const y = v => m.t + ih - ((v - lo) / (hi - lo)) * ih;
      const C = SERIES();

      let g = '';
      for (let v = lo; v <= hi + 1e-9; v += sc.step) {
        const yy = y(v);
        g += `<line class="grid" x1="${m.l}" y1="${yy}" x2="${m.l + iw}" y2="${yy}"/>
              <text x="${m.l - 8}" y="${yy + 3.4}" text-anchor="end">${fmt(v, cfg.dp ?? 2)}</text>`;
      }
      const step = Math.ceil(cfg.labels.length / (iw / 62));
      cfg.labels.forEach((l, i) => {
        if (i % step) return;
        g += `<text x="${x(i)}" y="${H - 8}" text-anchor="middle">${l}</text>`;
      });
      g += `<line class="axis" x1="${m.l}" y1="${m.t + ih}" x2="${m.l + iw}" y2="${m.t + ih}"/>`;

      let marks = '';
      const li = cfg.labels.length - 1;
      const ends = [];
      cfg.series.forEach((s, si) => {
        const col = s.color || C[si];
        const d = s.data.map((v, i) => `${i ? 'L' : 'M'}${x(i).toFixed(1)} ${y(v).toFixed(1)}`).join(' ');
        if (s.area) marks += `<path d="${d} L${x(li)} ${m.t + ih} L${m.l} ${m.t + ih} Z" fill="${col}" opacity=".10"/>`;
        marks += `<path class="ln" d="${d}" stroke="${col}" ${s.dash ? 'stroke-dasharray="5 4"' : ''}/>`;
        marks += `<circle class="dot-ring" cx="${x(li)}" cy="${y(s.data[li])}" r="4.5" fill="${col}"/>`;
        ends.push({ y0: y(s.data[li]), y: y(s.data[li]), t: fmt(s.data[li], cfg.dp ?? 2) + (cfg.unit || '') });
      });

      /* End labels never stack on top of each other: nudge them apart by the
         minimum needed and draw a leader back to the line end. */
      if (cfg.endLabels !== false) {
        const MIN = 13;
        ends.sort((a, b) => a.y - b.y);
        for (let i = 1; i < ends.length; i++)
          if (ends[i].y - ends[i - 1].y < MIN) ends[i].y = ends[i - 1].y + MIN;
        const over = ends.length ? ends[ends.length - 1].y - (m.t + ih) : 0;
        if (over > 0) ends.forEach(e => e.y -= over);
        ends.forEach(e => {
          if (Math.abs(e.y - e.y0) > 1)
            marks += `<path class="axis" fill="none" d="M${x(li) + 6} ${e.y0} L${x(li) + 11} ${e.y}"/>`;
          marks += `<text class="lbl" x="${x(li) + 13}" y="${e.y + 3.6}" fill="${css('--ink-2')}">${e.t}</text>`;
        });
      }

      let hot = '';
      cfg.labels.forEach((l, i) => {
        hot += `<rect x="${x(i) - iw / (cfg.labels.length - 1) / 2}" y="${m.t}"
                  width="${iw / (cfg.labels.length - 1)}" height="${ih}" fill="transparent" data-i="${i}"/>`;
      });

      host.innerHTML =
        `<svg class="chart" viewBox="0 0 ${W} ${H}" height="${H}">
           ${g}
           <line class="cross" x1="0" y1="${m.t}" x2="0" y2="${m.t + ih}"/>
           ${marks}
           <g class="hot">${hot}</g>
         </svg>
         <div class="tip"></div>`;

      const svg = host.querySelector('svg'), tip = host.querySelector('.tip'), cross = host.querySelector('.cross');
      svg.querySelectorAll('.hot rect').forEach(r => {
        r.addEventListener('mouseenter', () => {
          const i = +r.dataset.i;
          cross.setAttribute('x1', x(i)); cross.setAttribute('x2', x(i));
          cross.style.opacity = '.45';
          tip.innerHTML = `<div class="tip__t">${cfg.labels[i]}</div>` + cfg.series.map((s, si) =>
            `<div class="tip__r"><span class="tip__k" style="background:${s.color || C[si]}"></span>
             <span>${s.name}</span><span class="tip__n">${fmt(s.data[i], cfg.dp ?? 2)}${cfg.unit || ''}</span></div>`).join('');
          tip.classList.add('is-on');
          const px = (x(i) / W) * host.clientWidth;
          tip.style.left = Math.min(Math.max(px - 66, 0), host.clientWidth - 150) + 'px';
          tip.style.top = '6px';
        });
      });
      svg.addEventListener('mouseleave', () => { tip.classList.remove('is-on'); cross.style.opacity = '0'; });
    };
    draw();
    this._reg.push(draw);
    window.addEventListener('resize', debounce(draw, 120));
  },

  /* Horizontal bars: ≤24px thick, 4px rounded data-end, square at the baseline,
     2px surface gap between neighbours, value at the tip. */
  bars(sel, cfg) {
    const host = document.querySelector(sel);
    if (!host) return;
    const draw = () => {
      const W = host.clientWidth || 600;
      const labelW = cfg.labelW || 128, valW = 62;
      const rowH = 26, gap = 2;
      const H = cfg.rows.length * rowH;
      const iw = W - labelW - valW;
      const hi = cfg.max ?? niceMax(Math.max(...cfg.rows.map(r => r.v)));
      const C = SERIES();
      const th = Math.min(24, rowH - gap * 2 - 4);

      let s = '';
      cfg.rows.forEach((r, i) => {
        const yTop = i * rowH + (rowH - th) / 2;
        const w = Math.max(2, (r.v / hi) * iw);
        const col = r.color || C[cfg.slot ?? 0];
        s += `<rect x="${labelW}" y="${yTop}" width="${iw}" height="${th}" fill="${css('--panel-3')}"/>`;
        s += `<path d="${endCapPath(labelW, yTop, w, th, 4)}" fill="${col}" data-i="${i}"/>`;
        s += `<text x="${labelW - 10}" y="${yTop + th / 2 + 3.6}" text-anchor="end" class="lbl">${r.k}</text>`;
        s += `<text x="${labelW + iw + 8}" y="${yTop + th / 2 + 3.6}" fill="${css('--ink')}" font-weight="500">${fmt(r.v, cfg.dp ?? 0)}${cfg.unit || ''}</text>`;
      });

      host.innerHTML = `<svg class="chart" viewBox="0 0 ${W} ${H}" height="${H}">${s}</svg><div class="tip"></div>`;
      const tip = host.querySelector('.tip');
      host.querySelectorAll('path[data-i]').forEach(p => {
        p.addEventListener('mouseenter', e => {
          const r = cfg.rows[+p.dataset.i];
          tip.innerHTML = `<div class="tip__t">${r.k}</div>
            <div class="tip__r"><span class="tip__k" style="background:${r.color || SERIES()[cfg.slot ?? 0]}"></span>
            <span>${cfg.name || 'Value'}</span><span class="tip__n">${fmt(r.v, cfg.dp ?? 0)}${cfg.unit || ''}</span></div>
            ${r.note ? `<div class="tip__r"><span>${r.note}</span></div>` : ''}`;
          tip.classList.add('is-on');
          tip.style.left = Math.min(e.offsetX + 14, host.clientWidth - 152) + 'px';
          tip.style.top = (+p.getAttribute('y') - 6) + 'px';
        });
        p.addEventListener('mouseleave', () => tip.classList.remove('is-on'));
      });
    };
    draw();
    this._reg.push(draw);
    window.addEventListener('resize', debounce(draw, 120));
  },

  /* 12-point sparkline for stat tiles: de-emphasis hue, last segment in accent. */
  sparks() {
    document.querySelectorAll('[data-spark]').forEach(el => {
      const v = el.dataset.spark.split(',').map(Number);
      const W = 78, H = 22, lo = Math.min(...v), hi = Math.max(...v) || 1;
      const x = i => (W * i) / (v.length - 1);
      const y = n => H - 2 - ((n - lo) / (hi - lo || 1)) * (H - 4);
      const d = v.map((n, i) => `${i ? 'L' : 'M'}${x(i).toFixed(1)} ${y(n).toFixed(1)}`).join(' ');
      const tail = `M${x(v.length - 2).toFixed(1)} ${y(v[v.length - 2]).toFixed(1)} L${x(v.length - 1).toFixed(1)} ${y(v[v.length - 1]).toFixed(1)}`;
      el.innerHTML = `<svg class="spark" viewBox="0 0 ${W} ${H}" width="${W}" height="${H}">
        <path class="sl" d="${d}"/><path class="sl sl--sig" d="${tail}"/>
        <circle cx="${x(v.length - 1)}" cy="${y(v[v.length - 1])}" r="2.4" fill="${css('--signal')}"/></svg>`;
    });
    this._reg.push(() => this.sparks());
  }
};

/* rectangle with a rounded data-end, square at the baseline (left) */
function endCapPath(x, y, w, h, r) {
  r = Math.min(r, w, h / 2);
  return `M${x} ${y} H${x + w - r} A${r} ${r} 0 0 1 ${x + w} ${y + r} V${y + h - r} A${r} ${r} 0 0 1 ${x + w - r} ${y + h} H${x} Z`;
}

/* ── Heatmap matrix ───────────────────────────────────────────────────────── */
const Heat = {
  paint() {
    const ramp = RAMP();
    document.querySelectorAll('[data-heat]').forEach(tbl => {
      const stops = tbl.dataset.heat.split(',').map(Number);   // 6 thresholds → 7 steps
      tbl.querySelectorAll('td[data-v]').forEach(td => {
        const v = +td.dataset.v;
        let step = 0;
        while (step < stops.length && v >= stops[step]) step++;
        td.style.background = ramp[step];
        // ink or white by the fill's own luminance — the ramp inverts between modes
        td.style.color = onFill(ramp[step]);
        td.title = `${td.dataset.k || ''} — ${td.textContent.trim()}`;
      });
    });
    document.querySelectorAll('[data-ramp]').forEach(el => {
      el.innerHTML = ramp.map(c => `<i style="background:${c}"></i>`).join('');
    });
  }
};

/* ── Tables ───────────────────────────────────────────────────────────────── */
const Tables = {
  wire() {
    document.querySelectorAll('.dt th.is-sortable').forEach(th => {
      th.addEventListener('click', () => {
        const tbl = th.closest('table'), idx = [...th.parentNode.children].indexOf(th);
        const dir = th.classList.contains('is-sorted') && th.dataset.dir === 'asc' ? 'desc' : 'asc';
        tbl.querySelectorAll('th').forEach(o => { o.classList.remove('is-sorted'); o.removeAttribute('data-dir'); });
        th.classList.add('is-sorted'); th.dataset.dir = dir;
        const body = tbl.querySelector('tbody[data-sortable]');
        if (!body) return;
        [...body.rows]
          .sort((a, b) => {
            const av = cellVal(a.cells[idx]), bv = cellVal(b.cells[idx]);
            return (av > bv ? 1 : av < bv ? -1 : 0) * (dir === 'asc' ? 1 : -1);
          })
          .forEach(r => body.appendChild(r));
      });
    });
  }
};
const cellVal = c => {
  const t = (c?.textContent || '').replace(/[,%\s]/g, '');
  return isNaN(parseFloat(t)) ? (c?.textContent || '').trim().toLowerCase() : parseFloat(t);
};

/* ── Utils ────────────────────────────────────────────────────────────────── */
const fmt = (v, dp = 2) => v.toLocaleString(undefined, { minimumFractionDigits: dp, maximumFractionDigits: dp });

/* Axis ticks land on clean numbers (1 / 2 / 2.5 / 5 × 10^k), never on range/n. */
function niceScale(span, forced) {
  if (span <= 0) return { max: forced || 1, step: (forced || 1) / 4 };
  const raw = (forced || span) / 5;
  const e = Math.pow(10, Math.floor(Math.log10(raw)));
  const n = raw / e;
  const step = (n <= 1 ? 1 : n <= 2 ? 2 : n <= 2.5 ? 2.5 : n <= 5 ? 5 : 10) * e;
  return { max: forced || Math.ceil(span / step) * step, step };
}
function niceMax(v) { return niceScale(v).max; }
function debounce(fn, ms) { let t; return () => { clearTimeout(t); t = setTimeout(fn, ms); }; }

/* Label inside a coloured fill: pick whichever of ink / white clears contrast. */
function onFill(hex) {
  const p = hex.replace('#', '').match(/../g);
  if (!p) return '#0f1319';
  const [r, g, b] = p.map(h => {
    const c = parseInt(h, 16) / 255;
    return c <= 0.03928 ? c / 12.92 : Math.pow((c + 0.055) / 1.055, 2.4);
  });
  const L = 0.2126 * r + 0.7152 * g + 0.0722 * b;
  return (L + 0.05) / 0.05 > 1.05 / (L + 0.05) ? '#0f1319' : '#ffffff';
}
