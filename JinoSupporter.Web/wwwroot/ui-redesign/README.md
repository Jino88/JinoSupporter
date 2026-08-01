# INSTRUMENT — UI redesign mockup

Static HTML redesign of the whole GN LAB Supporter front end. No build step, no
dependencies, nothing wired to the app.

Two things live here:

1. **A working page.** `New Ver` in the nav (route `/new-ver`) is a real Blazor
   page — real menus from `AppMenus` + `MenuPermissionService`, real connected-user
   readouts, and a Users screen that reads and writes the same rows as
   `/admin/users`. Not a mockup; the writes are real.
2. **Design reference.** The static HTML below, for the screens not ported yet.
   Open them at `http://localhost:5050/ui-redesign/index.html`, or straight off
   disk (`file://` works — no build step, no dependencies).

### Two stylesheets, one source

| File | Used by | Notes |
|---|---|---|
| `assets/instrument.css` | the static reference pages | the source of truth — edit this one |
| `assets/instrument.scoped.css` | `/new-ver` inside the app | **generated**; every selector scoped under `.ins` |

The app loads Bootstrap globally, and this design system defines its own `.btn`,
`.badge`, `.row` and `.small` — plus a `*` reset and `:root` variables. Loading it
raw would restyle every existing page. `tools/scope-css.js` rewrites every
selector under a single `.ins` root, so `.ins .btn` (0,2,0) beats Bootstrap's
`.btn` (0,1,0) inside the container and changes nothing outside it. Section 23 of
`instrument.css` restates the handful of properties Bootstrap still lands
(`border-radius`, the `.row > *` grid columns) — delete it when Bootstrap goes.

**After editing `instrument.css`, re-run `node tools/scope-css.js`.** The scoped
file is generated and must never be hand-edited.

```
JinoSupporter.Web/wwwroot/ui-redesign/
  index.html          Daily Report      — dashboard archetype (hero, tiles, trend, heatmap, tree table)
  ng-rate.html        NG Rate           — heavy report archetype (setup deck, hierarchy table, run log)
  f-cost.html         F-Cost            — cost archetype (trend vs budget, ranked bars, wide table)
  ask-ai.html         Ask AI            — conversational archetype (chat + evidence rail)
  admin-users.html    Users             — admin archetype (list + detail form + permission matrix)
  login.html          Sign in           — unauthenticated archetype
  assets/instrument.css   the whole design system
  assets/instrument.js    shell renderer + chart runtime
```

Append `?theme=dark` to any page to force dark mode, or use the moon/sun button in
the top bar (the choice is remembered in `localStorage`).

---

## The concept

**The app is a measurement instrument.** A dark machined chassis — the rail and the
top bar — holds an illuminated data plate, which is the work surface. Everything
follows from that:

- **Squared, not rounded.** Radius is 0 almost everywhere. Separation is done with
  hairlines, not shadows and not rounded cards.
- **One signal colour.** Amber `#e3661f` marks the primary action, the active
  route, and rising magnitude. Nothing else competes with it.
- **Channel codes.** Every menu, tab, and panel carries a mono code (`NG-01`,
  `FC-02`). It makes 22 menus addressable in conversation and on the phone.
- **Numbers are the loudest thing on screen.** IBM Plex Mono, tabular figures,
  right-aligned. Chrome recedes; data does not.
- **Density is a feature.** 27px table rows, 29px controls, 13px base. These
  screens are read for eight hours, not glanced at.
- **Tables are spreadsheet grids.** They get screenshotted and pasted into
  reports and chat, so every cell is boxed and the header is a filled band —
  a crop of a table still reads as a table. See below.

### Tables are built to be captured

Data tables use a full Excel-style grid rather than the app's hairline styling:

- every cell has a right and bottom border (`--tbl-line`), the table a top and left one
- the header is a filled grey band (`--tbl-head`), bold, centred, with a heavier
  underline (`--tbl-line-hard`)
- column-group dividers and the total rule use that same heavier line, the way a
  manual border in a sheet does
- group rows and total rows are bold banded rows, not colour-only
- nothing meaningful depends on hover, and nothing depends on colour alone —
  every status cell carries a glyph and a word

`.dt--banded` adds Excel-style banded rows for tables long enough to lose your
place in; it is off by default because banding fights the heat-scale cells.
The heatmap (`.heat`) uses the same grid, so it reads as a conditional-format
colour scale rather than a chart.

### Type

| Role | Face | Notes |
|---|---|---|
| UI / headings | **Archivo** (variable, `wdth` axis) | headings run slightly expanded; industrial grotesque |
| Numbers, codes, labels | **IBM Plex Mono** | all tabular columns, all micro-labels, the run log |

Both load from Google Fonts with a local fallback stack (`Segoe UI` / `Consolas`),
so an offline factory PC degrades gracefully rather than breaking.

---

## Tokens

All colour lives in CSS custom properties at the top of `instrument.css`. Two
scopes: `:root` (light) and `[data-theme="dark"]`. The chassis tokens (`--ch-*`)
are deliberately **not** themed — the housing is dark in both modes.

| Group | Tokens |
|---|---|
| Chassis | `--ch-000…--ch-400`, `--ch-line`, `--ch-ink`, `--ch-ink-2`, `--ch-ink-3` |
| Work surface | `--plane`, `--panel`, `--panel-2`, `--panel-3`, `--line`, `--line-2` |
| Ink | `--ink`, `--ink-2`, `--ink-3` |
| Signal | `--signal`, `--signal-hot`, `--signal-lit`, `--signal-wash`, `--signal-line` |
| Status | `--ok`, `--warn`, `--serious`, `--crit` (+ `-wash` variants) |
| Series | `--s1…--s8` |
| Sequential ramp | `--q1…--q7` |
| Chart chrome | `--grid`, `--axis` |

Re-theming the whole app is editing those two blocks. Nothing else hard-codes a colour.

---

## Chart colour — validated, not eyeballed

The palette was run through the data-viz validator (six checks: lightness band,
chroma floor, CVD separation, normal-vision floor, contrast vs surface) in **both**
modes. Results:

| Set | Mode | Surface | Result |
|---|---|---|---|
| 8 categorical slots, adjacent pairs | light | `#ffffff` | PASS (worst CVD ΔE 7.2, normal-vision 19.6) |
| 8 categorical slots, adjacent pairs | dark | `#171b21` | PASS (worst CVD ΔE 6.9, normal-vision 19.3) |
| first 3 slots, all pairs | light / dark | — | PASS |
| sequential ramp `--q1…--q7` | light | `#ffffff` | monotone L, all ΔL ≥ 0.06, single hue (17° spread) |

Fixed slot order — **never cycled, never reordered per chart**:

| 1 orange | 2 blue | 3 aqua | 4 violet | 5 magenta | 6 yellow | 7 green | 8 red |
|---|---|---|---|---|---|---|---|
| `#eb6834` | `#2a78d6` | `#1baf7a` | `#4a3aa7` | `#e87ba4` | `#eda100` | `#008300` | `#e34948` |

Three light-mode slots (aqua, magenta, yellow) sit below 3:1 against white. That is
legal only with relief — which this UI always ships: a legend, direct end labels,
and the same numbers in a table underneath.

Rules the chart runtime enforces:

- 2px lines, ≥8px end markers with a 2px surface ring, area fills at 10%.
- Bars ≤24px thick, 4px rounded data-end, square at the baseline, value at the tip.
- Axis ticks land on clean numbers (1 / 2 / 2.5 / 5 × 10ⁿ) — never `range ÷ n`.
- End labels never stack: they are nudged apart by the minimum needed and a leader
  line connects each back to its line end.
- Heatmap cell text picks ink or white by the **fill's own luminance**, so the
  inverted dark ramp stays legible.
- Legend for ≥2 series; single-series charts get none (the title names it).
- Every chart has a hover layer (crosshair + tooltip on lines, per-mark on bars).
- Exactly one hero figure per view.

---

## Components

| Class | What it is |
|---|---|
| `.shell` / `.rail` / `.stage` | app frame; `.is-collapsed` gives the 58px icon rail |
| `.topbar` + `.readouts` + `.ro` | context bar and live instrument readouts |
| `.tabs` / `.tab` | the browser-style open-page strip |
| `.pagehead` | title, description, page actions |
| `.deck` | the filter deck — the amber-edged control strip above every report |
| `.stats` / `.stat` | KPI tile row (label, value, delta, sparkline) |
| `.hero` | the one big number a view leads with |
| `.panel` (+ `__head`, `__body`, `__foot`) | the universal container, with corner bezel ticks |
| `.dt` (`.dt--banded`) | spreadsheet-grid table: sticky head, group/total rows, `.ind1`/`.ind2` hierarchy, `.n` numerics, `.sep-l` column-group divider |
| `.heat` + `data-heat` | sequential matrix on the same grid; also works on `.dt` cells via `data-v` |
| `.console` | the dark run-log readout |
| `.chat` / `.msg` / `.composer` / `.ctxrail` | Ask AI |
| `.btn` (`--go`, `--ghost`, `--danger`), `.input`, `.select`, `.seg`, `.chip`, `.switch`, `.cbx` | controls |
| `.badge` (`--ok`, `--warn`, `--crit`, `--sig`), `.notice`, `.meter`, `.lvl` | status vocabulary |
| `.rise` + `--d` | the staggered page-load reveal |

---

## How this maps back to the Blazor app

| Mockup | Replaces |
|---|---|
| `.rail` markup + `NAV` array in `instrument.js` | `Components/Layout/NavMenu.razor` + `NavMenu.razor.css` (the `NAV` array is the stand-in for `AppMenus` + `MenuPermissionService`) |
| `.topbar`, `.tabs`, `.stage` | `Components/Layout/MainLayout.razor` + `MainLayout.razor.css` |
| `assets/instrument.css` | `wwwroot/app.css` — and it drops the Bootstrap dependency, so `wwwroot/bootstrap/` and the Bootstrap classes in every page go with it |
| `index.html` | `Pages/BmesDailyReportPage.razor` |
| `ng-rate.html` | `Pages/NgRatePage.razor` (+ `NgRateSetupPanel`, `NgRateModelGroupPicker`) |
| `f-cost.html` | `Pages/BmesFCostPage.razor` |
| `ask-ai.html` | `Pages/DataInferenceAskPage.razor` |
| `admin-users.html` | `Pages/UsersPage.razor` |
| `login.html` | `Pages/LoginPage.razor` |
| `Charts.line/bars/sparks` | replaces the `chart.umd.min.js` (Chart.js) usage for these forms |

Porting order that keeps the app working the whole way: **tokens + shell first**
(app.css, MainLayout, NavMenu), then one page at a time. Because every component
is a plain class, a page can be converted independently — a half-converted app
looks inconsistent but never broken.

### Not addressed here

Everything in the mockup is static. The tab strip does not persist page state, the
sort is a client-side re-sort of the rendered rows, chips and segmented controls
toggle but filter nothing, and there is no data layer. Charts are hand-fed
literals in each page's inline `<script>`.
