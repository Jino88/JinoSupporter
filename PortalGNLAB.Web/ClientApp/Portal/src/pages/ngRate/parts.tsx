import { Fragment, useEffect, useMemo, useRef, useState, type CSSProperties, type ReactNode } from "react";
import type { HierGroup, HierMid, HierSubRow, NgRateSetup, PickerGroup, Ppm } from "../../ngRateApi";
import { navigate } from "../../nav";
import {
  allKeys, formatPpm, hasAnyValue, isHierRowVisible, logClass, plural, ppmClass, ppmOf, toggled, usDate, type ColumnSet,
} from "./logic";

// Shared parts of the NG Rate pages: ports of NgRateViewNav, NgRateSetupPanel, NgRateModelGroupPicker,
// NgRateSimpleGroupPicker and HierSubRows (Components/Shared) plus the period header / cell blocks
// every pivot table repeats.

export const MODEL_GROUP_PAGE = "/app/bmes/make-model-group";
const SETTING_PAGE = "/app/setting";

export const Spinner = () => <span className="ngr-spinner" role="status" aria-hidden="true" />;

const VIEWS = [
  { key: "by-model", label: "By Model", href: "/app/ng-rate", title: "By Model view - show full Summary/Detail/Chart per model, one at a time" },
  { key: "by-section", label: "By Section", href: "/app/ng-rate-all", title: "By Section view - multiple models stacked vertically by Summary/Process/NG/Reason sections" },
  { key: "by-group", label: "By Group", href: "/app/ng-rate-by-group", title: "By Group view" },
] as const;

export function ViewNav({ current }: { current: (typeof VIEWS)[number]["key"] }) {
  return (
    <div className="ngrate-view-nav d-flex align-items-center gap-2 mb-3">
      <h5 className="mb-0 fw-semibold">NG Rate</h5>
      <nav className="ms-3 ngr-btn-group" aria-label="NG Rate view mode">
        {VIEWS.map((v) =>
          v.key === current ? (
            <button key={v.key} type="button" className="ngr-btn ngrate-view-current" disabled aria-current="page" title={`Current view: ${v.title}`}>
              {v.label}
            </button>
          ) : (
            <a key={v.key} href={v.href} className="ngr-btn ngr-btn-outline-secondary ngrate-view-link" title={`Switch to ${v.title}`}
              onClick={(e) => {
                if (e.ctrlKey || e.metaKey || e.shiftKey || e.button !== 0) return;
                e.preventDefault();
                navigate(v.href);
              }}>
              {v.label}
            </a>
          ),
        )}
      </nav>
    </div>
  );
}

export interface SetupPanelProps {
  flow?: boolean;
  start: string;
  end: string;
  onStart: (day: string) => void;
  onEnd: (day: string) => void;
  onToday: () => void;
  busy: boolean;
  setup: NgRateSetup | null;
  savedDbPath: string | null;
  errorMessage: string | null;
  showClearResult?: boolean;
  onClearResult?: () => void;
  logs: string[];
  hasSuccessfulResult: boolean;
  onClearLogs: () => void;
  children: ReactNode;
}

/** Period + BMES credentials + group picker on the left, progress log on the right; both can be hidden. */
export function SetupPanel(p: SetupPanelProps) {
  const [collapsed, setCollapsed] = useState(false);
  const logPane = useRef<HTMLDivElement>(null);
  const showClear = p.showClearResult ?? true;

  // Keep the newest log line in view (window.ngRateLog.scrollToBottom).
  useEffect(() => {
    const el = logPane.current;
    if (el) el.scrollTop = el.scrollHeight;
  }, [p.logs.length, p.busy, collapsed]);

  if (collapsed) {
    return (
      <button type="button" className="ngr-card ngrate-top-toggle ngrate-top-toggle-collapsed" onClick={() => setCollapsed(false)}
        title="Show report setup and progress log">
        <span className="fw-semibold">Report Setup / Progress Log</span>
        <span className="small text-muted">{usDate(p.start)} - {usDate(p.end)}</span>
        <span className="small text-muted ms-auto">{p.busy ? "Processing..." : `${p.logs.length} log(s)`}</span>
        <span className="small fw-semibold text-primary">Show</span>
      </button>
    );
  }

  const hide = (
    <button type="button" className="ngr-link text-secondary" onClick={() => setCollapsed(true)} title="Hide report setup and progress log">Hide</button>
  );
  const credentials = p.setup?.credentialsConfigured ?? false;

  return (
    <div className={p.flow ? "ngrate-flow" : "ngrate-grid"}>
      <div className={p.flow ? "ngrate-flow-data" : "ngrate-cell-tl"}>
        <div className="ngr-card h-100 ngrate-setup-card">
          <div className="card-header ngrate-card-header small fw-semibold d-flex justify-content-between align-items-center">
            <span>Report Setup</span>
            {hide}
          </div>
          <div className="card-body ngrate-card-body">
            <div className="ngrate-setup-grid">
              <div className="ngrate-setup-main d-flex flex-column gap-3">
                <div>
                  <div className="small fw-semibold mb-1">Period</div>
                  <div className="d-flex gap-2 align-items-center flex-wrap">
                    <input type="date" className="ngr-input" value={p.start} onChange={(e) => e.target.value && p.onStart(e.target.value)} />
                    <span className="small text-muted">-</span>
                    <input type="date" className="ngr-input" value={p.end} onChange={(e) => e.target.value && p.onEnd(e.target.value)} />
                    <button type="button" className="ngr-btn ngr-btn-outline-secondary ngr-btn-sm" onClick={p.onToday}>Today</button>
                  </div>
                </div>

                <div className="d-flex align-items-center gap-2">
                  <span className="small text-muted">BMES:</span>
                  {p.setup === null ? (
                    <span className="small text-muted">Loading…</span>
                  ) : credentials ? (
                    <>
                      <span className="ngr-badge bg-success">Credentials OK</span>
                      <span className="small text-muted font-monospace">{p.setup.loginId}</span>
                    </>
                  ) : (
                    <>
                      <span className="ngr-badge bg-warning">Not configured</span>
                      <a href={SETTING_PAGE} className="small text-primary">Setting</a>
                    </>
                  )}
                </div>

                {p.savedDbPath && (
                  <div className="ngr-alert ngr-alert-success small">
                    <div className="fw-semibold mb-1"><span className="ngr-badge bg-success me-1">Done</span> NG data ready</div>
                    <div className="font-monospace text-break" style={{ fontSize: 11 }}>{p.savedDbPath}</div>
                    {showClear && <button type="button" className="ngr-link text-secondary mt-1" onClick={p.onClearResult}>Clear</button>}
                  </div>
                )}
                {p.errorMessage && (
                  <div className="ngr-alert ngr-alert-danger small">
                    <span className="ngr-badge bg-danger me-1">Error</span> {p.errorMessage}
                    {showClear && <button type="button" className="ngr-link text-secondary ms-1" onClick={p.onClearResult}>Clear</button>}
                  </div>
                )}
              </div>
              <div className="ngrate-setup-groups">{p.children}</div>
            </div>
          </div>
        </div>
      </div>

      <div className={p.flow ? "ngrate-flow-progress" : "ngrate-cell-progress-log"}>
        <div className="ngr-card h-100 ngrate-setup-card ngrate-log-card">
          <div className="card-header ngrate-card-header small fw-semibold d-flex justify-content-between align-items-center">
            <span>Progress / Log</span>
            <span className="d-flex align-items-center gap-2">
              {p.logs.length > 0 && <button type="button" className="ngr-link text-secondary" onClick={p.onClearLogs}>Clear</button>}
              {hide}
            </span>
          </div>
          <div className="card-body ngrate-card-body p-0 ngrate-log-card-body">
            <div className="p-3 pb-2">
              <div className="ngr-progress">
                {p.busy ? (
                  <div className="ngr-progress-bar ngr-progress-bar-animated" style={{ width: "100%" }} />
                ) : p.logs.length > 0 ? (
                  <div className={"ngr-progress-bar " + (p.hasSuccessfulResult ? "bg-success" : "bg-danger")} style={{ width: "100%" }} />
                ) : null}
              </div>
            </div>
            <div ref={logPane} className="font-monospace p-2 ngrate-log-pane">
              {p.logs.length === 0 ? (
                <span style={{ opacity: 0.4 }}>No log entries yet.</span>
              ) : (
                <>
                  {p.logs.map((line, i) => <div key={i} className={logClass(line)}>{line}</div>)}
                  <div style={{ height: 12 }} />
                </>
              )}
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

function GetReportLabel({ busy, extra }: { busy: boolean; extra?: string }) {
  return busy ? (
    <><Spinner /> <span>Processing...</span></>
  ) : (
    <>
      <span>Get Report</span>
      {extra && <span className="ms-1" style={{ opacity: 0.75, fontSize: 10 }}>{extra}</span>}
    </>
  );
}

export interface ModelGroupPickerProps {
  setup: NgRateSetup;
  /** "children" = model + every named sub (By Model); "midChildren" = models only (By Section). */
  entries: "children" | "midChildren";
  selectedIds: ReadonlySet<number>;
  selectedEntries: ReadonlySet<string>;
  activeProductGroup: string;
  onActiveProductGroup: (pg: string) => void;
  busy: boolean;
  reportError: string;
  onRefresh: () => void;
  onToggleGroup: (id: number) => void;
  onToggleEntry: (key: string) => void;
  onGetReport: () => void;
}

/** Product-group tabs, then groups with their models / subs as checkboxes (NgRateModelGroupPicker). */
export function ModelGroupPicker(p: ModelGroupPickerProps) {
  const groups = p.setup.groups;
  const selectedCount = p.selectedIds.size + p.selectedEntries.size;
  const visible = groups.filter((g) => g.productGroup === p.activeProductGroup);
  const disabled = p.busy || !p.setup.credentialsConfigured;
  return (
    <>
      <div className="ngrate-setup-groups-head small fw-semibold d-flex justify-content-between align-items-center mb-2">
        <span>Model Groups</span>
        <div className="d-flex align-items-center gap-2">
          {selectedCount > 0 && <span className="ngrate-picker-selection">{selectedCount} selected</span>}
          <span className="text-muted" style={{ fontSize: 11 }}>{plural(groups.length, "group")}</span>
          <button type="button" className="ngr-link text-secondary" onClick={p.onRefresh} title="Refresh">Refresh</button>
        </div>
      </div>
      <div className="ngrate-picker d-flex flex-column min-h-0">
        {groups.length === 0 ? (
          <div className="text-muted small p-3 flex-grow-1">
            <span>No groups found. Define groups on the <a href={MODEL_GROUP_PAGE}>Model Group</a> page.</span>
          </div>
        ) : (
          <>
            <div className="ngrate-picker-tabs d-flex flex-wrap gap-1 px-2 py-2 border-bottom">
              {p.setup.productGroups.map((pg) => {
                const pgCount = groups.filter((g) => g.productGroup === pg).length;
                const pgSel = groups.filter((g) => g.productGroup === pg && p.selectedIds.has(g.id)).length;
                return (
                  <button key={pg} type="button" className={"pg-tab" + (p.activeProductGroup === pg ? " active" : "")}
                    title={`${pg}: ${pgSel} selected of ${pgCount}`} onClick={() => p.onActiveProductGroup(pg)}>
                    {pg}
                    {pgSel > 0 ? <span className="ms-1" style={{ opacity: 0.85 }}>({pgSel})</span> : <span className="ms-1 opacity-50">{pgCount}</span>}
                  </button>
                );
              })}
            </div>
            <div className="ngrate-setup-groups-list">
              <div className="px-2 py-1">
                {visible.map((g) => (
                  <Fragment key={g.id}>
                    <label className="json-file-item" title={g.name}>
                      <input type="checkbox" className="form-check-input" checked={p.selectedIds.has(g.id)} onChange={() => p.onToggleGroup(g.id)} />
                      <span className="ngrate-picker-level">GROUP</span>
                      <span>{g.name}</span>
                      <span className="file-date ngrate-picker-count">{g.lineShiftCount} shifts</span>
                    </label>
                    {g[p.entries].map((e) => (
                      <label key={e.key} className="json-file-item sub-row" title={e.label} style={{ paddingLeft: 8 + e.depth * 14 }}>
                        <input type="checkbox" className="form-check-input" checked={p.selectedEntries.has(e.key)} onChange={() => p.onToggleEntry(e.key)} />
                        <span className="ngrate-picker-level">{e.level}</span>
                        <span>{e.label}</span>
                        <span className="file-date ngrate-picker-count">{e.lsCount} shifts</span>
                      </label>
                    ))}
                  </Fragment>
                ))}
                {visible.length === 0 && <div className="text-muted small py-2 text-center">No groups in {p.activeProductGroup}</div>}
              </div>
            </div>
          </>
        )}

        <div className="border-top pt-2 mt-2">
          <button type="button" className="ngr-btn ngr-btn-success w-100 ngr-btn-sm" onClick={p.onGetReport} disabled={disabled}>
            <GetReportLabel busy={p.busy} extra={selectedCount > 0 ? `(${plural(selectedCount, "group")})` : undefined} />
          </button>
          {p.reportError && <div className="text-danger small mt-1">{p.reportError}</div>}
        </div>
      </div>
    </>
  );
}

/** Product group to show first: the current one while it still has groups, else the first with groups. */
export function activeProductGroupFor(setup: NgRateSetup, current: string): string {
  return setup.groups.some((g) => g.productGroup === current) ? current : setup.productGroups[0] ?? "ETC";
}

export interface SimpleGroupPickerProps {
  groups: PickerGroup[];
  selected: string[];
  busy: boolean;
  credentialsConfigured: boolean;
  onReload: () => void;
  onSelectAll: () => void;
  onClear: () => void;
  onToggle: (name: string, checked: boolean) => void;
  onGetReport: () => void;
}

/** Plain group checkboxes for By Group (NgRateSimpleGroupPicker). */
export function SimpleGroupPicker(p: SimpleGroupPickerProps) {
  const selected = new Set(p.selected);
  return (
    <>
      <div className="ngrate-setup-groups-head small fw-semibold d-flex align-items-center gap-2 mb-2">
        <span>Sub Groups</span>
        <span className="ngrate-picker-count">{p.groups.length} groups</span>
        <button type="button" className="ngr-btn ngr-btn-outline-secondary ngr-btn-sm ms-auto" onClick={p.onReload}>Refresh</button>
        <a className="ngr-btn ngr-btn-outline-secondary ngr-btn-sm" href={MODEL_GROUP_PAGE} target="_blank" rel="noreferrer" title="Edit groups">Edit</a>
      </div>
      <div className="ngrate-picker ngrate-setup-groups-body d-flex flex-column gap-2">
        {p.groups.length === 0 ? (
          <div className="text-muted small">
            No saved sub groups. Please create them first at <a href={MODEL_GROUP_PAGE} target="_blank" rel="noreferrer">Model Group</a>.
          </div>
        ) : (
          <>
            <div className="d-flex gap-2 small">
              <button type="button" className="ngr-link" onClick={p.onSelectAll}>All</button>
              <span className="text-muted">|</span>
              <button type="button" className="ngr-link" onClick={p.onClear}>None</button>
              <span className="ngrate-picker-selection ms-auto">{selected.size} selected</span>
            </div>
            <div className="ngrate-setup-groups-list">
              {p.groups.map((g) => (
                <label key={g.id} className="grp-chk">
                  <input type="checkbox" checked={selected.has(g.name)} onChange={(e) => p.onToggle(g.name, e.target.checked)} />
                  <span className="ngrate-picker-level">GROUP</span>
                  <span className="grp-name">{g.name}</span>
                  <span className="grp-count ngrate-picker-count">{g.midCount} models · {g.lineShiftCount} shifts</span>
                </label>
              ))}
            </div>
          </>
        )}
        <button type="button" className="ngr-btn ngr-btn-success mt-1" onClick={p.onGetReport}
          disabled={p.busy || selected.size === 0 || !p.credentialsConfigured}>
          <GetReportLabel busy={p.busy} />
        </button>
      </div>
    </>
  );
}

// ── Period blocks ─────────────────────────────────────────────────────────────────────────────

/** Single header row: Date | sep | Week | sep | Month (separators always drawn). */
export function FlatPeriodHeads({ cols }: { cols: ColumnSet }) {
  return (
    <>
      {cols.date.map((c) => <th key={c.key}>{c.header}</th>)}
      <th className="sep-th" />
      {cols.week.map((c) => <th key={c.key}>{c.header}</th>)}
      <th className="sep-th" />
      {cols.month.map((c) => <th key={c.key}>{c.header}</th>)}
    </>
  );
}

/** Value cells matching FlatPeriodHeads; `plain` = group sub-rows (no bold for values). */
export function FlatPeriodCells({ cols, ppm, plain }: { cols: ColumnSet; ppm: Ppm; plain?: boolean }) {
  const cell = (key: string) => {
    const v = ppmOf(ppm, key);
    return <td key={key} className={v <= 0 ? "ppm-zero" : plain ? "" : "ppm-val"}>{formatPpm(v)}</td>;
  };
  return (
    <>
      {cols.date.map((c) => cell(c.key))}
      <td className="sep-td" />
      {cols.week.map((c) => cell(c.key))}
      <td className="sep-td" />
      {cols.month.map((c) => cell(c.key))}
    </>
  );
}

/** Two header rows with Date / Week / Month block captions (separators only after non-empty blocks). */
export function BlockPeriodHeads({ cols, label, labelStyle }: { cols: ColumnSet; label: string; labelStyle?: CSSProperties }) {
  return (
    <thead>
      <tr>
        <th className="label-th" style={labelStyle} rowSpan={2}>{label}</th>
        {cols.date.length > 0 && <><th className="blk-date" colSpan={cols.date.length}>Date</th><th className="sep-th" rowSpan={2} /></>}
        {cols.week.length > 0 && <><th className="blk-week" colSpan={cols.week.length}>Week</th><th className="sep-th" rowSpan={2} /></>}
        {cols.month.length > 0 && <th className="blk-month" colSpan={cols.month.length}>Month</th>}
      </tr>
      <tr>
        {cols.date.map((c) => <th key={c.key}>{c.header}</th>)}
        {cols.week.map((c) => <th key={c.key}>{c.header}</th>)}
        {cols.month.map((c) => <th key={c.key}>{c.header}</th>)}
      </tr>
    </thead>
  );
}

export function BlockPeriodCells({ cols, value }: { cols: ColumnSet; value: (key: string) => number }) {
  const cell = (key: string) => {
    const v = value(key);
    return <td key={key} className={ppmClass(v)}>{formatPpm(v)}</td>;
  };
  return (
    <>
      {cols.date.map((c) => cell(c.key))}
      {cols.date.length > 0 && <td className="sep-td" />}
      {cols.week.map((c) => cell(c.key))}
      {cols.week.length > 0 && <td className="sep-td" />}
      {cols.month.map((c) => cell(c.key))}
    </>
  );
}

// ── Hierarchy ─────────────────────────────────────────────────────────────────────────────────

export interface HierRowClasses {
  sub1: string;
  sub2: string;
  lineShift: string;
  expandable: string;
}

export const HIER_CLASSES: HierRowClasses = { sub1: "hier-sub1-row", sub2: "hier-sub2-row", lineShift: "hier-ls-row", expandable: "hier-sub-expandable" };

/** Sub / Shift rows under one model row (HierSubRows.razor). */
export function HierSubRows({ mid, cols, expanded, onToggle, classes = HIER_CLASSES, basePad = 48, indent = 16 }: {
  mid: HierMid;
  cols: ColumnSet;
  expanded: ReadonlySet<string>;
  onToggle: (rowId: string) => void;
  classes?: HierRowClasses;
  basePad?: number;
  indent?: number;
}) {
  const byId = useMemo(() => new Map<string, HierSubRow>(mid.rows.map((r) => [r.rowId, r])), [mid]);
  return (
    <>
      {mid.rows.filter((r) => isHierRowVisible(r, byId, expanded)).map((row) => {
        const rowClass = row.isLineShift ? classes.lineShift : row.depth === 1 ? classes.sub1 : classes.sub2;
        const open = expanded.has(row.rowId);
        return (
          <tr key={row.rowId} className={rowClass + (row.hasChildren ? " " + classes.expandable : "")}
            onClick={() => row.hasChildren && onToggle(row.rowId)}>
            <td className="label-td font-monospace" style={{ fontSize: 11, paddingLeft: basePad + (row.depth - 1) * indent }}>
              <span className="tog-icon" style={row.hasChildren ? undefined : { visibility: "hidden" }}>{row.hasChildren && open ? "▼" : "▶"}</span>
              <span className="ngrate-hier-level">{row.isLineShift ? "SHIFT" : `SUB ${row.depth}`}</span>
              <span className="ms-1">{row.display}</span>
            </td>
            <BlockPeriodCells cols={cols} value={(k) => ppmOf(row.ppm, k)} />
          </tr>
        );
      })}
    </>
  );
}

/** Group → Model → Sub → Shift card shared by By Model and By Section (collapsible rows). */
export function HierCard({ hier, cols, title, label }: { hier: HierGroup[]; cols: ColumnSet; title: string; label: string }) {
  const [groups, setGroups] = useState<Set<string>>(new Set());
  const [mids, setMids] = useState<Set<string>>(new Set());
  const [subs, setSubs] = useState<Set<string>>(new Set());
  const keys = allKeys(cols);
  return (
    <div className="ngr-card mb-3 ngrate-report-card">
      <div className="card-header ngrate-card-header small fw-semibold">{title}</div>
      <div className="card-body ngrate-card-body p-0">
        <div className="pivot-wrap">
          <table className="pivot-table text-fit-table">
            <BlockPeriodHeads cols={cols} label={label} labelStyle={{ minWidth: 280 }} />
            <tbody>
              {hier.map((g) => {
                const open = groups.has(g.name);
                const visibleMids = g.mids.filter((m) => hasAnyValue(m.ppm, keys));
                return (
                  <Fragment key={g.name}>
                    <tr className="hier-group-row" onClick={() => setGroups(toggled(groups, g.name))}>
                      <td className="label-td">
                        <span className="tog-icon">{open ? "▼" : "▶"}</span>
                        <span className="ngrate-hier-level">GROUP</span>
                        <strong>{g.name}</strong>
                        <span className="text-muted ms-1" style={{ fontSize: 10 }}>{visibleMids.length} models</span>
                      </td>
                      <BlockPeriodCells cols={cols} value={(k) => ppmOf(g.ppm, k)} />
                    </tr>
                    {open && visibleMids.map((m) => {
                      const midKey = `${g.name}::${m.material}`;
                      const midOpen = mids.has(midKey);
                      return (
                        <Fragment key={midKey}>
                          <tr className="hier-mid-row" onClick={() => setMids(toggled(mids, midKey))}>
                            <td className="label-td" style={{ paddingLeft: 32, fontSize: 11.5 }}>
                              <span className="tog-icon">{midOpen ? "▼" : "▶"}</span>
                              <span className="ngrate-hier-level">MODEL</span>
                              <span className="ms-1">{m.material}</span>
                            </td>
                            <BlockPeriodCells cols={cols} value={(k) => ppmOf(m.ppm, k)} />
                          </tr>
                          {midOpen && <HierSubRows mid={m} cols={cols} expanded={subs} onToggle={(id) => setSubs(toggled(subs, id))} />}
                        </Fragment>
                      );
                    })}
                  </Fragment>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}
