import { useEffect, useState, type DragEvent, type KeyboardEvent } from "react";
import { toolSend } from "../../toolsApi";
import {
  addTag,
  dash,
  editLine,
  formatDay,
  groupPictures,
  removeTag,
  renameMetric,
  renameTable,
  replaceText,
  setArms,
  setMetricValue,
} from "./logic";
import {
  pictureDragType,
  sections,
  type Placement,
  type PreviewTable,
  type Report,
  type ReportDoc,
} from "./types";

interface Props {
  report: Report;
  /** False while a request runs: the text is shown but cannot be edited. */
  editable: boolean;
  onChange: (doc: ReportDoc) => void;
  onMove: (no: number, section: string) => void;
  onRemovePicture: (no: number) => void;
  onError: (message: string) => void;
}

/**
 * Text a person can click and change in place. The edit is committed on blur (Enter for
 * single-line text, Escape cancels). The key includes the value so React replaces the
 * element instead of reconciling text the browser has already edited.
 */
function Editable({
  value,
  editable,
  multiline = false,
  placeholder = "(empty)",
  onCommit,
}: {
  value: string;
  editable: boolean;
  multiline?: boolean;
  /** Shown while the text is empty, so an empty field can still be clicked. */
  placeholder?: string;
  onCommit: (value: string) => void;
}) {
  if (!editable) return <>{value}</>;
  function keyDown(e: KeyboardEvent<HTMLSpanElement>) {
    if (e.key === "Escape") {
      e.currentTarget.innerText = value;
      e.currentTarget.blur();
    } else if (e.key === "Enter" && !multiline) {
      e.preventDefault();
      e.currentTarget.blur();
    }
  }
  return (
    <span
      key={value}
      className="report-editable"
      contentEditable="plaintext-only"
      suppressContentEditableWarning
      title="Click to edit"
      data-placeholder={placeholder}
      onKeyDown={keyDown}
      onBlur={(e) => {
        const next = e.currentTarget.innerText.replace(/ /g, " ");
        if (next.trim() !== value.trim()) onCommit(next.trim());
      }}
    >
      {value}
    </span>
  );
}

/**
 * A result table worked out by the server. Every cell writes back to what it was made from:
 * the title to the conditions' table name, Date / Model / Type to the conditions, headers and
 * numbers to the metrics. The NG ppm and an added Total NG are worked out from the counts, so
 * they change when a count or the input changes.
 */
function ResultGrid({
  table,
  number,
  doc,
  editable,
  onChange,
}: {
  table: PreviewTable;
  number: number;
  doc: ReportDoc;
  editable: boolean;
  onChange: (doc: ReportDoc) => void;
}) {
  const rows = table.rows;
  const spanned = (spans: number[], i: number) =>
    rows.slice(i, i + spans[i]).map((row) => row.armId);
  const cell = (
    value: string,
    commit: ((value: string) => ReportDoc) | null,
  ) =>
    commit && editable ? (
      <Editable
        value={value}
        editable
        placeholder="-"
        onCommit={(next) => onChange(commit(next))}
      />
    ) : (
      value
    );
  const cellClass = (row: PreviewTable["rows"][number], j: number) =>
    [
      row.test ? "report-test" : "",
      table.columns[j]?.rateColumn ? "report-rate" : "",
    ]
      .filter(Boolean)
      .join(" ");
  return (
    <div className="scroll-x report-table-block">
      <h4>
        {number}.{" "}
        {cell(table.title, (title) => renameTable(doc, table.tableKey, title))}
      </h4>
      <table className="report-result-grid">
        <thead>
          <tr>
            <th>No</th>
            {table.showDate && <th>Date</th>}
            {table.showModel && <th>Model</th>}
            <th>Type</th>
            {table.columns.map((column, i) => (
              <th key={i}>
                {cell(
                  column.header,
                  column.mainId && !column.computed
                    ? (header) => renameMetric(doc, column.mainId, header)
                    : null,
                )}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {rows.flatMap((row, i) => [
            <tr key={`${i}a`}>
              <td rowSpan={2}>{i + 1}</td>
              {table.showDate && table.dateSpans[i] > 0 && (
                <td rowSpan={table.dateSpans[i] * 2}>
                  {cell(row.dateLabel || row.date, (date) =>
                    setArms(doc, spanned(table.dateSpans, i), { date }),
                  )}
                </td>
              )}
              {table.showModel && table.modelSpans[i] > 0 && (
                <td rowSpan={table.modelSpans[i] * 2}>
                  {cell(row.model, (model) =>
                    setArms(doc, spanned(table.modelSpans, i), { model }),
                  )}
                </td>
              )}
              <td rowSpan={2} className={row.test ? "report-test" : ""}>
                {cell(row.label, (label) =>
                  setArms(doc, [row.armId], { label }),
                )}
              </td>
              {row.cells.map((value, j) => {
                const column = table.columns[j];
                const split = value.hasBelow && !column?.rateColumn;
                return (
                  <td
                    key={j}
                    rowSpan={split ? 1 : 2}
                    className={cellClass(row, j)}
                  >
                    {cell(
                      value.main,
                      column?.mainId && !column.computed
                        ? (v) =>
                            setMetricValue(doc, column.mainId, row.armId, v)
                        : null,
                    )}
                  </td>
                );
              })}
            </tr>,
            <tr key={`${i}b`} className="report-below-row">
              {row.cells.map((value, j) => {
                const column = table.columns[j];
                if (!value.hasBelow || column?.rateColumn) return null;
                return (
                  <td key={j} className={`report-rate ${cellClass(row, j)}`}>
                    {cell(
                      value.below,
                      column?.rateId
                        ? (v) =>
                            setMetricValue(doc, column.rateId, row.armId, v)
                        : null,
                    )}
                  </td>
                );
              })}
            </tr>,
          ])}
        </tbody>
      </table>
    </div>
  );
}

/** Replaces a word everywhere in the report at once (e.g. a model name the AI misread). */
function ReplaceBar({
  doc,
  onChange,
}: {
  doc: ReportDoc;
  onChange: (doc: ReportDoc) => void;
}) {
  const [find, setFind] = useState("");
  const [replacement, setReplacement] = useState("");
  const [result, setResult] = useState("");
  function replace() {
    const next = replaceText(doc, find, replacement);
    if (next.count) onChange(next.doc);
    setResult(
      next.count
        ? `Replaced ${next.count} place${next.count === 1 ? "" : "s"}.`
        : `"${find}" is not in the report.`,
    );
  }
  return (
    <div className="form-row report-replace">
      <input
        className="inp sm"
        placeholder="Find"
        aria-label="Find in report"
        value={find}
        onChange={(e) => setFind(e.target.value)}
      />
      <input
        className="inp sm"
        placeholder="Replace with"
        aria-label="Replace with"
        value={replacement}
        onChange={(e) => setReplacement(e.target.value)}
        onKeyDown={(e) => {
          if (e.key === "Enter" && find) replace();
        }}
      />
      <button className="btn sm" disabled={!find} onClick={replace}>
        Replace all
      </button>
      {result && <span className="muted small">{result}</span>}
    </div>
  );
}

function TagEditor({
  doc,
  editable,
  onChange,
}: {
  doc: ReportDoc;
  editable: boolean;
  onChange: (doc: ReportDoc) => void;
}) {
  const [draft, setDraft] = useState("");
  function add() {
    onChange(addTag(doc, draft));
    setDraft("");
  }
  return (
    <>
      {doc.tags.map((tag, i) => (
        <span className="tag ok report-tag" key={`${i}:${tag}`}>
          {tag}
          {editable && (
            <button
              className="report-tag-remove"
              title={`Remove tag ${tag}`}
              aria-label={`Remove tag ${tag}`}
              onClick={() => onChange(removeTag(doc, i))}
            >
              ×
            </button>
          )}
        </span>
      ))}
      {editable && (
        <span className="report-tag-add">
          <input
            className="inp sm"
            placeholder="Add tag"
            value={draft}
            onChange={(e) => setDraft(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === "Enter") {
                e.preventDefault();
                add();
              }
            }}
          />
          <button className="btn sm" disabled={!draft.trim()} onClick={add}>
            Add
          </button>
        </span>
      )}
    </>
  );
}

export default function ReportPreview({
  report,
  editable,
  onChange,
  onMove,
  onRemovePicture,
  onError,
}: Props) {
  const [tables, setTables] = useState<PreviewTable[]>([]);
  const [dropSection, setDropSection] = useState("");
  const doc = report.doc;
  useEffect(() => {
    let active = true;
    const timer = window.setTimeout(() => {
      toolSend<PreviewTable[]>("review-report/preview", { report })
        .then((value) => {
          if (active) setTables(value);
        })
        .catch((error: Error) => {
          if (active) onError(error.message);
        });
    }, 150);
    return () => {
      active = false;
      window.clearTimeout(timer);
    };
  }, [report]);

  if (!doc)
    return <p className="muted">The report appears here after Write report.</p>;
  const d = doc;

  function set(change: Partial<ReportDoc>) {
    onChange({ ...d, ...change });
  }

  /** "- line" paragraphs of a string list; clearing a line removes it. */
  function lines(
    key: "content" | "findings" | "limitations" | "nextActions",
    addLabel: string,
    prefix?: string,
    bold = false,
  ) {
    const list = d[key];
    return (
      <>
        {list.map((line, i) =>
          line.trim() ? (
            <p
              key={`${i}:${line}`}
              className={bold ? "report-bold" : undefined}
            >
              {editable ? (
                <>
                  {prefix ?? (/^[-•]/.test(line.trim()) ? "" : "- ")}
                  <Editable
                    value={line}
                    editable
                    onCommit={(value) =>
                      set({
                        [key]: editLine(d[key], i, value),
                      } as Partial<ReportDoc>)
                    }
                  />
                </>
              ) : prefix ? (
                `${prefix}${line}`
              ) : (
                dash(line)
              )}
            </p>
          ) : null,
        )}
        {editable && (
          <p className="report-add">
            {prefix ?? "+ "}
            <Editable
              key={`add:${list.length}`}
              value=""
              editable
              placeholder={addLabel}
              onCommit={(value) =>
                value && set({ [key]: [...list, value] } as Partial<ReportDoc>)
              }
            />
          </p>
        )}
      </>
    );
  }

  function text(
    key: "purpose" | "background" | "conclusion",
    addLabel: string,
    multiline = true,
    arrow = false,
  ) {
    const lead = arrow ? "=> " : "- ";
    const led = arrow ? /^(=>|→)/ : /^[-•]/;
    const className = arrow ? "report-bold" : undefined;
    if (!d[key].trim())
      return editable ? (
        <p key={key} className="report-add">
          {arrow && lead}
          <Editable
            value=""
            editable
            multiline={multiline}
            placeholder={addLabel}
            onCommit={(value) => set({ [key]: value } as Partial<ReportDoc>)}
          />
        </p>
      ) : null;
    return (
      <p key={key} className={className}>
        {editable ? (
          <>
            {led.test(d[key].trim()) ? "" : lead}
            <Editable
              value={d[key]}
              editable
              multiline={multiline}
              onCommit={(value) => set({ [key]: value } as Partial<ReportDoc>)}
            />
          </>
        ) : arrow ? (
          led.test(d[key].trim()) ? (
            d[key].trim()
          ) : (
            `${lead}${d[key].trim()}`
          )
        ) : (
          dash(d[key])
        )}
      </p>
    );
  }

  /** A report section that takes a picture dragged from the picture list or from another section. */
  function dropZone(section: string) {
    const accepts = (e: DragEvent) =>
      editable && e.dataTransfer.types.includes(pictureDragType);
    return {
      className: `report-section ${dropSection === section ? "drop-over" : ""}`,
      onDragOver: (e: DragEvent<HTMLElement>) => {
        if (!accepts(e)) return;
        e.preventDefault();
        e.dataTransfer.dropEffect = "move";
        setDropSection(section);
      },
      onDragLeave: (e: DragEvent<HTMLElement>) => {
        if (!e.currentTarget.contains(e.relatedTarget as Node | null))
          setDropSection("");
      },
      onDrop: (e: DragEvent<HTMLElement>) => {
        setDropSection("");
        const no = Number(e.dataTransfer.getData(pictureDragType));
        if (!accepts(e) || !no) return;
        e.preventDefault();
        onMove(no, section);
      },
    };
  }

  function setPlacement(no: number, change: Partial<Placement>) {
    set({
      images: d.images.map((p) => (p.no === no ? { ...p, ...change } : p)),
    });
  }

  function pictures(section: string) {
    const placements = d.images.filter(
      (p) =>
        p.section === section &&
        report.images.some((image) => image.no === p.no),
    );
    return (
      <div
        className={`report-pictures ${placements.length === 1 ? "single" : ""}`}
      >
        {groupPictures(placements).map((group, i) => {
          // One bar under a whole group, as in the Excel form, when it has three or more
          // pictures or none of them has a caption of its own.
          const oneBar =
            group.length >= 3 ||
            (group.length > 1 && group.every((p) => !p.caption.trim()));
          return (
            <div key={i} className="report-picture-group">
              <div className="report-picture-row">
                {group.map((placement: Placement) => {
                  const image = report.images.find(
                    (image) => image.no === placement.no,
                  )!;
                  return (
                    <figure key={placement.no}>
                      <img
                        src={`data:${image.mediaType};base64,${image.data}`}
                        alt={`IMG ${image.no}: ${placement.caption}`}
                        draggable={editable}
                        title={editable ? "Drag to another section" : undefined}
                        onDragStart={(e) => {
                          e.dataTransfer.setData(
                            pictureDragType,
                            String(image.no),
                          );
                          e.dataTransfer.effectAllowed = "move";
                        }}
                      />
                      {!oneBar && (
                        <figcaption>
                          <Editable
                            value={placement.caption || placement.group}
                            editable={editable}
                            onCommit={(caption) =>
                              setPlacement(placement.no, { caption })
                            }
                          />
                        </figcaption>
                      )}
                      <label className="form-row">
                        IMG {image.no}
                        <select
                          className="inp sm"
                          title="Move this picture to another section"
                          value={placement.section}
                          onChange={(e) => onMove(image.no, e.target.value)}
                        >
                          {sections.map((name) => (
                            <option key={name}>{name}</option>
                          ))}
                        </select>
                        <button
                          className="btn sm danger"
                          title="Delete this picture from the report"
                          onClick={() => onRemovePicture(image.no)}
                        >
                          Delete
                        </button>
                      </label>
                    </figure>
                  );
                })}
              </div>
              {oneBar && (
                <div className="report-picture-caption">
                  <Editable
                    value={group[0].group || group[0].caption}
                    editable={editable}
                    onCommit={(name) =>
                      set({
                        images: d.images.map((p) =>
                          group.some((g) => g.no === p.no)
                            ? { ...p, group: name }
                            : p,
                        ),
                      })
                    }
                  />
                </div>
              )}
            </div>
          );
        })}
      </div>
    );
  }

  return (
    <div className="report-document">
      {editable && (
        <>
          <p className="muted small">
            Click any text or table cell in the report to change it; clear a
            line to delete it. Drag a picture from the left into a section to
            place it. Save to keep your changes.
          </p>
          <ReplaceBar doc={d} onChange={onChange} />
        </>
      )}
      {(d.productModel || d.tags.length > 0 || editable) && (
        <div className="report-meta small">
          {(d.productModel || editable) && (
            <span>
              Product model:{" "}
              <Editable
                value={d.productModel}
                editable={editable}
                onCommit={(productModel) => set({ productModel })}
              />
            </span>
          )}
          <span>
            Tags: <TagEditor doc={d} editable={editable} onChange={onChange} />
          </span>
        </div>
      )}
      <article className="report-sheet">
        <div className="report-head">
          <div className="report-head-label">TITLE</div>
          <h2 className="report-head-title">
            <Editable
              value={d.title}
              editable={editable}
              onCommit={(title) => set({ title })}
            />
          </h2>
          <div className="report-head-info">
            <div className="report-head-label">Dept</div>
            <div>
              <Editable
                value={d.dept ?? ""}
                editable={editable}
                placeholder="-"
                onCommit={(dept) => set({ dept })}
              />
            </div>
            <div className="report-head-label">Date</div>
            <div>
              {editable ? (
                <Editable
                  value={formatDay(d.reportDate)}
                  editable
                  placeholder="-"
                  onCommit={(reportDate) => set({ reportDate })}
                />
              ) : (
                formatDay(d.reportDate)
              )}
            </div>
            <div className="report-head-label">Marker</div>
            <div>
              <Editable
                value={d.maker || report.createdBy}
                editable={editable}
                placeholder="-"
                onCommit={(maker) => set({ maker })}
              />
            </div>
          </div>
        </div>
        <section {...dropZone("purpose")}>
          <h3>I. Purpose.</h3>
          {text("purpose", "Click to add the purpose")}
          {text("background", "Click to add the background")}
          {pictures("purpose")}
        </section>
        <section {...dropZone("conditions")}>
          <h3>II. Content.</h3>
          {lines("content", "Add a content line")}
          {d.changedFactors.map((factor, i) =>
            factor.name.trim() ? (
              <p key={i}>
                -{" "}
                <Editable
                  value={factor.name}
                  editable={editable}
                  onCommit={(name) =>
                    set({
                      changedFactors: name
                        ? d.changedFactors.map((f, j) =>
                            j === i ? { ...f, name } : f,
                          )
                        : d.changedFactors.filter((_, j) => j !== i),
                    })
                  }
                />
                :{" "}
                <Editable
                  value={factor.baseline}
                  editable={editable}
                  onCommit={(baseline) =>
                    set({
                      changedFactors: d.changedFactors.map((f, j) =>
                        j === i ? { ...f, baseline } : f,
                      ),
                    })
                  }
                />{" "}
                →{" "}
                <Editable
                  value={factor.changed}
                  editable={editable}
                  onCommit={(changed) =>
                    set({
                      changedFactors: d.changedFactors.map((f, j) =>
                        j === i ? { ...f, changed } : f,
                      ),
                    })
                  }
                />
              </p>
            ) : null,
          )}
          {d.conditions.map((condition, i) =>
            condition.details.trim() ? (
              <p key={i}>
                -{" "}
                <Editable
                  value={condition.label}
                  editable={editable}
                  onCommit={(label) =>
                    set({
                      conditions: d.conditions.map((c, j) =>
                        j === i ? { ...c, label } : c,
                      ),
                    })
                  }
                />
                :{" "}
                <Editable
                  value={condition.details}
                  editable={editable}
                  onCommit={(details) =>
                    set({
                      conditions: d.conditions.map((c, j) =>
                        j === i ? { ...c, details } : c,
                      ),
                    })
                  }
                />
              </p>
            ) : null,
          )}
          {pictures("conditions")}
        </section>
        <section {...dropZone("results")}>
          <h3>III. Result</h3>
          {tables.map((table, i) => (
            <ResultGrid
              table={table}
              number={i + 1}
              doc={d}
              editable={editable}
              onChange={onChange}
              key={i}
            />
          ))}
          {d.resultTables.map((table, i) => {
            const setTable = (change: Partial<typeof table>) =>
              set({
                resultTables: d.resultTables.map((t, j) =>
                  j === i ? { ...t, ...change } : t,
                ),
              });
            return (
              <div className="scroll-x" key={i}>
                <h4>
                  {tables.length + i + 1}.{" "}
                  <Editable
                    value={table.caption || "Result"}
                    editable={editable}
                    onCommit={(caption) => setTable({ caption })}
                  />
                </h4>
                <table className="grid report-result-grid">
                  <thead>
                    <tr>
                      {table.columns.map((column, j) => (
                        <th key={j}>
                          <Editable
                            value={column}
                            editable={editable}
                            onCommit={(value) =>
                              setTable({
                                columns: table.columns.map((c, k) =>
                                  k === j ? value : c,
                                ),
                              })
                            }
                          />
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {table.rows.map((row, j) => (
                      <tr key={j}>
                        {Array.from(
                          {
                            length: Math.max(row.length, table.columns.length),
                          },
                          (_, k) => (
                            <td key={k}>
                              <Editable
                                value={row[k] ?? ""}
                                editable={editable}
                                onCommit={(value) =>
                                  setTable({
                                    rows: table.rows.map((r, rj) =>
                                      rj === j
                                        ? Array.from(
                                            {
                                              length: Math.max(r.length, k + 1),
                                            },
                                            (_, rk) =>
                                              rk === k ? value : (r[rk] ?? ""),
                                          )
                                        : r,
                                    ),
                                  })
                                }
                              />
                            </td>
                          ),
                        )}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            );
          })}
          {editable && tables.length > 0 && (
            <p className="muted small">
              Rates (%) and OK / Total NG / Total NG rate are worked out from
              the counts and the input.
            </p>
          )}
          {pictures("results")}
          {lines("findings", "Add a finding")}
        </section>
        <section {...dropZone("conclusion")}>
          <h3>IV. Decision</h3>
          {text("conclusion", "Click to add the conclusion", true, true)}
          {lines("limitations", "Add a limitation")}
          {lines("nextActions", "Add a next action", "=> ", true)}
          {pictures("conclusion")}
        </section>
        {d.images.some(
          (p) =>
            p.section === "appendix" &&
            report.images.some((image) => image.no === p.no),
        ) && (
          <section {...dropZone("appendix")}>
            <h3>Appendix</h3>
            {pictures("appendix")}
          </section>
        )}
      </article>
    </div>
  );
}
