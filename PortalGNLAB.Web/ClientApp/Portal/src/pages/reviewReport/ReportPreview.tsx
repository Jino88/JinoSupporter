import { useEffect, useState, type KeyboardEvent } from "react";
import { toolSend } from "../../toolsApi";
import { addTag, dash, editLine, groupPictures, removeTag } from "./logic";
import {
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
  onCommit,
}: {
  value: string;
  editable: boolean;
  multiline?: boolean;
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

function ResultGrid({ table }: { table: PreviewTable }) {
  return (
    <div className="scroll-x">
      <h4>{table.title}</h4>
      <table className="grid report-result-grid">
        <thead>
          <tr>
            <th>No</th>
            {table.showDate && <th>Date</th>}
            {table.showModel && <th>Model</th>}
            <th>Type</th>
            {table.columns.map((column, i) => (
              <th key={i}>
                {column.header}
                {column.hasBelow && (
                  <div className="report-rate">{column.belowHeader}</div>
                )}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {table.rows.map((row, i) => (
            <tr key={i}>
              <td>{i + 1}</td>
              {table.showDate && table.dateSpans[i] > 0 && (
                <td rowSpan={table.dateSpans[i]}>{row.date}</td>
              )}
              {table.showModel && table.modelSpans[i] > 0 && (
                <td rowSpan={table.modelSpans[i]}>{row.model}</td>
              )}
              <td className={row.test ? "report-test" : ""}>{row.label}</td>
              {row.cells.map((cell, j) => (
                <td key={j} className={row.test ? "report-test" : ""}>
                  <div className={cell.rate ? "report-rate" : ""}>
                    {cell.main}
                  </div>
                  {cell.hasBelow && (
                    <div className="report-rate report-below">{cell.below}</div>
                  )}
                </td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
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
    prefix?: string,
  ) {
    return d[key].map((line, i) =>
      line.trim() ? (
        <p key={`${i}:${line}`}>
          {editable ? (
            <>
              {prefix ?? (/^[-•]/.test(line.trim()) ? "" : "- ")}
              <Editable
                value={line}
                editable
                onCommit={(value) =>
                  set({ [key]: editLine(d[key], i, value) } as Partial<ReportDoc>)
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
    );
  }

  function text(
    key: "purpose" | "background" | "conclusion",
    multiline = true,
  ) {
    return d[key].trim() ? (
      <p key={key}>
        {editable ? (
          <>
            {/^[-•]/.test(d[key].trim()) ? "" : "- "}
            <Editable
              value={d[key]}
              editable
              multiline={multiline}
              onCommit={(value) => set({ [key]: value } as Partial<ReportDoc>)}
            />
          </>
        ) : (
          dash(d[key])
        )}
      </p>
    ) : null;
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
        {groupPictures(placements).map((group, i) => (
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
                    />
                    {group.length < 3 && (
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
            {group.length >= 3 && (
              <div className="report-picture-caption">
                {group[0].group || group[0].caption}
              </div>
            )}
          </div>
        ))}
      </div>
    );
  }

  return (
    <article className="report-document">
      {editable && (
        <p className="muted small">
          Click any text in the report to change it; clear a line to delete
          it. Save to keep your changes.
        </p>
      )}
      <h2>
        <Editable
          value={d.title}
          editable={editable}
          onCommit={(title) => set({ title })}
        />
      </h2>
      <table className="grid">
        <tbody>
          {(d.productModel || editable) && (
            <tr>
              <th>Product model</th>
              <td>
                <Editable
                  value={d.productModel}
                  editable={editable}
                  onCommit={(productModel) => set({ productModel })}
                />
              </td>
            </tr>
          )}
          {(d.reportDate || editable) && (
            <tr>
              <th>Report date</th>
              <td>
                <Editable
                  value={d.reportDate}
                  editable={editable}
                  onCommit={(reportDate) => set({ reportDate })}
                />
              </td>
            </tr>
          )}
          <tr>
            <th>Author</th>
            <td>{report.createdBy}</td>
          </tr>
          {(d.tags.length > 0 || editable) && (
            <tr>
              <th>Tags</th>
              <td>
                <TagEditor doc={d} editable={editable} onChange={onChange} />
              </td>
            </tr>
          )}
        </tbody>
      </table>
      <h3>I. Purpose.</h3>
      {text("purpose")}
      {text("background")}
      {pictures("purpose")}
      <h3>II. Content.</h3>
      {lines("content")}
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
      <h3>III. Result</h3>
      {tables.map((table, i) => (
        <ResultGrid
          table={{ ...table, title: `${i + 1}. ${table.title}` }}
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
                      { length: Math.max(row.length, table.columns.length) },
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
                                        { length: Math.max(r.length, k + 1) },
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
          The numbers in the calculated tables come from Metrics; change them
          under Edit report details.
        </p>
      )}
      {pictures("results")}
      {lines("findings")}
      <h3>IV. Conclusion</h3>
      {text("conclusion")}
      {lines("limitations")}
      {lines("nextActions", "→ ")}
      {pictures("conclusion")}
      {d.images.some(
        (p) =>
          p.section === "appendix" &&
          report.images.some((image) => image.no === p.no),
      ) && (
        <>
          <h3>Appendix</h3>
          {pictures("appendix")}
        </>
      )}
    </article>
  );
}
