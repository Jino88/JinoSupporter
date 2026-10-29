import { useEffect, useState } from "react";
import { toolSend } from "../../toolsApi";
import { dash, groupPictures } from "./logic";
import {
  sections,
  type Placement,
  type PreviewTable,
  type Report,
} from "./types";

interface Props {
  report: Report;
  onMove: (no: number, section: string) => void;
  onError: (message: string) => void;
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

export default function ReportPreview({ report, onMove, onError }: Props) {
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

  function pictures(section: string) {
    const placements = doc!.images.filter(
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
                        {placement.caption || placement.group}
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
      <h2>{doc.title}</h2>
      <table className="grid">
        <tbody>
          {doc.productModel && (
            <tr>
              <th>Product model</th>
              <td>{doc.productModel}</td>
            </tr>
          )}
          {doc.reportDate && (
            <tr>
              <th>Report date</th>
              <td>{doc.reportDate}</td>
            </tr>
          )}
          <tr>
            <th>Author</th>
            <td>{report.createdBy}</td>
          </tr>
          {doc.tags.length > 0 && (
            <tr>
              <th>Tags</th>
              <td>
                {doc.tags.map((tag, i) => (
                  <span className="tag ok" key={i}>
                    {tag}
                  </span>
                ))}
              </td>
            </tr>
          )}
        </tbody>
      </table>
      <h3>I. Purpose.</h3>
      {[doc.purpose, doc.background]
        .filter((line) => line.trim())
        .map((line, i) => (
          <p key={i}>{dash(line)}</p>
        ))}
      {pictures("purpose")}
      <h3>II. Content.</h3>
      {doc.content
        .filter((line) => line.trim())
        .map((line, i) => (
          <p key={i}>{dash(line)}</p>
        ))}
      {doc.changedFactors
        .filter((factor) => factor.name.trim())
        .map((factor, i) => (
          <p key={i}>
            - {factor.name}: {factor.baseline} → {factor.changed}
          </p>
        ))}
      {doc.conditions
        .filter((condition) => condition.details.trim())
        .map((condition, i) => (
          <p key={i}>
            - {condition.label}: {condition.details}
          </p>
        ))}
      {pictures("conditions")}
      <h3>III. Result</h3>
      {tables.map((table, i) => (
        <ResultGrid
          table={{ ...table, title: `${i + 1}. ${table.title}` }}
          key={i}
        />
      ))}
      {doc.resultTables.map((table, i) => (
        <div className="scroll-x" key={i}>
          <h4>
            {tables.length + i + 1}. {table.caption || "Result"}
          </h4>
          <table className="grid report-result-grid">
            <thead>
              <tr>
                {table.columns.map((column, j) => (
                  <th key={j}>{column}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {table.rows.map((row, j) => (
                <tr key={j}>
                  {Array.from(
                    { length: Math.max(row.length, table.columns.length) },
                    (_, k) => (
                      <td key={k}>{row[k] ?? ""}</td>
                    ),
                  )}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      ))}
      {pictures("results")}
      {doc.findings
        .filter((line) => line.trim())
        .map((line, i) => (
          <p key={i}>{dash(line)}</p>
        ))}
      <h3>IV. Conclusion</h3>
      {[doc.conclusion, ...doc.limitations]
        .filter((line) => line.trim())
        .map((line, i) => (
          <p key={i}>{dash(line)}</p>
        ))}
      {doc.nextActions
        .filter((line) => line.trim())
        .map((line, i) => (
          <p key={i}>→ {line}</p>
        ))}
      {pictures("conclusion")}
      {doc.images.some(
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
