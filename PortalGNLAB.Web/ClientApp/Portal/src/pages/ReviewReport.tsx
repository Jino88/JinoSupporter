import { useEffect, useRef, useState } from "react";
import {
  downloadBlob,
  getJson,
  imageData,
  toolPath,
  toolSend,
} from "../toolsApi";
import { useImagePaste } from "./translate/useImagePaste";
import {
  filterReports,
  placePicture,
  removePicture,
} from "./reviewReport/logic";
import {
  freshReport,
  pictureDragType,
  type Report,
  type ReportDoc,
  type ReportIndex,
} from "./reviewReport/types";
import ReportPreview from "./reviewReport/ReportPreview";
import ReportEditor from "./reviewReport/ReportEditor";
import "../tools.css";

export default function ReviewReport() {
  const [report, setReport] = useState<Report>(freshReport);
  const [index, setIndex] = useState<ReportIndex>({
    reports: [],
    languages: [],
    configured: false,
  });
  const [filter, setFilter] = useState("");
  const [request, setRequest] = useState("");
  const [busy, setBusy] = useState(false);
  const [revising, setRevising] = useState(false);
  const [elapsed, setElapsed] = useState(0);
  const [error, setError] = useState("");
  const [status, setStatus] = useState("");
  const [deleteOpen, setDeleteOpen] = useState(false);
  const busyRef = useRef(false);
  busyRef.current = busy;
  const paste = useImagePaste(addPastedImage, setError);
  const controller = useRef<AbortController | null>(null);
  const hasInput = !!report.inputText.trim() || report.images.length > 0;

  async function refresh() {
    setIndex(await getJson<ReportIndex>(toolPath("review-report")));
  }
  useEffect(() => {
    refresh().catch((e: Error) => setError(e.message));
    return () => controller.current?.abort();
  }, []);
  useEffect(() => {
    if (!busy) return;
    setElapsed(0);
    const timer = window.setInterval(
      () => setElapsed((value) => value + 1),
      1000,
    );
    return () => window.clearInterval(timer);
  }, [busy]);

  function update(change: Partial<Report>) {
    setReport((current) => ({ ...current, ...change }));
  }
  function addPastedImage(mediaType: string, data: string) {
    if (busyRef.current) return;
    setReport((current) => ({
      ...current,
      images: [
        ...current.images,
        {
          no: Math.max(0, ...current.images.map((image) => image.no)) + 1,
          fileName: `paste-${Date.now()}.png`,
          mediaType,
          data,
        },
      ],
    }));
  }
  async function addFiles(files: FileList | null) {
    setError("");
    const selected = Array.from(files ?? []);
    if (selected.length > 30)
      return setError("Select at most 30 pictures at once.");
    for (const file of selected) {
      if (!file.type.startsWith("image/")) continue;
      if (file.size > 15 * 1024 * 1024) {
        setError(`${file.name}: maximum file size is 15 MB.`);
        continue;
      }
      try {
        const data = await imageData(file);
        setReport((current) => ({
          ...current,
          images: [
            ...current.images,
            {
              no: Math.max(0, ...current.images.map((image) => image.no)) + 1,
              fileName: file.name,
              mediaType: file.type,
              data,
            },
          ],
        }));
      } catch (e) {
        setError((e as Error).message);
      }
    }
  }
  function newReport() {
    setReport(freshReport());
    setRequest("");
    setError("");
    setStatus("");
  }
  async function openReport(id: string) {
    setError("");
    setStatus("");
    setRequest("");
    try {
      setReport(
        await getJson<Report>(
          toolPath(`review-report/${encodeURIComponent(id)}`),
        ),
      );
    } catch (e) {
      setError((e as Error).message);
      await refresh();
    }
  }
  async function saveOrGenerate(action: "save" | "generate" | "revise") {
    if (busyRef.current) return;
    setBusy(true);
    setRevising(action === "revise");
    setError("");
    setStatus("");
    controller.current = new AbortController();
    try {
      setReport(
        await toolSend<Report>(
          `review-report/${action}`,
          { report, request },
          "POST",
          controller.current.signal,
        ),
      );
      setRequest("");
      await refresh();
      setStatus(action === "save" ? "Saved." : "Report written and saved.");
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setBusy(false);
      controller.current = null;
    }
  }
  async function deleteReport() {
    setBusy(true);
    try {
      await toolSend(
        `review-report/${encodeURIComponent(report.id)}`,
        undefined,
        "DELETE",
      );
      setDeleteOpen(false);
      newReport();
      await refresh();
      setStatus("Deleted.");
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setBusy(false);
    }
  }
  async function copyMarkdown() {
    try {
      const result = await toolSend<{ text: string }>(
        "review-report/markdown",
        { report },
      );
      await navigator.clipboard.writeText(result.text);
      setStatus("Markdown copied.");
    } catch (e) {
      setError((e as Error).message);
    }
  }
  async function exportExcel() {
    try {
      const blob = await toolSend<Blob>("review-report/export", { report });
      const title = (report.doc?.title || "review-report").replace(
        /[<>:"/\\|?*]/g,
        "_",
      );
      downloadBlob(blob, `${title}.xlsx`);
    } catch (e) {
      setError((e as Error).message);
    }
  }
  function movePicture(no: number, section: string) {
    if (!report.doc) return;
    update({ doc: placePicture(report.doc, no, section) });
    setStatus(`IMG ${no} placed in ${section}. Save to keep it.`);
  }
  function editDoc(doc: ReportDoc) {
    update({ doc });
    setStatus("Report edited. Save to keep your changes.");
  }
  function deletePicture(no: number) {
    setReport((current) => removePicture(current, no));
    setStatus(`IMG ${no} deleted. Save to keep it.`);
  }

  return (
    <div className="tools-page">
      <header className="page-head">
        <h1>Review Report</h1>
        <p className="muted">
          Paste notes, Excel cells and pictures. AI writes your report for
          WorkHub Review Data.
        </p>
      </header>
      {error && (
        <div className="alert" role="alert">
          {error}
        </div>
      )}
      <p role="status">
        {busy ? `${revising ? "Changing" : "Working"}… ${elapsed} s` : status}
      </p>
      <div className="report-layout">
        <aside className="card flush report-list">
          <h2 className="split">
            Reports
            <button
              className="btn sm primary"
              disabled={busy}
              onClick={newReport}
            >
              + New
            </button>
          </h2>
          <label className="tool-field report-filter">
            Filter reports
            <input
              className="inp"
              value={filter}
              onChange={(e) => setFilter(e.target.value)}
            />
          </label>
          {filterReports(index.reports, filter).map((item) => (
            <button
              className={`report-list-item ${item.id === report.id ? "on" : ""}`}
              disabled={busy}
              key={item.id}
              onClick={() => void openReport(item.id)}
              title={item.title}
            >
              <strong>{item.title}</strong>
              <span className="muted small">
                {item.createdBy} ·{" "}
                {item.updatedAt.slice(0, 16).replace("T", " ")}
              </span>
            </button>
          ))}
          {index.reports.length === 0 && (
            <p className="muted report-filter">No saved reports yet.</p>
          )}
        </aside>
        <section className="card">
          <fieldset className="tool-fieldset" disabled={busy}>
            <label className="tool-field">
              Content
              <textarea
                className="inp tool-textarea"
                rows={12}
                value={report.inputText}
                onChange={(e) => update({ inputText: e.target.value })}
                placeholder="What was tested and why, conditions, results and your conclusion. Paste Excel cells here."
              />
            </label>
            <label className="tool-field">
              Pictures
              <input
                type="file"
                multiple
                accept="image/*"
                onChange={(e) => void addFiles(e.target.files)}
              />
            </label>
            <p className="muted small">
              Ctrl+V anywhere on the page, or pick files. Up to 15 MB per file.
            </p>
            <button className="btn sm" onClick={paste.openCapture}>
              Enable image paste
            </button>
            <div className="tool-thumbnails">
              {report.images.map((image) => (
                <figure
                  key={image.no}
                  title={`IMG ${image.no} · ${image.fileName} · ${Math.round((image.data.length * 0.75) / 1024)} KB${report.doc ? " · drag into the report" : ""}`}
                  draggable={!!report.doc && !busy}
                  onDragStart={(e) => {
                    e.dataTransfer.setData(pictureDragType, String(image.no));
                    e.dataTransfer.effectAllowed = "move";
                  }}
                >
                  <img
                    src={`data:${image.mediaType};base64,${image.data}`}
                    alt={`IMG ${image.no}`}
                  />
                  <button
                    className="btn sm"
                    onClick={() =>
                      setReport((current) => removePicture(current, image.no))
                    }
                  >
                    Remove IMG {image.no}
                  </button>
                </figure>
              ))}
            </div>
            <p className="muted small">
              The report is always written in English.
            </p>
            <label className="tool-field">
              Instruction to AI (optional)
              <input
                className="inp"
                value={report.extraInstruction}
                onChange={(e) => update({ extraInstruction: e.target.value })}
              />
            </label>
            {!index.configured && (
              <div className="alert">
                Set the Codex API key in Settings to write a report.
              </div>
            )}
            <div className="form-row">
              <button
                className="btn sm primary"
                disabled={!hasInput || !index.configured}
                onClick={() => void saveOrGenerate("generate")}
              >
                {report.doc ? "Rewrite" : "Write report"}
              </button>
              <button
                className="btn sm"
                disabled={!hasInput}
                onClick={() => void saveOrGenerate("save")}
              >
                Save
              </button>
              {report.id && (
                <button
                  className="btn sm danger"
                  onClick={() => setDeleteOpen(true)}
                >
                  Delete
                </button>
              )}
            </div>
            {report.doc && (
              <section className="report-revise">
                <h2>Ask AI to change this report</h2>
                <ol>
                  {report.revisions.map((revision, i) => (
                    <li key={i}>
                      {revision.request}{" "}
                      <span className="muted small">
                        {revision.at.slice(5, 16).replace("T", " ")}
                      </span>
                    </li>
                  ))}
                </ol>
                <label className="tool-field">
                  Change request
                  <textarea
                    className="inp tool-textarea"
                    rows={3}
                    value={request}
                    onChange={(e) => setRequest(e.target.value)}
                  />
                </label>
                <button
                  className="btn sm primary"
                  disabled={!request.trim() || !index.configured}
                  onClick={() => void saveOrGenerate("revise")}
                >
                  Apply change
                </button>
              </section>
            )}
          </fieldset>
        </section>
        <section className="card report-preview">
          {report.doc && (
            <div className="form-row report-preview-bar">
              <span className="muted small">
                {report.model} ·{" "}
                {report.updatedAt.slice(0, 16).replace("T", " ")}
              </span>
              <button
                className="btn sm"
                disabled={busy}
                onClick={() => void copyMarkdown()}
              >
                Copy Markdown
              </button>
              <button
                className="btn sm primary"
                disabled={busy}
                onClick={() => void exportExcel()}
              >
                Excel for WorkHub
              </button>
            </div>
          )}
          <fieldset className="tool-fieldset" disabled={busy}>
            <ReportPreview
              report={report}
              editable={!busy}
              onChange={editDoc}
              onMove={movePicture}
              onRemovePicture={deletePicture}
              onError={setError}
            />
            {report.doc && (
              <ReportEditor
                doc={report.doc}
                onChange={(doc) => update({ doc })}
              />
            )}
          </fieldset>
        </section>
      </div>
      {deleteOpen && (
        <div className="modal-back">
          <section
            className="modal"
            role="dialog"
            aria-modal="true"
            aria-labelledby="delete-report-title"
          >
            <h2 id="delete-report-title">Delete report?</h2>
            <p>{report.doc?.title || report.title || "This report"}</p>
            <div className="modal-actions">
              <button
                className="btn sm"
                disabled={busy}
                onClick={() => setDeleteOpen(false)}
              >
                Cancel
              </button>
              <button
                className="btn sm danger"
                disabled={busy}
                onClick={() => void deleteReport()}
              >
                Delete
              </button>
            </div>
          </section>
        </div>
      )}
    </div>
  );
}
