export interface Status { text: string; error?: boolean; success?: boolean }

/** Status line under the page header: spinner while busy, dismiss button when idle. */
export default function StatusBar({ status, busy, onDismiss }: { status: Status | null; busy: boolean; onDismiss: () => void }) {
  if (!busy && !status?.text) return null;
  return (
    <div className={"bm-status" + (status?.error ? " error" : status?.success ? " success" : "")} role={status?.error ? "alert" : "status"}>
      {busy && <span className="bm-spin" />}
      <span className="grow">{status?.text || "Working..."}</span>
      {!busy && <button type="button" className="bm-x" title="Dismiss" onClick={onDismiss}>×</button>}
    </div>
  );
}
