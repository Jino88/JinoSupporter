import { useEffect, type ReactNode } from "react";

/**
 * AI input of a tab in a popup: its title, a close button (Esc too) that asks first when what was typed or the AI answer
 * would be lost, and the AI box as the body.
 */
export default function AiPopup(props: { title: string; dirty: boolean; onClose: () => void; children: ReactNode }) {
  const close = () => {
    if (!props.dirty || window.confirm("Close AI input? What you typed and the AI answer are lost.")) props.onClose();
  };
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => { if (e.key === "Escape") close(); };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  });
  return (
    <div className="modal-back">
      <div className="modal wk-ai-pop" role="dialog" aria-modal="true" aria-label={props.title}>
        <div className="wk-ai-pop-h">
          <h2>{props.title}</h2>
          <button className="btn sm" aria-label="Close" onClick={close}>×</button>
        </div>
        <div className="wk-ai-pop-b">{props.children}</div>
      </div>
    </div>
  );
}
