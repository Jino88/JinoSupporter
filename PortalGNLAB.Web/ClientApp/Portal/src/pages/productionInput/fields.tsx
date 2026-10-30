import { useEffect, useState, type KeyboardEvent } from "react";

const blurOnEnter = (e: KeyboardEvent<HTMLInputElement>) => {
  if (e.key === "Enter") e.currentTarget.blur();
};

/**
 * Number box that keeps what is typed and commits on blur / Enter. Blank commits null when
 * `allowEmpty`, otherwise 0; a negative or non-number entry goes back to the last value.
 */
export function NumInput(props: {
  value: number | null;
  onCommit: (value: number | null) => void;
  className?: string;
  decimal?: boolean;
  allowEmpty?: boolean;
  disabled?: boolean;
  title?: string;
  label?: string;
}) {
  const { value, onCommit, decimal, allowEmpty } = props;
  const shown = value == null ? "" : String(value);
  const [text, setText] = useState(shown);
  useEffect(() => setText(shown), [shown]);
  const commit = () => {
    const t = text.trim();
    if (t === "") {
      const empty = allowEmpty ? null : 0;
      if (empty !== value) onCommit(empty);
      else setText(shown);
      return;
    }
    const v = Number(t.replace(/,/g, ""));
    if (!Number.isFinite(v) || v < 0) { setText(shown); return; }
    const next = decimal ? v : Math.round(v);
    if (next !== value) onCommit(next);
    else setText(shown);
  };
  return (
    <input
      className={"pi-ni " + (props.className ?? "")}
      inputMode={decimal ? "decimal" : "numeric"}
      value={text}
      disabled={props.disabled}
      title={props.title}
      aria-label={props.label}
      onChange={(e) => setText(e.target.value)}
      onBlur={commit}
      onKeyDown={blurOnEnter}
    />
  );
}

/** Text box that commits on blur / Enter. */
export function TextInput(props: {
  value: string;
  onCommit: (value: string) => void;
  className?: string;
  placeholder?: string;
  label?: string;
}) {
  const [text, setText] = useState(props.value);
  useEffect(() => setText(props.value), [props.value]);
  return (
    <input
      className={"pi-ti " + (props.className ?? "")}
      value={text}
      placeholder={props.placeholder}
      aria-label={props.label}
      onChange={(e) => setText(e.target.value)}
      onBlur={() => { if (text !== props.value) props.onCommit(text); }}
      onKeyDown={blurOnEnter}
    />
  );
}
