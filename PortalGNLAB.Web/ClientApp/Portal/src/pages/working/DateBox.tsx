import { useEffect, useState } from "react";
import { dayText, readDayText } from "./dateText";

/**
 * Date box of Working: the date typed as yyMMdd (261112) and kept as yyyy-MM-dd. What is not a day yet is passed on as
 * typed, so the form shows that it must be a date; a date set from outside (e.g. Done ticked) is shown again.
 */
export default function DateBox(props: {
  value: string; onChange: (v: string) => void; label: string; className?: string; autoFocus?: boolean;
}) {
  const [text, setText] = useState(dayText(props.value));
  useEffect(() => { if (readDayText(text) !== props.value) setText(dayText(props.value)); }, [props.value]);
  return (
    <input className={"inp wk-datebox" + (props.className ? " " + props.className : "")} value={text} placeholder="yyMMdd"
      inputMode="numeric" maxLength={10} aria-label={props.label} title="yyMMdd, e.g. 261112" autoFocus={props.autoFocus}
      onChange={(e) => { setText(e.target.value); props.onChange(readDayText(e.target.value) ?? e.target.value); }} />
  );
}
