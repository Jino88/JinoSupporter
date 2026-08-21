import type { Granularity } from "../logic";

interface GranularitySegmentProps {
  value: Granularity;
  onChange: (value: Granularity) => void;
}

export function GranularitySegment({ value, onChange }: GranularitySegmentProps) {
  return (
    <div className="dr-segment" aria-label="Report granularity">
      <button type="button" aria-pressed={value === "week"} onClick={() => onChange("week")}>WEEKLY</button>
      <button type="button" aria-pressed={value === "day"} onClick={() => onChange("day")}>DAILY</button>
    </div>
  );
}
