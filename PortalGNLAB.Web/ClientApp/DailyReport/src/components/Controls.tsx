import type { GroupFilter } from "../logic";

const groups: Array<{ value: GroupFilter; label: string }> = [
  { value: "ALL", label: "All" },
  { value: "SPK", label: "SPK" },
  { value: "MODULE", label: "Module" },
  { value: "UNIT", label: "Unit" },
  { value: "TWS", label: "TWS" },
];

interface ControlsProps {
  group: GroupFilter;
  onGroupChange: (value: GroupFilter) => void;
}

export function Controls(props: ControlsProps) {
  return (
    <div className="dr-controls">
      <div className="dr-chips" aria-label="Model group filter">
        <span className="dr-control-label">GROUP</span>
        {groups.map((group) => (
          <button
            className="dr-chip"
            key={group.value}
            type="button"
            aria-pressed={props.group === group.value}
            onClick={() => props.onGroupChange(group.value)}
          >
            {group.label}
          </button>
        ))}
      </div>
    </div>
  );
}
