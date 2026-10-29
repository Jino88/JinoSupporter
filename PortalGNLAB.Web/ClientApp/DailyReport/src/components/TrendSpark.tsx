import { useId } from "react";
import { direction } from "../format";
import { formatPpm } from "../format";
import { buildSparkGeometry } from "../logic";

interface TrendSparkProps {
  modelName: string;
  noun: "week" | "day";
  values: number[];
  baselinePpm: number;
  latestPpm: number;
  deltaPpm: number;
}

const points = (segment: Array<{ x: number; y: number }>) =>
  segment.map((point) => `${point.x.toFixed(1)},${point.y.toFixed(1)}`).join(" ");

export function TrendSpark(props: TrendSparkProps) {
  const geometry = buildSparkGeometry(props.values, props.baselinePpm);
  const id = `drspark${useId().replaceAll(":", "")}`;
  if (!geometry) return <span className="dr-none">—</span>;

  const marker = direction(props.deltaPpm);
  return (
    <div className="dr-spark-wrap">
      <svg
        className="dr-spark"
        width={geometry.width}
        height={geometry.height}
        viewBox={`0 0 ${geometry.width} ${geometry.height}`}
        role="img"
        aria-label={`${props.modelName} ${props.noun}ly PPM trend; zero-based Y scale; latest on the left, older periods on the right; dashed line is the baseline`}
      >
        <line x1="0" y1={geometry.zeroY} x2={geometry.width} y2={geometry.zeroY} stroke="var(--rule)" strokeWidth="1" />
        {geometry.baselineY === null ? null : (
          <defs>
            <clipPath id={`${id}w`}><rect x="0" y="0" width={geometry.width} height={geometry.baselineY} /></clipPath>
            <clipPath id={`${id}b`}><rect x="0" y={geometry.baselineY} width={geometry.width} height={geometry.height - geometry.baselineY} /></clipPath>
          </defs>
        )}
        {geometry.baselineY === null ? null : geometry.segments.filter((segment) => segment.length > 1).map((segment, index) => {
          const first = segment[0];
          const last = segment[segment.length - 1];
          const path = `M ${first.x.toFixed(1)},${geometry.baselineY!.toFixed(1)} ${segment.map((point) => `L ${point.x.toFixed(1)},${point.y.toFixed(1)}`).join(" ")} L ${last.x.toFixed(1)},${geometry.baselineY!.toFixed(1)} Z`;
          return (
            <g key={index}>
              <path d={path} fill="var(--worse-wash)" clipPath={`url(#${id}w)`} />
              <path d={path} fill="var(--better-wash)" clipPath={`url(#${id}b)`} />
            </g>
          );
        })}
        {geometry.baselineY === null ? null : (
          <line x1="0" y1={geometry.baselineY} x2={geometry.width} y2={geometry.baselineY} stroke="var(--rule-hard)" strokeWidth="1" strokeDasharray="2 2" />
        )}
        {geometry.segments.filter((segment) => segment.length > 1).map((segment, index) => (
          <polyline key={index} points={points(segment)} fill="none" stroke="var(--ink-2)" strokeWidth="1.4" strokeLinejoin="round" strokeLinecap="round" />
        ))}
        <circle
          cx={geometry.latestPoint.x}
          cy={geometry.latestPoint.y}
          r="2.9"
          fill={`var(--${marker === "flat" ? "ink-2" : marker})`}
          stroke="var(--paper)"
          strokeWidth="1.4"
        />
      </svg>
      <span className="dr-spark-key">{formatPpm(props.latestPpm)} → {props.baselinePpm > 0 ? formatPpm(props.baselinePpm) : "—"}</span>
    </div>
  );
}
