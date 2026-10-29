import { useEffect, useId } from "react";
import { loadLegacyTools } from "../../toolsApi";
export interface Payload {
  kind: string;
  title: string;
  stats: { name: string; value: number }[];
  statGroups: { label: string; stats: { name: string; value: number }[] }[];
}
interface Renderer {
  render(id: string, payload: Payload): void;
  destroy(id: string): void;
}
function renderer() {
  return (window as unknown as { graphMaker?: Renderer }).graphMaker;
}
export function formatStat(value: number): string {
  return Math.abs(value) >= 1000
    ? value.toLocaleString("en-US", {
        minimumFractionDigits: 1,
        maximumFractionDigits: 1,
      })
    : String(Number(value.toFixed(4)));
}
export default function Charts({
  payloads,
  onError,
}: {
  payloads: Payload[];
  onError: (message: string) => void;
}) {
  const prefix = useId().replaceAll(":", "");
  useEffect(() => {
    let active = true;
    const ids = payloads.map((_, index) => `graph-${prefix}-${index}`);
    loadLegacyTools()
      .then(() => {
        if (active)
          payloads.forEach((payload, index) =>
            renderer()!.render(ids[index], payload),
          );
      })
      .catch((e: Error) => {
        if (active) onError(e.message);
      });
    return () => {
      active = false;
      ids.forEach((id) => renderer()?.destroy(id));
    };
  }, [payloads, prefix]);
  return (
    <div className="graph-chart-stack">
      {payloads.map((payload, index) => (
        <section className="card" key={index}>
          <h2>{payload.title}</h2>
          <div className="graph-chart-frame">
            <canvas id={`graph-${prefix}-${index}`} />
          </div>
          <div className="graph-stat-panel">
            {payload.statGroups.map((group, i) => (
              <section className="graph-stat-card" key={i}>
                <h3>{group.label}</h3>
                <dl>
                  {group.stats.map((stat, j) => (
                    <div key={j}>
                      <dt>{stat.name}</dt>
                      <dd>{formatStat(stat.value)}</dd>
                    </div>
                  ))}
                </dl>
              </section>
            ))}
          </div>
        </section>
      ))}
    </div>
  );
}
