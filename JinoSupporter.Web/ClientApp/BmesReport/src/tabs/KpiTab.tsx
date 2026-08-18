import type { KpiTabData } from "../contract";
import { formatMetric } from "../format";
import { selectVisiblePeriods, type ViewerPreferences } from "../logic";
import { EmptySection, ReportTable, Section } from "../components/ReportTable";

export function KpiTab({ data, preferences }: { data: KpiTabData; preferences: ViewerPreferences }) {
  const periods = selectVisiblePeriods(data.periods, preferences);
  return (
    <div className="tab-content kpi-grid">
      {data.metrics.length === 0 ? <EmptySection /> : data.metrics.map((metric) => (
        <Section key={metric.id} title={metric.name} count={metric.lines.length}>
          <dl className="metric-meta">
            <div><dt>Type</dt><dd>{metric.type}</dd></div>
            <div><dt>Baseline</dt><dd>{formatMetric(metric.baselineValue, metric.unit)}</dd></div>
            <div><dt>Target</dt><dd>{formatMetric(metric.targetValue, metric.unit)}</dd></div>
          </dl>
          {metric.lines.length === 0 ? <EmptySection /> : (
            <ReportTable caption={`${metric.name} KPI`} headers={[
              "지표", "종류", "연간", ...periods.map((period) => period.header),
            ]}>
              {metric.lines.map((line, index) => (
                <tr key={`${metric.id}-${line.kind}-${index}`}>
                  <th scope="row">{line.label}</th>
                  <td>{line.kind}</td>
                  <td className="numeric">{formatMetric(line.annualValue, line.unit)}</td>
                  {periods.map((period) => (
                    <td className="numeric" key={period.key}>{formatMetric(line.valuesByPeriod[period.key], line.unit)}</td>
                  ))}
                </tr>
              ))}
            </ReportTable>
          )}
        </Section>
      ))}
    </div>
  );
}
