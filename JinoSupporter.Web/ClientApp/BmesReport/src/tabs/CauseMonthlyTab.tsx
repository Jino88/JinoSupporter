import type { CauseMonthlyTabData } from "../contract";
import { formatPpm, formatShareRatio } from "../format";
import { selectVisiblePeriods, type ViewerPreferences } from "../logic";
import { EmptySection, ReportTable, Section } from "../components/ReportTable";

export function CauseMonthlyTab({ data, preferences }: { data: CauseMonthlyTabData; preferences: ViewerPreferences }) {
  const periods = selectVisiblePeriods(data.periods, preferences);
  return (
    <div className="tab-content">
      <Section title="원인 비중" count={data.rows.length}>
        {data.rows.length === 0 ? <EmptySection /> : (
          <ReportTable caption="모델별 원인 비중과 가중 PPM" headers={[
            "Model", "Type", "Process", "NG", "No.", "Cause", "Share",
            ...periods.flatMap((period) => [`${period.header} PPM`, `${period.header} 가중 PPM`]),
          ]}>
            {data.rows.map((row) => (
              <tr key={row.rowId} className={row.isSubtotal ? "total-row" : ""}>
                <th scope="row">{row.model}</th>
                <td>{row.type ?? "-"}</td>
                <td>{row.process ?? "-"}</td>
                <td>{row.ngName ?? "-"}</td>
                <td className="numeric">{row.number ?? "-"}</td>
                <td>{row.cause ?? "-"}</td>
                <td className="numeric">{formatShareRatio(row.shareRatio)}</td>
                {periods.flatMap((period) => [
                  <td className="numeric" key={`${period.key}-ppm`}>{formatPpm(row.ppmByPeriod[period.key])}</td>,
                  <td className="numeric" key={`${period.key}-weighted`}>{formatPpm(row.weightedPpmByPeriod[period.key])}</td>,
                ])}
              </tr>
            ))}
          </ReportTable>
        )}
      </Section>

      <Section title="모델 월별 PPM" count={data.modelMonthlyRows.length}>
        {data.modelMonthlyRows.length === 0 ? <EmptySection /> : (
          <ReportTable caption="모델별 월간 PPM" headers={["Model", ...periods.map((period) => `${period.header} PPM`)]}>
            {data.modelMonthlyRows.map((row) => (
              <tr key={row.model}>
                <th scope="row">{row.model}</th>
                {periods.map((period) => <td className="numeric" key={period.key}>{formatPpm(row.ppmByPeriod[period.key])}</td>)}
              </tr>
            ))}
          </ReportTable>
        )}
      </Section>
    </div>
  );
}
