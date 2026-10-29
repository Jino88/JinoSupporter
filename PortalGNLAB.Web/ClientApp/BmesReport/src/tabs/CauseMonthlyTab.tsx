import type { CauseMonthlyTabData } from "../contract";
import { formatPpm, formatShareRatio } from "../format";
import { selectVisiblePeriods, type ViewerPreferences } from "../logic";
import { dim, EmptySection, ReportTable, Section } from "../components/ReportTable";

export function CauseMonthlyTab({ data, preferences }: { data: CauseMonthlyTabData; preferences: ViewerPreferences }) {
  const periods = selectVisiblePeriods(data.periods, preferences);
  return (
    <div className="tab-content">
      <Section title="Cause share" count={data.rows.length}>
        {data.rows.length === 0 ? <EmptySection /> : (
          <ReportTable caption="Cause share and weighted PPM by model" numericFrom={6} numericCols={[4]} headers={[
            "Model", "Type", "Process", "NG", "No.", "Cause", "Share",
            ...periods.flatMap(() => ["PPM", "Weighted PPM"]),
          ]} groups={{ lead: 7, items: periods.map((period) => ({ label: period.header, span: 2 })) }}>
            {data.rows.map((row) => (
              <tr key={row.rowId} className={row.isSubtotal ? "total-row" : ""}>
                <th scope="row">{row.model}</th>
                <td>{dim(row.type ?? "-")}</td>
                <td>{dim(row.process ?? "-")}</td>
                <td>{dim(row.ngName ?? "-")}</td>
                <td className="numeric">{dim(row.number ?? "-")}</td>
                <td>{dim(row.cause ?? "-")}</td>
                <td className="numeric">{dim(formatShareRatio(row.shareRatio))}</td>
                {periods.flatMap((period) => [
                  <td className="numeric" key={`${period.key}-ppm`}>{dim(formatPpm(row.ppmByPeriod[period.key]))}</td>,
                  <td className="numeric" key={`${period.key}-weighted`}>{dim(formatPpm(row.weightedPpmByPeriod[period.key]))}</td>,
                ])}
              </tr>
            ))}
          </ReportTable>
        )}
      </Section>

      <Section title="Model monthly PPM" count={data.modelMonthlyRows.length}>
        {data.modelMonthlyRows.length === 0 ? <EmptySection /> : (
          <ReportTable caption="Monthly PPM by model" numericFrom={1} headers={["Model", ...periods.map((period) => `${period.header} PPM`)]}>
            {data.modelMonthlyRows.map((row) => (
              <tr key={row.model}>
                <th scope="row">{row.model}</th>
                {periods.map((period) => <td className="numeric" key={period.key}>{dim(formatPpm(row.ppmByPeriod[period.key]))}</td>)}
              </tr>
            ))}
          </ReportTable>
        )}
      </Section>
    </div>
  );
}
