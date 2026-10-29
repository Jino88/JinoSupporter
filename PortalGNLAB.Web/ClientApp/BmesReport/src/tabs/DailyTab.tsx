import type { DailyTabData, ReportPeriod } from "../contract";
import { formatCount, formatPpm } from "../format";
import { filterDailyReasonRows, selectVisiblePeriods, sumVisiblePpmByPeriod, type ViewerPreferences } from "../logic";
import { dim, EmptySection, indentStyle, ReportTable, Section } from "../components/ReportTable";

export function DailyTab({ data, preferences }: { data: DailyTabData; preferences: ViewerPreferences }) {
  const periods = selectVisiblePeriods(data.periods, preferences);
  return (
    <div className="tab-content">
      <Section title="Daily summary" count={data.summaryRows.length}>
        {data.summaryRows.length === 0 ? <EmptySection /> : (
          <ReportTable
            caption="Daily Input, NG and PPM by hierarchy"
            headers={["Item", "Level", ...periods.flatMap(() => ["Input", "NG", "PPM"])]}
            groups={{ lead: 2, items: periods.map((period) => ({ label: period.header, span: 3 })) }}
          >
            {data.summaryRows.map((row) => (
              <tr key={row.rowId} className={row.level === "total" ? "total-row" : ""}>
                <th scope="row" style={indentStyle(row.depth)}>{row.display}</th>
                <td>{row.level}{row.processType ? ` · ${row.processType}` : ""}</td>
                {periods.flatMap((period) => [
                  <td key={`${row.rowId}-${period.key}-input`} className="numeric">{dim(formatCount(row.inputByPeriod[period.key]))}</td>,
                  <td key={`${row.rowId}-${period.key}-ng`} className="numeric">{dim(formatCount(row.ngByPeriod[period.key]))}</td>,
                  <td key={`${row.rowId}-${period.key}-ppm`} className="numeric">{dim(formatPpm(row.ppmByPeriod[period.key]))}</td>,
                ])}
              </tr>
            ))}
          </ReportTable>
        )}
      </Section>

      <Section title="Defect causes by model" count={data.modelSections.length}>
        {data.modelSections.length === 0 ? <EmptySection /> : data.modelSections.map((section) => {
          const rows = filterDailyReasonRows(
            section.reasonRows,
            data.referenceDatePeriodKey,
            preferences.minimumPpm,
          );
          const visibleTotal = sumVisiblePpmByPeriod(rows, periods);
          return (
            <section className="model-section" key={section.id} aria-labelledby={`daily-model-${section.id}`}>
              <div className="model-heading">
                <h3 id={`daily-model-${section.id}`}>{section.groupName} · {section.modelName}</h3>
                <span>{section.lineShiftCount.toLocaleString("en-US")} line shifts</span>
              </div>
              {rows.filter((row) => !row.isTotal).length === 0 ? (
                <EmptySection message={`No causes at Minimum PPM ${formatPpm(preferences.minimumPpm)} or higher.`} />
              ) : (
                <ReportTable
                  caption={`${section.modelName} defect cause PPM`}
                  headers={["Rank", "Cause", "Process", "Process name", "NG", ...periods.map((period) => `${period.header} PPM`)]}
                >
                  {rows.map((row) => (
                    <tr key={row.rowId} className={row.isTotal ? "total-row" : ""}>
                      <td className="numeric">{dim(row.rank ?? "-")}</td>
                      <th scope="row">{row.reason}</th>
                      <td>{dim(row.processType ?? "-")}</td>
                      <td>{dim(row.processName ?? "-")}</td>
                      <td>{dim(row.ngName ?? "-")}</td>
                      {periods.map((period) => (
                        <td className="numeric" key={period.key}>{dim(formatPpm(row.ppmByPeriod[period.key]))}</td>
                      ))}
                    </tr>
                  ))}
                  <tr className="visible-total-row">
                    <td />
                    <th scope="row" colSpan={4}>Visible total (current filter)</th>
                    {periods.map((period) => (
                      <td className="numeric" key={period.key}>{dim(formatPpm(visibleTotal[period.key]))}</td>
                    ))}
                  </tr>
                </ReportTable>
              )}
            </section>
          );
        })}
      </Section>
      {periods.length === 0 && <EmptySection message="All period counts are 0. Increase them in the options above." />}
    </div>
  );
}

export const dailyPeriodHeaders = (periods: ReportPeriod[]) => periods.map((period) => period.header);
