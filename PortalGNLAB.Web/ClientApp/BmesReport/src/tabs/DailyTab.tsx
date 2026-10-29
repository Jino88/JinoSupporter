import type { DailyTabData, ReportPeriod } from "../contract";
import { formatCount, formatPpm } from "../format";
import { filterDailyReasonRows, selectVisiblePeriods, sumVisiblePpmByPeriod, type ViewerPreferences } from "../logic";
import { EmptySection, indentStyle, ReportTable, Section } from "../components/ReportTable";

export function DailyTab({ data, preferences }: { data: DailyTabData; preferences: ViewerPreferences }) {
  const periods = selectVisiblePeriods(data.periods, preferences);
  return (
    <div className="tab-content">
      <Section title="Daily 요약" count={data.summaryRows.length}>
        {data.summaryRows.length === 0 ? <EmptySection /> : (
          <ReportTable
            caption="Daily 계층별 Input, NG, PPM 요약"
            headers={["항목", "Level", ...periods.flatMap((period) => [
              `${period.header} Input`, `${period.header} NG`, `${period.header} PPM`,
            ])]}
          >
            {data.summaryRows.map((row) => (
              <tr key={row.rowId} className={row.level === "total" ? "total-row" : ""}>
                <th scope="row" style={indentStyle(row.depth)}>{row.display}</th>
                <td>{row.level}{row.processType ? ` · ${row.processType}` : ""}</td>
                {periods.flatMap((period) => [
                  <td key={`${row.rowId}-${period.key}-input`} className="numeric">{formatCount(row.inputByPeriod[period.key])}</td>,
                  <td key={`${row.rowId}-${period.key}-ng`} className="numeric">{formatCount(row.ngByPeriod[period.key])}</td>,
                  <td key={`${row.rowId}-${period.key}-ppm`} className="numeric">{formatPpm(row.ppmByPeriod[period.key])}</td>,
                ])}
              </tr>
            ))}
          </ReportTable>
        )}
      </Section>

      <Section title="모델별 불량 원인" count={data.modelSections.length}>
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
                <EmptySection message={`Minimum PPM ${formatPpm(preferences.minimumPpm)} 이상인 원인이 없습니다.`} />
              ) : (
                <ReportTable
                  caption={`${section.modelName} 불량 원인 PPM`}
                  headers={["순위", "원인", "Process", "공정", "NG", ...periods.map((period) => `${period.header} PPM`)]}
                >
                  {rows.map((row) => (
                    <tr key={row.rowId} className={row.isTotal ? "total-row" : ""}>
                      <td className="numeric">{row.rank ?? "-"}</td>
                      <th scope="row">{row.reason}</th>
                      <td>{row.processType ?? "-"}</td>
                      <td>{row.processName ?? "-"}</td>
                      <td>{row.ngName ?? "-"}</td>
                      {periods.map((period) => (
                        <td className="numeric" key={period.key}>{formatPpm(row.ppmByPeriod[period.key])}</td>
                      ))}
                    </tr>
                  ))}
                  <tr className="visible-total-row">
                    <td />
                    <th scope="row" colSpan={4}>현재 필터 표시 합계</th>
                    {periods.map((period) => (
                      <td className="numeric" key={period.key}>{formatPpm(visibleTotal[period.key])}</td>
                    ))}
                  </tr>
                </ReportTable>
              )}
            </section>
          );
        })}
      </Section>
      {periods.length === 0 && <EmptySection message="기간 표시 개수가 모두 0입니다. 상단 필터에서 기간을 늘려 주세요." />}
    </div>
  );
}

export const dailyPeriodHeaders = (periods: ReportPeriod[]) => periods.map((period) => period.header);
