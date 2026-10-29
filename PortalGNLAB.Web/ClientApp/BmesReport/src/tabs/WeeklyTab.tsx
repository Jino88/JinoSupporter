import type { WeeklyTabData } from "../contract";
import { formatPercent, formatPpm } from "../format";
import { selectVisiblePeriods, type ViewerPreferences } from "../logic";
import { dim, EmptySection, indentStyle, ReportTable, Section } from "../components/ReportTable";

export function WeeklyTab({ data, preferences }: { data: WeeklyTabData; preferences: ViewerPreferences }) {
  const periods = selectVisiblePeriods(data.periods, preferences);
  return (
    <div className="tab-content">
      <Section title="Weekly target" count={data.targetRows.length}>
        {data.targetRows.length === 0 ? <EmptySection /> : (
          <ReportTable caption="Weekly target and achievement" numericFrom={2} headers={[
            "Item", "Kind", "Baseline PPM", "Target PPM", "Achievement", ...periods.map((period) => `${period.header} PPM`),
          ]}>
            {data.targetRows.map((row) => (
              <tr key={row.rowId} className={row.rowKind === "group" ? "total-row" : ""}>
                <th scope="row" style={indentStyle(row.depth)}>{row.display}</th>
                <td>{row.rowKind}{row.isLineShift ? " · line shift" : ""}</td>
                <td className="numeric">{dim(formatPpm(row.baselinePpm))}</td>
                <td className="numeric">{dim(formatPpm(row.targetPpm))}</td>
                <td className="numeric">{dim(formatPercent(row.achievementPercent))}</td>
                {periods.map((period) => <td className="numeric" key={period.key}>{dim(formatPpm(row.ppmByPeriod[period.key]))}</td>)}
              </tr>
            ))}
          </ReportTable>
        )}
      </Section>

      <Section title="Hierarchy summary" count={data.summaryRows.length}>
        {data.summaryRows.length === 0 ? <EmptySection /> : (
          <ReportTable caption="Weekly PPM by hierarchy" numericFrom={2} headers={["Item", "Level / Process", ...periods.map((period) => `${period.header} PPM`)]}>
            {data.summaryRows.map((row) => (
              <tr key={row.rowId}>
                <th scope="row" style={indentStyle(row.depth)}>{row.display}</th>
                <td>{row.level}{row.processType ? ` · ${row.processType}` : ""}</td>
                {periods.map((period) => <td className="numeric" key={period.key}>{dim(formatPpm(row.ppmByPeriod[period.key]))}</td>)}
              </tr>
            ))}
          </ReportTable>
        )}
      </Section>

      <Section title="Trend series" count={data.trendSeries.length}>
        {data.trendSeries.length === 0 ? <EmptySection /> : (
          <ReportTable caption="Weekly trend series" numericFrom={2} headers={["Series", "Group", ...periods.map((period) => `${period.header} PPM`)]}>
            {data.trendSeries.map((row) => (
              <tr key={row.id}>
                <th scope="row">{row.label}</th>
                <td>{dim(row.groupName ?? "-")}</td>
                {periods.map((period) => <td className="numeric" key={period.key}>{dim(formatPpm(row.ppmByPeriod[period.key]))}</td>)}
              </tr>
            ))}
          </ReportTable>
        )}
      </Section>

      <Section title="Top Defects" count={data.topDefects.length}>
        {data.topDefects.length === 0 ? <EmptySection /> : (
          <ReportTable caption="Weekly Top Defects" numericFrom={5} numericCols={[0]} headers={[
            "Rank", "Line Shift", "Process Type", "Process", "NG", ...periods.map((period) => `${period.header} PPM`),
          ]}>
            {data.topDefects.map((row, index) => (
              <tr key={`${row.rank}-${row.lineShift ?? "all"}-${row.processName}-${row.ngName}-${index}`}>
                <td className="numeric">{dim(row.rank)}</td>
                <td>{dim(row.lineShift ?? "-")}</td>
                <td>{row.processType}</td>
                <td>{row.processName}</td>
                <th scope="row">{row.ngName}</th>
                {periods.map((period) => <td className="numeric" key={period.key}>{dim(formatPpm(row.ppmByPeriod[period.key]))}</td>)}
              </tr>
            ))}
          </ReportTable>
        )}
      </Section>
    </div>
  );
}
