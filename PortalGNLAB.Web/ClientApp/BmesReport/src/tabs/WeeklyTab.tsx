import type { WeeklyTabData } from "../contract";
import { formatPercent, formatPpm } from "../format";
import { selectVisiblePeriods, type ViewerPreferences } from "../logic";
import { EmptySection, indentStyle, ReportTable, Section } from "../components/ReportTable";

export function WeeklyTab({ data, preferences }: { data: WeeklyTabData; preferences: ViewerPreferences }) {
  const periods = selectVisiblePeriods(data.periods, preferences);
  return (
    <div className="tab-content">
      <Section title="주간 목표" count={data.targetRows.length}>
        {data.targetRows.length === 0 ? <EmptySection /> : (
          <ReportTable caption="주간 목표와 달성률" headers={[
            "항목", "구분", "기준 PPM", "목표 PPM", "달성률", ...periods.map((period) => `${period.header} PPM`),
          ]}>
            {data.targetRows.map((row) => (
              <tr key={row.rowId} className={row.rowKind === "group" ? "total-row" : ""}>
                <th scope="row" style={indentStyle(row.depth)}>{row.display}</th>
                <td>{row.rowKind}{row.isLineShift ? " · line shift" : ""}</td>
                <td className="numeric">{formatPpm(row.baselinePpm)}</td>
                <td className="numeric">{formatPpm(row.targetPpm)}</td>
                <td className="numeric">{formatPercent(row.achievementPercent)}</td>
                {periods.map((period) => <td className="numeric" key={period.key}>{formatPpm(row.ppmByPeriod[period.key])}</td>)}
              </tr>
            ))}
          </ReportTable>
        )}
      </Section>

      <Section title="계층 요약" count={data.summaryRows.length}>
        {data.summaryRows.length === 0 ? <EmptySection /> : (
          <ReportTable caption="Weekly 계층 PPM 요약" headers={["항목", "Level / Process", ...periods.map((period) => `${period.header} PPM`)]}>
            {data.summaryRows.map((row) => (
              <tr key={row.rowId}>
                <th scope="row" style={indentStyle(row.depth)}>{row.display}</th>
                <td>{row.level}{row.processType ? ` · ${row.processType}` : ""}</td>
                {periods.map((period) => <td className="numeric" key={period.key}>{formatPpm(row.ppmByPeriod[period.key])}</td>)}
              </tr>
            ))}
          </ReportTable>
        )}
      </Section>

      <Section title="추세 시리즈" count={data.trendSeries.length}>
        {data.trendSeries.length === 0 ? <EmptySection /> : (
          <ReportTable caption="Weekly 추세 시리즈" headers={["시리즈", "그룹", ...periods.map((period) => `${period.header} PPM`)]}>
            {data.trendSeries.map((row) => (
              <tr key={row.id}>
                <th scope="row">{row.label}</th>
                <td>{row.groupName ?? "-"}</td>
                {periods.map((period) => <td className="numeric" key={period.key}>{formatPpm(row.ppmByPeriod[period.key])}</td>)}
              </tr>
            ))}
          </ReportTable>
        )}
      </Section>

      <Section title="Top Defects" count={data.topDefects.length}>
        {data.topDefects.length === 0 ? <EmptySection /> : (
          <ReportTable caption="Weekly Top Defects" headers={[
            "순위", "Line Shift", "Process Type", "Process", "NG", ...periods.map((period) => `${period.header} PPM`),
          ]}>
            {data.topDefects.map((row, index) => (
              <tr key={`${row.rank}-${row.lineShift ?? "all"}-${row.processName}-${row.ngName}-${index}`}>
                <td className="numeric">{row.rank}</td>
                <td>{row.lineShift ?? "-"}</td>
                <td>{row.processType}</td>
                <td>{row.processName}</td>
                <th scope="row">{row.ngName}</th>
                {periods.map((period) => <td className="numeric" key={period.key}>{formatPpm(row.ppmByPeriod[period.key])}</td>)}
              </tr>
            ))}
          </ReportTable>
        )}
      </Section>
    </div>
  );
}
