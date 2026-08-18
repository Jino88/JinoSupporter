import type {
  FCostDataset,
  FCostMaterialRow,
  FCostView,
  ReportPeriod,
} from "../contract";
import { formatCount, formatNumber2, formatPercent, formatPpm, formatUsd, formatVnd } from "../format";
import { selectVisiblePeriods, type ViewerPreferences } from "../logic";
import { EmptySection, indentStyle, ReportTable, Section } from "../components/ReportTable";

export function FCostTab({
  dataset,
  view,
  preferences,
}: {
  dataset: FCostDataset;
  view: FCostView;
  preferences: ViewerPreferences;
}) {
  if (view.mode === "target-defect-rate") {
    return <TargetDefectRateView dataset={dataset} view={view} preferences={preferences} />;
  }

  const periods = selectVisiblePeriods(dataset.periods, preferences, view.allPeriods);
  return (
    <div className="tab-content">
      <p className="view-mode-note">{view.allPeriods ? "전체 기간" : "최근 기간"} · Regular F-COST · USD</p>
      <Section title="F-COST 합계">
        {periods.length === 0 ? <EmptySection /> : (
          <ReportTable caption="F-COST 기간별 합계" headers={["지표", ...periods.map((period) => period.header)]}>
            <tr><th scope="row">Input</th>{periods.map((period) => <td className="numeric" key={period.key}>{formatCount(dataset.totals.inputQtyByPeriod[period.key])}</td>)}</tr>
            <tr><th scope="row">F-COST (USD)</th>{periods.map((period) => <td className="numeric" key={period.key}>{formatUsd(dataset.totals.fcostUsdByPeriod[period.key])}</td>)}</tr>
            <tr><th scope="row">F-COST Rate</th>{periods.map((period) => <td className="numeric" key={period.key}>{formatPercent(dataset.totals.ratePercentByPeriod[period.key])}</td>)}</tr>
          </ReportTable>
        )}
      </Section>

      <Section title="Model / NG 추세" count={dataset.trendRows.length}>
        {dataset.trendRows.length === 0 ? <EmptySection /> : (
          <ReportTable caption="모델 및 NG 그룹 F-COST 추세" headers={[
            "Group", "Model", "NG Group",
            ...periods.flatMap((period) => [
              `${period.header} Input`, `${period.header} F-COST`, `${period.header} NG PPM`, `${period.header} Share`,
            ]),
          ]}>
            {dataset.trendRows.map((row) => (
              <tr key={row.rowId}>
                <td style={indentStyle(row.depth)}>{row.groupName}</td>
                <td>{row.modelName}</td>
                <th scope="row">{row.ngGroupKey}</th>
                {periods.flatMap((period) => [
                  <td className="numeric" key={`${period.key}-input`}>{formatCount(row.inputQtyByPeriod[period.key])}</td>,
                  <td className="numeric" key={`${period.key}-cost`}>{formatUsd(row.fcostUsdByPeriod[period.key])}</td>,
                  <td className="numeric" key={`${period.key}-ppm`}>{formatPpm(row.ngPpmByPeriod[period.key])}</td>,
                  <td className="numeric" key={`${period.key}-share`}>{formatPercent(row.fcostSharePercentByPeriod[period.key])}</td>,
                ])}
              </tr>
            ))}
          </ReportTable>
        )}
      </Section>

      <Section title="F-COST 계층" count={dataset.hierarchyRows.length}>
        {dataset.hierarchyRows.length === 0 ? <EmptySection /> : (
          <ReportTable caption="Mid/Sub 계층 F-COST" headers={[
            "항목", "Group", "Model", "NG Group", "Matched",
            ...periods.flatMap((period) => [`${period.header} Input`, `${period.header} F-COST`, `${period.header} Rate`]),
          ]}>
            {dataset.hierarchyRows.map((row) => (
              <tr key={row.rowId}>
                <th scope="row" style={indentStyle(row.depth)}>{row.display}</th>
                <td>{row.groupName}</td><td>{row.modelName}</td><td>{row.ngGroupKey}</td>
                <td className="numeric">{row.matchedMaterialCount}</td>
                {periods.flatMap((period) => [
                  <td className="numeric" key={`${period.key}-input`}>{formatCount(row.inputQtyByPeriod[period.key])}</td>,
                  <td className="numeric" key={`${period.key}-cost`}>{formatUsd(row.fcostUsdByPeriod[period.key])}</td>,
                  <td className="numeric" key={`${period.key}-rate`}>{formatPercent(row.sourceRatePercentByPeriod[period.key])}</td>,
                ])}
              </tr>
            ))}
          </ReportTable>
        )}
      </Section>

      <MaterialTable title="Material" rows={dataset.materials} periods={periods} />
      <MaterialTable title="Unmapped Material" rows={dataset.unmappedMaterials} periods={periods} />
      <RawBreakdown dataset={dataset} preferences={preferences} allPeriods={view.allPeriods} />
    </div>
  );
}

function MaterialTable({ title, rows, periods }: { title: string; rows: FCostMaterialRow[]; periods: ReportPeriod[] }) {
  return (
    <Section title={title} count={rows.length}>
      {rows.length === 0 ? <EmptySection /> : (
        <ReportTable caption={`${title} F-COST`} headers={[
          "Name", "Product Group", "Model No", "Material", "VERID",
          ...periods.flatMap((period) => [`${period.header} Input`, `${period.header} F-COST`, `${period.header} Rate`]),
        ]}>
          {rows.map((row, index) => (
            <tr key={`${row.displayName}-${row.material ?? "none"}-${row.verid ?? "none"}-${index}`}>
              <th scope="row">{row.displayName}</th>
              <td>{row.productGroup ?? "-"}</td><td>{row.modelNo ?? "-"}</td><td>{row.material ?? "-"}</td><td>{row.verid ?? "-"}</td>
              {periods.flatMap((period) => [
                <td className="numeric" key={`${period.key}-input`}>{formatCount(row.inputQtyByPeriod[period.key])}</td>,
                <td className="numeric" key={`${period.key}-cost`}>{formatUsd(row.fcostUsdByPeriod[period.key])}</td>,
                <td className="numeric" key={`${period.key}-rate`}>{formatPercent(row.sourceRatePercentByPeriod[period.key])}</td>,
              ])}
            </tr>
          ))}
        </ReportTable>
      )}
    </Section>
  );
}

function RawBreakdown({
  dataset,
  preferences,
  allPeriods,
}: {
  dataset: FCostDataset;
  preferences: ViewerPreferences;
  allPeriods: boolean;
}) {
  const raw = dataset.rawBreakdown;
  if (!raw) {
    return <Section title="Raw Material F-COST"><EmptySection message="Raw breakdown source를 사용할 수 없습니다." /></Section>;
  }
  const periods = selectVisiblePeriods(raw.periods, preferences, allPeriods);
  return (
    <Section title="Raw Material F-COST" count={raw.rows.length}>
      <dl className="source-meta">
        <div><dt>Source Table</dt><dd>{raw.sourceTable || "-"}</dd></div>
        <div><dt>Name Source</dt><dd>{raw.nameSource || "-"}</dd></div>
      </dl>
      {raw.warningMessage && <p className="inline-warning" role="status">{raw.warningMessage}</p>}
      {raw.exchangeRates.length > 0 && (
        <ReportTable caption="Raw breakdown 환율" headers={["Period", "기준일", "KRW/USD", "KRW/VND"]}>
          {raw.exchangeRates.map((rate) => (
            <tr key={rate.periodKey}>
              <th scope="row">{rate.periodKey}</th><td>{rate.standardDate}</td>
              <td className="numeric">{formatNumber2(rate.krwPerUsd)}</td>
              <td className="numeric">{formatNumber2(rate.krwPerVnd)}</td>
            </tr>
          ))}
        </ReportTable>
      )}
      {raw.rows.length === 0 ? <EmptySection /> : (
        <ReportTable caption="Raw material VND breakdown" headers={[
          "Group", "Model", "Material Code", "Material Name", "Total F-COST", "Source Rows",
          ...periods.flatMap((period) => [
            `${period.header} F-COST`, `${period.header} Eq. Qty`, `${period.header} Price`, `${period.header} Price VND`,
          ]),
        ]}>
          {raw.rows.map((row, index) => (
            <tr key={`${row.groupName}-${row.modelName}-${row.materialCode}-${index}`}>
              <td>{row.groupName}</td><td>{row.modelName}</td><td>{row.materialCode}</td><th scope="row">{row.materialName}</th>
              <td className="numeric">{formatVnd(row.totalFcostVnd)}</td><td className="numeric">{row.sourceRows.toLocaleString("en-US")}</td>
              {periods.flatMap((period) => {
                const price = row.priceByPeriod[period.key];
                const priceText = price?.unitPrice === null || price?.unitPrice === undefined
                  ? "-"
                  : `${formatNumber2(price.unitPrice)} ${price.currency ?? ""}${price.priceUnit ? `/${price.priceUnit}` : ""}${price.isMixed ? " (mixed)" : ""}`.trim();
                return [
                  <td className="numeric" key={`${period.key}-cost`}>{formatVnd(row.fcostVndByPeriod[period.key])}</td>,
                  <td className="numeric" key={`${period.key}-qty`}>{formatNumber2(row.equivalentQtyByPeriod[period.key])}</td>,
                  <td className="numeric" key={`${period.key}-price`}>{priceText}</td>,
                  <td className="numeric" key={`${period.key}-price-vnd`}>{formatVnd(price?.unitPriceVnd)}</td>,
                ];
              })}
            </tr>
          ))}
        </ReportTable>
      )}
    </Section>
  );
}

function TargetDefectRateView({
  dataset,
  view,
  preferences,
}: {
  dataset: FCostDataset;
  view: FCostView;
  preferences: ViewerPreferences;
}) {
  const target = dataset.targetDefectRate;
  const periods = selectVisiblePeriods(dataset.periods, preferences, view.allPeriods);
  return (
    <div className="tab-content">
      <p className="view-mode-note">{view.allPeriods ? "전체 기간" : "최근 기간"} · 목표 불량률</p>
      <Section title="목표 기준">
        <dl className="metric-meta">
          <div><dt>기본 Baseline</dt><dd>{formatPercent(target.defaultBaselineRatePercent)}</dd></div>
          {target.defaultTargetRatePercent.map((value, index) => (
            <div key={index}><dt>기본 Target {index + 1}</dt><dd>{formatPercent(value)}</dd></div>
          ))}
        </dl>
        {target.actionItems.length === 0 ? <EmptySection message="등록된 action item이 없습니다." /> : (
          <ol className="action-list">{target.actionItems.map((item, index) => <li key={`${index}-${item}`}>{item}</li>)}</ol>
        )}
      </Section>
      <Section title="목표 불량률" count={target.rows.length}>
        {target.rows.length === 0 ? <EmptySection /> : (
          <ReportTable caption="모델별 목표 불량률" headers={[
            "Model", "Target PPM", "Achievement", ...periods.map((period) => `${period.header} Actual PPM`),
          ]}>
            <tr className="total-row">
              <th scope="row">전체 Actual</th><td>-</td><td>-</td>
              {periods.map((period) => <td className="numeric" key={period.key}>{formatPpm(target.totalActualPpmByPeriod[period.key])}</td>)}
            </tr>
            {target.rows.map((row, index) => (
              <tr key={`${row.modelName}-${index}`}>
                <th scope="row">{row.modelName}</th>
                <td className="numeric">{formatPpm(row.targetPpm)}</td>
                <td className="numeric">{formatPercent(row.achievementPercent)}</td>
                {periods.map((period) => <td className="numeric" key={period.key}>{formatPpm(row.actualPpmByPeriod[period.key])}</td>)}
              </tr>
            ))}
          </ReportTable>
        )}
      </Section>
    </div>
  );
}
