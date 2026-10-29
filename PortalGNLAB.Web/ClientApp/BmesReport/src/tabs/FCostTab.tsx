import type {
  FCostDataset,
  FCostMaterialRow,
  FCostView,
  ReportPeriod,
} from "../contract";
import { formatCount, formatNumber2, formatPercent, formatPpm, formatUsd, formatVnd } from "../format";
import { selectVisiblePeriods, type ViewerPreferences } from "../logic";
import { dim, EmptySection, indentStyle, ReportTable, Section } from "../components/ReportTable";

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
      <p className="view-mode-note">{view.allPeriods ? "All periods" : "Recent periods"} · Regular F-COST · USD</p>
      <Section title="F-COST totals">
        {periods.length === 0 ? <EmptySection /> : (
          <ReportTable caption="F-COST totals by period" numericFrom={1} headers={["Metric", ...periods.map((period) => period.header)]}>
            <tr><th scope="row">Input</th>{periods.map((period) => <td className="numeric" key={period.key}>{dim(formatCount(dataset.totals.inputQtyByPeriod[period.key]))}</td>)}</tr>
            <tr><th scope="row">F-COST (USD)</th>{periods.map((period) => <td className="numeric" key={period.key}>{dim(formatUsd(dataset.totals.fcostUsdByPeriod[period.key]))}</td>)}</tr>
            <tr><th scope="row">F-COST Rate</th>{periods.map((period) => <td className="numeric" key={period.key}>{dim(formatPercent(dataset.totals.ratePercentByPeriod[period.key]))}</td>)}</tr>
          </ReportTable>
        )}
      </Section>

      <Section title="Model / NG trend" count={dataset.trendRows.length}>
        {dataset.trendRows.length === 0 ? <EmptySection /> : (
          <ReportTable caption="F-COST trend by model and NG group" numericFrom={3} headers={[
            "Group", "Model", "NG Group",
            ...periods.flatMap(() => ["Input", "F-COST", "NG PPM", "Share"]),
          ]} groups={{ lead: 3, items: periods.map((period) => ({ label: period.header, span: 4 })) }}>
            {dataset.trendRows.map((row) => (
              <tr key={row.rowId}>
                <td style={indentStyle(row.depth)}>{row.groupName}</td>
                <td>{row.modelName}</td>
                <th scope="row">{row.ngGroupKey}</th>
                {periods.flatMap((period) => [
                  <td className="numeric" key={`${period.key}-input`}>{dim(formatCount(row.inputQtyByPeriod[period.key]))}</td>,
                  <td className="numeric" key={`${period.key}-cost`}>{dim(formatUsd(row.fcostUsdByPeriod[period.key]))}</td>,
                  <td className="numeric" key={`${period.key}-ppm`}>{dim(formatPpm(row.ngPpmByPeriod[period.key]))}</td>,
                  <td className="numeric" key={`${period.key}-share`}>{dim(formatPercent(row.fcostSharePercentByPeriod[period.key]))}</td>,
                ])}
              </tr>
            ))}
          </ReportTable>
        )}
      </Section>

      <Section title="F-COST hierarchy" count={dataset.hierarchyRows.length}>
        {dataset.hierarchyRows.length === 0 ? <EmptySection /> : (
          <ReportTable caption="F-COST by Mid/Sub hierarchy" numericFrom={4} headers={[
            "Item", "Group", "Model", "NG Group", "Matched",
            ...periods.flatMap(() => ["Input", "F-COST", "Rate"]),
          ]} groups={{ lead: 5, items: periods.map((period) => ({ label: period.header, span: 3 })) }}>
            {dataset.hierarchyRows.map((row) => (
              <tr key={row.rowId}>
                <th scope="row" style={indentStyle(row.depth)}>{row.display}</th>
                <td>{row.groupName}</td><td>{row.modelName}</td><td>{row.ngGroupKey}</td>
                <td className="numeric">{row.matchedMaterialCount}</td>
                {periods.flatMap((period) => [
                  <td className="numeric" key={`${period.key}-input`}>{dim(formatCount(row.inputQtyByPeriod[period.key]))}</td>,
                  <td className="numeric" key={`${period.key}-cost`}>{dim(formatUsd(row.fcostUsdByPeriod[period.key]))}</td>,
                  <td className="numeric" key={`${period.key}-rate`}>{dim(formatPercent(row.sourceRatePercentByPeriod[period.key]))}</td>,
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
        <ReportTable caption={`${title} F-COST`} numericFrom={5} headers={[
          "Name", "Product Group", "Model No", "Material", "VERID",
          ...periods.flatMap(() => ["Input", "F-COST", "Rate"]),
        ]} groups={{ lead: 5, items: periods.map((period) => ({ label: period.header, span: 3 })) }}>
          {rows.map((row, index) => (
            <tr key={`${row.displayName}-${row.material ?? "none"}-${row.verid ?? "none"}-${index}`}>
              <th scope="row">{row.displayName}</th>
              <td>{dim(row.productGroup ?? "-")}</td><td>{dim(row.modelNo ?? "-")}</td><td>{dim(row.material ?? "-")}</td><td>{dim(row.verid ?? "-")}</td>
              {periods.flatMap((period) => [
                <td className="numeric" key={`${period.key}-input`}>{dim(formatCount(row.inputQtyByPeriod[period.key]))}</td>,
                <td className="numeric" key={`${period.key}-cost`}>{dim(formatUsd(row.fcostUsdByPeriod[period.key]))}</td>,
                <td className="numeric" key={`${period.key}-rate`}>{dim(formatPercent(row.sourceRatePercentByPeriod[period.key]))}</td>,
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
    return <Section title="Raw Material F-COST"><EmptySection message="Raw breakdown source is unavailable." /></Section>;
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
        <ReportTable caption="Raw breakdown exchange rates" numericFrom={2} headers={["Period", "Standard date", "KRW/USD", "KRW/VND"]}>
          {raw.exchangeRates.map((rate) => (
            <tr key={rate.periodKey}>
              <th scope="row">{rate.periodKey}</th><td>{rate.standardDate}</td>
              <td className="numeric">{dim(formatNumber2(rate.krwPerUsd))}</td>
              <td className="numeric">{dim(formatNumber2(rate.krwPerVnd))}</td>
            </tr>
          ))}
        </ReportTable>
      )}
      {raw.rows.length === 0 ? <EmptySection /> : (
        <ReportTable caption="Raw material VND breakdown" numericFrom={4} headers={[
          "Group", "Model", "Material Code", "Material Name", "Total F-COST", "Source Rows",
          ...periods.flatMap(() => ["F-COST", "Eq. Qty", "Price", "Price VND"]),
        ]} groups={{ lead: 6, items: periods.map((period) => ({ label: period.header, span: 4 })) }}>
          {raw.rows.map((row, index) => (
            <tr key={`${row.groupName}-${row.modelName}-${row.materialCode}-${index}`}>
              <td>{row.groupName}</td><td>{row.modelName}</td><td>{row.materialCode}</td><th scope="row">{row.materialName}</th>
              <td className="numeric">{dim(formatVnd(row.totalFcostVnd))}</td><td className="numeric">{row.sourceRows.toLocaleString("en-US")}</td>
              {periods.flatMap((period) => {
                const price = row.priceByPeriod[period.key];
                const priceText = price?.unitPrice === null || price?.unitPrice === undefined
                  ? "-"
                  : `${formatNumber2(price.unitPrice)} ${price.currency ?? ""}${price.priceUnit ? `/${price.priceUnit}` : ""}${price.isMixed ? " (mixed)" : ""}`.trim();
                return [
                  <td className="numeric" key={`${period.key}-cost`}>{dim(formatVnd(row.fcostVndByPeriod[period.key]))}</td>,
                  <td className="numeric" key={`${period.key}-qty`}>{dim(formatNumber2(row.equivalentQtyByPeriod[period.key]))}</td>,
                  <td className="numeric" key={`${period.key}-price`}>{priceText}</td>,
                  <td className="numeric" key={`${period.key}-price-vnd`}>{dim(formatVnd(price?.unitPriceVnd))}</td>,
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
      <p className="view-mode-note">{view.allPeriods ? "All periods" : "Recent periods"} · Target defect rate</p>
      <Section title="Target basis">
        <dl className="metric-meta">
          <div><dt>Default baseline</dt><dd>{formatPercent(target.defaultBaselineRatePercent)}</dd></div>
          {target.defaultTargetRatePercent.map((value, index) => (
            <div key={index}><dt>Default target {index + 1}</dt><dd>{formatPercent(value)}</dd></div>
          ))}
        </dl>
        {target.actionItems.length === 0 ? <EmptySection message="No action items registered." /> : (
          <ol className="action-list">{target.actionItems.map((item, index) => <li key={`${index}-${item}`}>{item}</li>)}</ol>
        )}
      </Section>
      <Section title="Target defect rate" count={target.rows.length}>
        {target.rows.length === 0 ? <EmptySection /> : (
          <ReportTable caption="Target defect rate by model" numericFrom={1} headers={[
            "Model", "Target PPM", "Achievement", ...periods.map((period) => `${period.header} Actual PPM`),
          ]}>
            <tr className="total-row">
              <th scope="row">Total actual</th><td className="numeric">{dim("-")}</td><td className="numeric">{dim("-")}</td>
              {periods.map((period) => <td className="numeric" key={period.key}>{dim(formatPpm(target.totalActualPpmByPeriod[period.key]))}</td>)}
            </tr>
            {target.rows.map((row, index) => (
              <tr key={`${row.modelName}-${index}`}>
                <th scope="row">{row.modelName}</th>
                <td className="numeric">{dim(formatPpm(row.targetPpm))}</td>
                <td className="numeric">{dim(formatPercent(row.achievementPercent))}</td>
                {periods.map((period) => <td className="numeric" key={period.key}>{dim(formatPpm(row.actualPpmByPeriod[period.key]))}</td>)}
              </tr>
            ))}
          </ReportTable>
        )}
      </Section>
    </div>
  );
}
