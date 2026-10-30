import { useEffect, useMemo, useRef } from "react";
import { Chart, registerables, type ChartDataset, type Plugin } from "chart.js";
import type { SummaryRow } from "../../ngRateApi";
import { processTypeColor, summaryChartAxis, type ColumnSet } from "./logic";

// Ports of window.ngRateChart (By Model / By Section Summary: stacked process-type bars + TOTAL line)
// and window.ngRateGroupChart (By Group trend lines) from wwwroot/js/app.js.

Chart.register(...registerables);

const kilo = (v: string | number) => (Number(v) >= 1000 ? (Number(v) / 1000).toFixed(0) + "k" : v);

/** Draws each point's value above it; `lineOnly` limits it to line datasets. */
function valueLabels(id: string, lineOnly: boolean): Plugin {
  return {
    id,
    afterDatasetsDraw(chart) {
      const options = (chart.options.plugins as Record<string, { display?: boolean } | undefined>)[id];
      if (options?.display === false) return;
      const ctx = chart.ctx;
      ctx.save();
      ctx.font = "600 10px sans-serif";
      ctx.textAlign = "center";
      ctx.textBaseline = "bottom";
      ctx.fillStyle = "#111827";
      chart.data.datasets.forEach((dataset, datasetIndex) => {
        if (lineOnly && (dataset as { type?: string }).type !== "line") return;
        if (!chart.isDatasetVisible(datasetIndex)) return;
        chart.getDatasetMeta(datasetIndex).data.forEach((point, pointIndex) => {
          const value = dataset.data[pointIndex];
          if (value == null || !Number.isFinite(Number(value))) return;
          const y = Math.max(chart.chartArea.top + 12, point.y - 8);
          ctx.fillText(Number(value).toLocaleString(), point.x, y);
        });
      });
      ctx.restore();
    },
  };
}

const summaryLabels = valueLabels("ngRateSummaryValueLabels", true);
const groupLabels = valueLabels("ngRateValueLabels", false);

const separators = (labels: string[]) => new Set(labels.flatMap((l, i) => (l === "" ? [i] : [])));

export interface SummaryBar {
  label: string;
  data: (number | null)[];
  color: string;
}

/** Stacked bars per process type with the TOTAL line on the right axis. */
export function SummaryChart({ labels, bars, total }: {
  labels: string[];
  bars: SummaryBar[];
  total: { label: string; data: (number | null)[] } | null;
}) {
  const canvas = useRef<HTMLCanvasElement>(null);
  useEffect(() => {
    if (!canvas.current) return;
    const sep = separators(labels);
    const datasets: ChartDataset<"bar" | "line", (number | null)[]>[] = bars.map((b) => ({
      type: "bar" as const,
      label: b.label,
      data: b.data,
      backgroundColor: b.color,
      stack: "pt",
      order: 1,
    }));
    if (total) {
      datasets.push({
        type: "line" as const,
        label: total.label,
        data: total.data,
        borderColor: "#4f46e5",
        backgroundColor: "rgba(79,70,229,0.08)",
        borderWidth: 2,
        pointRadius: 3,
        pointBackgroundColor: "#4f46e5",
        pointBorderColor: "#4f46e5",
        tension: 0.4,
        fill: false,
        yAxisID: "y1",
        order: 0,
        spanGaps: false,
      });
    }
    const chart = new Chart(canvas.current, {
      type: "bar",
      data: { labels, datasets },
      plugins: [summaryLabels],
      options: {
        responsive: true,
        maintainAspectRatio: false,
        interaction: { mode: "index", intersect: false },
        plugins: {
          ...({ ngRateSummaryValueLabels: { display: labels.filter((l) => l !== "").length <= 14 } } as object),
          legend: { position: "bottom", labels: { boxWidth: 12, padding: 10, font: { size: 10 } } },
          tooltip: {
            filter: (item) => !sep.has(item.dataIndex),
            callbacks: { label: (c) => ` ${c.dataset.label}: ${c.parsed.y != null ? c.parsed.y.toLocaleString() : "-"}` },
          },
        },
        scales: {
          x: {
            stacked: true,
            ticks: { font: { size: 9 }, maxRotation: 0, callback: (_v, i) => (sep.has(i) ? "" : labels[i]) },
            grid: { color: (c) => (sep.has(c.index) ? "transparent" : "#f1f5f9") },
          },
          y: { stacked: true, ticks: { font: { size: 9 }, callback: kilo }, grid: { color: "#e2e8f0" } },
          y1: { position: "right", ticks: { font: { size: 9 }, callback: kilo }, grid: { drawOnChartArea: false } },
        },
      },
    });
    return () => chart.destroy();
  }, [labels, bars, total]);
  return <canvas ref={canvas} />;
}

/** Summary chart of one report over the visible columns (stacked process types + TOTAL line). */
export function ReportSummaryChart({ summary, cols, totalLabel = "TOTAL" }: { summary: SummaryRow[]; cols: ColumnSet; totalLabel?: string }) {
  const data = useMemo(() => {
    const axis = summaryChartAxis(cols);
    const totalRow = summary.find((r) => r.isTotal);
    return {
      labels: axis.labels,
      bars: summary.filter((r) => !r.isTotal).map((r, i) => ({ label: r.processType, data: axis.values(r.ppm), color: processTypeColor(r.processType, i) })),
      total: totalRow ? { label: totalLabel, data: axis.values(totalRow.ppm) } : null,
    };
  }, [summary, cols, totalLabel]);
  return <SummaryChart labels={data.labels} bars={data.bars} total={data.total} />;
}

const PALETTE = [
  "#4f46e5", "#2563eb", "#14b8a6", "#f97316", "#ef4444",
  "#8b5cf6", "#0ea5e9", "#f59e0b", "#10b981", "#ec4899",
  "#64748b", "#a855f7", "#22c55e", "#eab308", "#06b6d4",
];

/** One trend line per group. */
export function GroupChart({ labels, series }: { labels: string[]; series: { name: string; values: (number | null)[] }[] }) {
  const canvas = useRef<HTMLCanvasElement>(null);
  useEffect(() => {
    if (!canvas.current) return;
    const sep = separators(labels);
    const datasets = series.map((s, i) => {
      const color = PALETTE[i % PALETTE.length];
      return {
        label: s.name,
        data: s.values,
        borderColor: color,
        backgroundColor: color + "18",
        borderWidth: 2,
        pointBackgroundColor: color,
        pointBorderColor: color,
        pointRadius: 3,
        pointHoverRadius: 5,
        spanGaps: false,
        tension: 0.4,
      };
    });
    const chart = new Chart(canvas.current, {
      type: "line",
      data: { labels, datasets },
      plugins: [groupLabels],
      options: {
        responsive: true,
        maintainAspectRatio: false,
        interaction: { mode: "index", intersect: false },
        plugins: {
          ...({ ngRateValueLabels: { display: datasets.length <= 2 && labels.filter((l) => l !== "").length <= 14 } } as object),
          legend: {
            position: "bottom",
            labels: { boxWidth: 10, padding: 8, usePointStyle: true, font: { size: 10 }, color: "#111827" },
          },
          tooltip: {
            filter: (item) => !sep.has(item.dataIndex),
            callbacks: { label: (c) => ` ${c.dataset.label}: ${c.parsed.y != null ? c.parsed.y.toLocaleString() : "-"}` },
          },
        },
        scales: {
          x: {
            ticks: { font: { size: 10 }, maxRotation: 0, callback: (_v, i) => (sep.has(i) ? "" : labels[i]) },
            grid: { color: (c) => (sep.has(c.index) ? "#cbd5e1" : "#edf2f7") },
            border: { color: "#cbd5e1" },
          },
          y: {
            ticks: { font: { size: 10 }, callback: kilo },
            grid: { color: "#e5e7eb" },
            border: { display: false },
            beginAtZero: true,
          },
        },
      },
    });
    return () => chart.destroy();
  }, [labels, series]);
  return <canvas ref={canvas} />;
}
