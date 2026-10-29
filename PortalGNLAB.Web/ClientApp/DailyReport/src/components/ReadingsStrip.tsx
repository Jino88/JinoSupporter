import type { DailyReportSnapshot } from "../contract";
import { directionClass, formatPpm, formatSignedPpm } from "../format";

export function ReadingsStrip({ snapshot }: { snapshot: DailyReportSnapshot }) {
  const weekDelta = snapshot.overallRecentPpm - snapshot.overallPreviousPpm;
  const dayDelta = snapshot.overallRecentDayPpm - snapshot.overallPreviousDayPpm;
  const readings = [
    { label: `LATEST WEEK · ${snapshot.recentWeekHeader}`, value: formatPpm(snapshot.overallRecentPpm), note: "overall PPM", direction: "" },
    { label: `BASELINE WEEK · ${snapshot.previousWeekHeader || "—"}`, value: formatPpm(snapshot.overallPreviousPpm), note: "last week with data", direction: "" },
    { label: "WEEK CHANGE", value: formatSignedPpm(weekDelta), note: "latest − baseline", direction: directionClass(weekDelta) },
    { label: `LATEST DAY · ${snapshot.recentDayHeader}`, value: formatPpm(snapshot.overallRecentDayPpm), note: "overall PPM", direction: "" },
    { label: `BASELINE DAY · ${snapshot.previousDayHeader || "—"}`, value: formatPpm(snapshot.overallPreviousDayPpm), note: "last day with data", direction: "" },
    { label: "DAY CHANGE", value: formatSignedPpm(dayDelta), note: "latest − baseline", direction: directionClass(dayDelta) },
  ];

  return (
    <section className="dr-readings" aria-label="Weekly and daily overall readings">
      {readings.map((reading) => (
        <div className="dr-reading" key={reading.label}>
          <span className="dr-reading-key">{reading.label}</span>
          <strong className={`dr-reading-value ${reading.direction}`.trim()}>{reading.value}</strong>
          <span className="dr-reading-note">{reading.note}</span>
        </div>
      ))}
    </section>
  );
}
