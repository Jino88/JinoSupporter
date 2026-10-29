import { useState } from "react";
import { addDays, calendarItems, daysBetween, spanOf, today } from "./logic";
import type { Summary, Translate } from "./types";
function monthDate(value: string, delta: number) {
  const d = new Date(value + "T12:00:00");
  d.setMonth(d.getMonth() + delta, 1);
  return today(d);
}
function MonthBar({
  month,
  setMonth,
}: {
  month: string;
  setMonth: (m: string) => void;
}) {
  return (
    <div className="memo-bar">
      <button onClick={() => setMonth(monthDate(month, -1))}>
        Previous month
      </button>
      <b>
        {new Date(month + "T12:00:00").toLocaleDateString("en-US", {
          month: "long",
          year: "numeric",
        })}
      </b>
      <button onClick={() => setMonth(monthDate(month, 1))}>Next month</button>
      <button onClick={() => setMonth(today().slice(0, 7) + "-01")}>
        Today
      </button>
    </div>
  );
}
export function Gantt({
  memos,
  t,
  open,
}: {
  memos: Summary[];
  t: Translate;
  open: (id: string) => void;
}) {
  const [month, setMonth] = useState(today().slice(0, 7) + "-01");
  const end = addDays(monthDate(month, 1), -1);
  const days = daysBetween(month, monthDate(month, 1));
  const rows = memos
    .map((m) => ({ m, span: spanOf(m) }))
    .filter(
      (
        x,
      ): x is {
        m: Summary;
        span: {
          start: string;
          end: string;
        };
      } => !!x.span && x.span.start <= end && x.span.end >= month,
    )
    .sort((a, b) => a.span.start.localeCompare(b.span.start));
  return (
    <section className="memo-gantt">
      <MonthBar month={month} setMonth={setMonth} />
      <p>Start to due date. Open and overdue work extends to today.</p>
      <div className="memo-gantt-axis">
        <span>Memo</span>
        <div>
          {Array.from({ length: days }, (_, i) => (
            <small key={i}>{i + 1}</small>
          ))}
        </div>
      </div>
      {rows.map(({ m, span }) => {
        const start = span.start < month ? month : span.start;
        const stop = span.end > end ? end : span.end;
        return (
          <div className="memo-gantt-row" key={m.id}>
            <button onClick={() => open(m.id)}>{t(m.title)}</button>
            <div className="memo-gantt-track">
              <button
                title={`${m.title}: ${span.start} to ${span.end}`}
                style={{
                  left: `${(daysBetween(month, start) * 100) / days}%`,
                  width: `${((daysBetween(start, stop) + 1) * 100) / days}%`,
                }}
                onClick={() => open(m.id)}
                className={`memo-gantt-bar status-${m.status}`}
              >
                {start.slice(5)} – {stop.slice(5)}
              </button>
            </div>
          </div>
        );
      })}
      {!rows.length && <p>No scheduled memo in this month.</p>}
    </section>
  );
}
export function CalendarView({
  memos,
  t,
  open,
}: {
  memos: Summary[];
  t: Translate;
  open: (id: string) => void;
}) {
  const [month, setMonth] = useState(today().slice(0, 7) + "-01");
  const [plan, setPlan] = useState(true);
  const [work, setWork] = useState(true);
  const [due, setDue] = useState(true);
  const [expanded, setExpanded] = useState<string | null>(null);
  const first = addDays(month, -new Date(month + "T12:00:00").getDay());
  const last = addDays(first, 41);
  const items = calendarItems(memos, first, last, due).filter((i) =>
    i.plan ? plan : work,
  );
  return (
    <section className="memo-calendar">
      <MonthBar month={month} setMonth={setMonth} />
      <div className="memo-bar">
        <label>
          <input
            type="checkbox"
            checked={plan}
            onChange={(e) => setPlan(e.target.checked)}
          />
          Repeating plans
        </label>
        <label>
          <input
            type="checkbox"
            checked={work}
            onChange={(e) => setWork(e.target.checked)}
          />
          Work
        </label>
        <label>
          <input
            type="checkbox"
            checked={due}
            onChange={(e) => setDue(e.target.checked)}
          />
          Due dates and spans
        </label>
      </div>
      <div className="memo-calendar-grid">
        {["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"].map((d) => (
          <b key={d}>{d}</b>
        ))}
        {Array.from({ length: 42 }, (_, i) => addDays(first, i)).map((day) => {
          const dayItems = items.filter((it) => it.day === day);
          return (
            <div
              className={`memo-calendar-day ${day === today() ? "today" : ""} ${day.slice(0, 7) !== month.slice(0, 7) ? "outside" : ""}`}
              key={day}
            >
              <time>{day.slice(8)}</time>
              {(expanded === day ? dayItems : dayItems.slice(0, 4)).map(
                (it) => (
                  <button
                    key={`${it.memo.id}-${it.kind}`}
                    className={`status-${it.memo.status}`}
                    onClick={() => open(it.memo.id)}
                    title={it.memo.title}
                  >
                    {it.time}{" "}
                    {it.kind === "Repeat"
                      ? "↻ "
                      : it.kind === "Due"
                        ? "Due: "
                        : ""}
                    {t(it.memo.title)}
                  </button>
                ),
              )}
              {dayItems.length > 4 && (
                <button
                  onClick={() => setExpanded(expanded === day ? null : day)}
                >
                  {expanded === day
                    ? "Show less"
                    : `+${dayItems.length - 4} more`}
                </button>
              )}
            </div>
          );
        })}
      </div>
    </section>
  );
}
