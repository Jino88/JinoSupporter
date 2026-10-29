import { useEffect, useRef, useState } from "react";
import { memoApi } from "../../teamMemoApi";
import { ago, headerDiffers, today, weekdays } from "./logic";
import { AiSummaryPanel, Avatar } from "./shared";
import Timeline from "./Timeline";
import {
  statuses,
  statusLabel,
  type Memo,
  type MemoView,
  type Translate,
} from "./types";
export default function MemoDetail({
  memo,
  view,
  t,
  changed,
  close,
  plan = false,
}: {
  memo: Memo;
  view: MemoView;
  t: Translate;
  changed: () => Promise<void>;
  close: () => void;
  plan?: boolean;
}) {
  const [draft, setDraft] = useState(memo);
  const [original, setOriginal] = useState(memo);
  const [stale, setStale] = useState(false);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const state = useRef({ draft, original });
  state.current = { draft, original };
  useEffect(() => {
    const { draft: current, original: base } = state.current;
    if (!headerDiffers(current, base)) {
      setDraft(memo);
      setOriginal(memo);
      setStale(false);
    } else setStale(headerDiffers(memo, base));
  }, [memo]);
  const dirty = headerDiffers(draft, original);
  const display = (u: string) =>
    view.users.find((x) => x.username.toLowerCase() === u.toLowerCase())
      ?.displayName ?? u;
  async function run(action: () => Promise<void>) {
    setBusy(true);
    setError("");
    try {
      await action();
      await changed();
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setBusy(false);
    }
  }
  function revert() {
    setDraft(memo);
    setOriginal(memo);
    setStale(false);
    setError("");
  }
  async function save() {
    await run(async () => {
      await memoApi.send("PUT", `/items/${memo.id}/header`, {
        original,
        edited: draft,
      });
      const fresh = await memoApi.memo(memo.id);
      setOriginal(fresh);
      setDraft(fresh);
      setStale(false);
    });
  }
  const canDelete =
    view.isAdmin ||
    [memo.owner, memo.createdBy].some(
      (u) => u.toLowerCase() === view.currentUser.toLowerCase(),
    );
  const participants =
    view.memos.find((m) => m.id === memo.id)?.participants ?? [];
  return (
    <section className="memo-detail">
      <div className="memo-bar">
        <button onClick={close}>Close</button>
        <span>
          {view.folders.find((f) => f.id === memo.folderId)?.path ?? "Unfiled"}
        </span>
        <small>
          Revision {memo.revision} · edited {ago(memo.updatedAt)} by{" "}
          {display(memo.updatedBy)}
        </small>
        <button
          disabled={busy}
          onClick={() =>
            void run(() =>
              memoApi.send("POST", `/items/${memo.id}/pin`, {
                pinned: !memo.isPinned,
              }),
            )
          }
        >
          {memo.isPinned ? "Unpin" : "Pin"}
        </button>
        {canDelete && (
          <button
            disabled={busy}
            onClick={() => {
              if (
                window.confirm(
                  `Delete '${memo.title}' with its history and comments? It will be deleted from WorkHub too.`,
                )
              )
                void run(async () => {
                  await memoApi.send("DELETE", `/items/${memo.id}`);
                  close();
                });
            }}
          >
            Delete memo
          </button>
        )}
      </div>
      <div className="memo-bar">
        {participants.slice(0, 6).map((p) => (
          <Avatar key={p} name={p} display={display(p)} />
        ))}
        <small>{participants.length} participants</small>
      </div>
      {error && <p role="alert">{error}</p>}
      <fieldset disabled={busy}>
        <label>
          Title
          <input
            className="memo-title"
            value={draft.title}
            onChange={(e) => setDraft({ ...draft, title: e.target.value })}
          />
        </label>
        {t(memo.title) !== memo.title && (
          <p className="memo-translation">{t(memo.title)}</p>
        )}
        <div className="memo-fields">
          <label>
            Type
            <select
              value={draft.kind}
              onChange={(e) => setDraft({ ...draft, kind: e.target.value })}
            >
              {["Task", "History", "Review"].map((k) => (
                <option key={k}>{k}</option>
              ))}
            </select>
          </label>
          <label>
            Status
            <select
              value={memo.status}
              onChange={(e) =>
                void run(() =>
                  memoApi.send("POST", `/items/${memo.id}/status`, {
                    status: e.target.value,
                  }),
                )
              }
            >
              {statuses.map((s) => (
                <option key={s} value={s}>
                  {statusLabel(s)}
                </option>
              ))}
            </select>
          </label>
          <label>
            Folder
            <select
              value={draft.folderId ?? ""}
              onChange={(e) =>
                setDraft({ ...draft, folderId: e.target.value || null })
              }
            >
              <option value="">Unfiled</option>
              {view.folders.map((f) => (
                <option key={f.id} value={f.id}>
                  {f.path}
                </option>
              ))}
            </select>
          </label>
          <label>
            Owner
            <select
              value={draft.owner}
              onChange={(e) => setDraft({ ...draft, owner: e.target.value })}
            >
              <option value="">None</option>
              {[
                ...new Set(
                  [...view.users.map((u) => u.username), draft.owner].filter(
                    Boolean,
                  ),
                ),
              ]
                .sort((a, b) => display(a).localeCompare(display(b)))
                .map((u) => (
                  <option key={u} value={u}>
                    {display(u)}
                  </option>
                ))}
            </select>
          </label>
          <label>
            Start
            <input
              type="date"
              value={draft.startDate ?? ""}
              onChange={(e) =>
                setDraft({ ...draft, startDate: e.target.value || null })
              }
            />
          </label>
          <label>
            Due
            <input
              type="date"
              value={draft.dueDate ?? ""}
              onChange={(e) =>
                setDraft({ ...draft, dueDate: e.target.value || null })
              }
            />
          </label>
        </div>
        {memo.dueDate &&
          memo.dueDate < today() &&
          !["Done", "Canceled"].includes(memo.status) && (
            <p className="memo-late">Overdue since {memo.dueDate}</p>
          )}
        <div className="memo-bar">
          <b>Repeat weekly</b>
          {["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"].map((day, i) => (
            <label key={day}>
              <input
                type="checkbox"
                checked={weekdays(draft.repeatWeekdays).includes(i)}
                onChange={() => {
                  const days = weekdays(draft.repeatWeekdays);
                  const next = days.includes(i)
                    ? days.filter((d) => d !== i)
                    : [...days, i].sort();
                  setDraft({
                    ...draft,
                    repeatWeekdays: next.join(","),
                    repeatUntil: next.length ? draft.repeatUntil : null,
                  });
                }}
              />
              {day}
            </label>
          ))}
          {draft.repeatWeekdays && (
            <label>
              Until
              <input
                type="date"
                value={draft.repeatUntil ?? ""}
                onChange={(e) =>
                  setDraft({ ...draft, repeatUntil: e.target.value || null })
                }
              />
            </label>
          )}
        </div>
        <label>
          Tags
          <input
            value={draft.tags}
            onChange={(e) => setDraft({ ...draft, tags: e.target.value })}
          />
        </label>
        <label>
          {plan ? "Notes" : "Summary"}
          <textarea
            value={draft.description}
            onChange={(e) =>
              setDraft({ ...draft, description: e.target.value })
            }
          />
        </label>
        {t(memo.description) !== memo.description && (
          <p className="memo-translation">{t(memo.description)}</p>
        )}
        <div className="memo-bar">
          <button disabled={!dirty} onClick={() => void save()}>
            Save
          </button>
          <button onClick={revert}>Revert</button>
          {dirty && <strong>Unsaved changes</strong>}
          {stale && (
            <strong role="alert">
              The header changed elsewhere. Revert to load that version.
            </strong>
          )}
          <small>
            Created by {display(memo.createdBy)} · {memo.createdAt.slice(0, 10)}
          </small>
        </div>
      </fieldset>
      {!plan && (
        <AiSummaryPanel
          summary={memo}
          memo={memo}
          view={view}
          t={t}
          changed={changed}
        />
      )}
      <Timeline
        key={memo.id}
        memo={memo}
        view={view}
        t={t}
        changed={changed}
        showHistory={!plan}
      />
    </section>
  );
}
