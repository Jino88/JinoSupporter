import { useState } from "react";
import { memoApi } from "../../teamMemoApi";
import { ago, newestFirst, stamp, today } from "./logic";
import {
  AiSummaryPanel,
  appendFiles,
  Attachments,
  Avatar,
  FilePicker,
  pasteFiles,
} from "./shared";
import type { Comment, Entry, Memo, MemoView, Translate } from "./types";
interface Props {
  memo: Memo;
  view: MemoView;
  t: Translate;
  changed: () => Promise<void>;
  showHistory?: boolean;
}
export default function Timeline({
  memo,
  view,
  t,
  changed,
  showHistory = true,
}: Props) {
  const [collapsed, setCollapsed] = useState(
    new Set(
      memo.entries
        .filter(
          (e) =>
            e.doneDate &&
            !e.attachments.length &&
            !e.comments.length &&
            !e.body.trim(),
        )
        .map((e) => e.id),
    ),
  );
  const [form, setForm] = useState<Partial<Entry> | null>(null);
  const [files, setFiles] = useState<File[]>([]);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const [inline, setInline] = useState<{
    id: string;
    field: "title" | "body";
    value: string;
  } | null>(null);
  const own = (author: string) =>
    view.isAdmin || author.toLowerCase() === view.currentUser.toLowerCase();
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
  async function saveEntry(entry: Partial<Entry>, attachments: File[] = []) {
    const data = new FormData();
    data.set(
      "entry",
      JSON.stringify({
        id: "",
        entryDate: today(),
        title: "",
        body: "",
        ...entry,
      }),
    );
    attachments.forEach((f) => data.append("files", f));
    await memoApi.upload(`/items/${memo.id}/entries`, data);
  }
  function start(entry?: Entry) {
    setFiles([]);
    setForm(
      entry
        ? { ...entry }
        : {
            id: "",
            entryDate: today(),
            title: "",
            body: "",
            targetDate: null,
            doneDate: null,
          },
    );
    setInline(null);
  }
  const toggle = (id: string) =>
    setCollapsed((s) => {
      const n = new Set(s);
      if (!n.delete(id)) n.add(id);
      return n;
    });
  const editor = form && (
    <section
      className="memo-entry-form"
      onPaste={(e) =>
        pasteFiles(e, (incoming) => {
          try {
            setFiles(appendFiles(files, incoming));
          } catch (ex) {
            setError((ex as Error).message);
          }
        })
      }
    >
      <h3>{form.id ? "Edit entry" : "New entry"}</h3>
      <div className="memo-fields">
        {(["entryDate", "targetDate", "doneDate"] as const).map((key) => (
          <label key={key}>
            {key === "entryDate"
              ? "Date"
              : key === "targetDate"
                ? "Target"
                : "Done date"}
            <input
              type="date"
              value={form[key] ?? ""}
              onChange={(e) =>
                setForm({ ...form, [key]: e.target.value || null })
              }
            />
          </label>
        ))}
        <label>
          Title
          <input
            value={form.title ?? ""}
            onChange={(e) => setForm({ ...form, title: e.target.value })}
          />
        </label>
      </div>
      <label>
        Detail
        <textarea
          value={form.body ?? ""}
          onChange={(e) => setForm({ ...form, body: e.target.value })}
        />
      </label>
      <FilePicker files={files} setFiles={setFiles} error={setError} />
      <p>Paste screenshots into any field above. Up to 20 files, 20 MB each.</p>
      <button
        disabled={busy || !form.title?.trim()}
        onClick={() =>
          void run(async () => {
            await saveEntry(form, files);
            setForm(null);
            setFiles([]);
          })
        }
      >
        Save entry
      </button>
      <button onClick={() => setForm(null)}>Cancel</button>
    </section>
  );
  return (
    <div className="memo-timeline">
      {error && <p role="alert">{error}</p>}
      {showHistory && (
        <>
          <div className="memo-bar">
            <h2>History ({memo.entries.length})</h2>
            <button
              onClick={() =>
                setCollapsed(new Set(memo.entries.map((e) => e.id)))
              }
            >
              Collapse all
            </button>
            <button onClick={() => setCollapsed(new Set())}>Expand all</button>
            <button onClick={() => start()}>Add entry</button>
          </div>
          {form && !form.id && editor}
          {!memo.entries.length && !form && (
            <p>No history yet. Add the first entry.</p>
          )}
          {newestFirst(memo.entries).map((entry) =>
            form?.id === entry.id ? (
              <div key={entry.id}>{editor}</div>
            ) : (
              <article
                key={entry.id}
                className="memo-entry"
                style={{ marginLeft: Math.min(entry.depth, 4) * 22 }}
              >
                <div className="memo-bar">
                  <button
                    aria-label="Toggle entry"
                    onClick={() => toggle(entry.id)}
                  >
                    {collapsed.has(entry.id) ? "Expand" : "Fold"}
                  </button>
                  <time>{entry.entryDate.slice(5)}</time>
                  {inline?.id === entry.id && inline.field === "title" ? (
                    <input
                      autoFocus
                      value={inline.value}
                      onChange={(e) =>
                        setInline({ ...inline, value: e.target.value })
                      }
                      onKeyDown={(e) => {
                        if (e.key === "Escape") setInline(null);
                        if (e.key === "Enter" && inline.value.trim())
                          void run(async () => {
                            await saveEntry({ ...entry, title: inline.value });
                            setInline(null);
                          });
                      }}
                      onBlur={() => {
                        if (inline.value.trim())
                          void run(async () => {
                            await saveEntry({ ...entry, title: inline.value });
                            setInline(null);
                          });
                      }}
                    />
                  ) : (
                    <button
                      className="memo-entry-title"
                      title="Click to edit title"
                      onClick={() =>
                        setInline({
                          id: entry.id,
                          field: "title",
                          value: entry.title,
                        })
                      }
                    >
                      {t(entry.title)}
                    </button>
                  )}
                  <label
                    className={
                      !entry.doneDate &&
                      entry.targetDate &&
                      entry.targetDate < today()
                        ? "memo-late"
                        : ""
                    }
                  >
                    {entry.doneDate ? "Done" : "Target"}
                    <input
                      type="date"
                      value={entry.targetDate ?? ""}
                      onChange={(e) =>
                        void run(() =>
                          saveEntry({
                            ...entry,
                            targetDate: e.target.value || null,
                          }),
                        )
                      }
                    />
                  </label>
                  <button
                    disabled={busy}
                    onClick={() =>
                      void run(() =>
                        memoApi.send("PUT", `/items/${memo.id}/entry-done`, {
                          id: entry.id,
                          doneDate: entry.doneDate ? null : today(),
                        }),
                      )
                    }
                  >
                    {entry.doneDate ? "Reopen" : "Done"}
                  </button>
                  <button onClick={() => start(entry)}>Edit / attach</button>
                  {own(entry.author) && (
                    <button
                      disabled={busy}
                      onClick={() =>
                        void run(() =>
                          memoApi.send(
                            "POST",
                            `/items/${memo.id}/entry-delete`,
                            { id: entry.id },
                          ),
                        )
                      }
                    >
                      Delete
                    </button>
                  )}
                  <Avatar name={entry.author} />
                  <small>
                    {entry.comments.length} comments ·{" "}
                    {entry.attachments.length} files
                  </small>
                </div>
                {!collapsed.has(entry.id) && (
                  <>
                    {inline?.id === entry.id && inline.field === "body" ? (
                      <div>
                        <textarea
                          autoFocus
                          value={inline.value}
                          onChange={(e) =>
                            setInline({ ...inline, value: e.target.value })
                          }
                          onKeyDown={(e) => {
                            if (e.key === "Escape") setInline(null);
                            if (e.key === "Enter" && e.ctrlKey)
                              void run(async () => {
                                await saveEntry({
                                  ...entry,
                                  body: inline.value,
                                });
                                setInline(null);
                              });
                          }}
                        />
                        <button
                          onClick={() =>
                            void run(async () => {
                              await saveEntry({ ...entry, body: inline.value });
                              setInline(null);
                            })
                          }
                        >
                          Save detail
                        </button>
                        <button onClick={() => setInline(null)}>Cancel</button>
                      </div>
                    ) : (
                      <div
                        className="memo-body memo-editable"
                        onClick={() =>
                          setInline({
                            id: entry.id,
                            field: "body",
                            value: entry.body,
                          })
                        }
                      >
                        {t(entry.body) || "Add detail"}
                      </div>
                    )}
                    <Attachments
                      files={entry.attachments}
                      view={view}
                      memoId={memo.id}
                      changed={changed}
                    />
                    <AiSummaryPanel
                      summary={entry}
                      memo={memo}
                      entryId={entry.id}
                      view={view}
                      t={t}
                      changed={changed}
                    />
                    <Thread
                      memo={memo}
                      comments={entry.comments}
                      entryId={entry.id}
                      view={view}
                      t={t}
                      changed={changed}
                    />
                  </>
                )}
              </article>
            ),
          )}
        </>
      )}
      <h2>Discussion ({memo.comments.length})</h2>
      <Attachments
        files={memo.attachments}
        view={view}
        memoId={memo.id}
        changed={changed}
      />
      <Thread
        memo={memo}
        comments={memo.comments}
        entryId=""
        view={view}
        t={t}
        changed={changed}
      />
    </div>
  );
}
function Thread({
  memo,
  comments,
  entryId,
  view,
  t,
  changed,
}: Props & {
  comments: Comment[];
  entryId: string;
}) {
  const [all, setAll] = useState(false);
  const [draft, setDraft] = useState("");
  const [result, setResult] = useState(false);
  const [files, setFiles] = useState<File[]>([]);
  const [edit, setEdit] = useState<{
    id: number;
    body: string;
  } | null>(null);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const own = (author: string) =>
    view.isAdmin || author.toLowerCase() === view.currentUser.toLowerCase();
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
  async function post() {
    if (busy || (!draft.trim() && !files.length)) return;
    await run(async () => {
      const data = new FormData();
      data.set("entryId", entryId);
      data.set("body", draft.trim());
      data.set("isResult", String(result));
      files.forEach((f) => data.append("files", f));
      await memoApi.upload(`/items/${memo.id}/comments`, data);
      setDraft("");
      setFiles([]);
      setResult(false);
      setAll(true);
    });
  }
  return (
    <div className="memo-thread">
      {!all && comments.length > 3 && (
        <button onClick={() => setAll(true)}>
          Show {comments.length - 3} earlier comments
        </button>
      )}
      {(all ? comments : comments.slice(-3)).map((c) => (
        <article
          className={c.isResult ? "memo-comment result" : "memo-comment"}
          key={c.id}
        >
          <div className="memo-bar">
            <Avatar name={c.author} />
            <b>
              {view.users.find(
                (u) => u.username.toLowerCase() === c.author.toLowerCase(),
              )?.displayName ?? c.author}
            </b>
            {c.isResult && <strong>RESULT</strong>}
            <time title={c.createdAt}>
              {stamp(c.createdAt)} · {ago(c.updatedAt)}
            </time>
            {(own(c.author) || own(memo.owner)) && (
              <button
                disabled={busy}
                onClick={() =>
                  void run(() =>
                    memoApi.send(
                      "POST",
                      `/items/${memo.id}/comments/${c.id}/result`,
                      { isResult: !c.isResult },
                    ),
                  )
                }
              >
                {c.isResult ? "Unmark result" : "Mark as result"}
              </button>
            )}
            {own(c.author) && (
              <>
                <button onClick={() => setEdit({ id: c.id, body: c.body })}>
                  Edit
                </button>
                <button
                  disabled={busy}
                  onClick={() =>
                    void run(() =>
                      memoApi.send(
                        "POST",
                        `/items/${memo.id}/comments/${c.id}/delete`,
                        {},
                      ),
                    )
                  }
                >
                  Delete
                </button>
              </>
            )}
          </div>
          {edit?.id === c.id ? (
            <>
              <textarea
                value={edit.body}
                onChange={(e) => setEdit({ ...edit, body: e.target.value })}
              />
              <button
                disabled={busy}
                onClick={() =>
                  void run(async () => {
                    await memoApi.send(
                      "POST",
                      `/items/${memo.id}/comments/${c.id}/edit`,
                      { body: edit.body },
                    );
                    setEdit(null);
                  })
                }
              >
                Save
              </button>
              <button onClick={() => setEdit(null)}>Cancel</button>
            </>
          ) : (
            <div className="memo-body" title={c.body}>
              {t(c.body)}
            </div>
          )}
          <Attachments
            files={c.attachments}
            view={view}
            memoId={memo.id}
            changed={changed}
          />
        </article>
      ))}
      <fieldset
        disabled={busy}
        onPaste={(e) =>
          pasteFiles(e, (incoming) => {
            try {
              setFiles(appendFiles(files, incoming));
            } catch (ex) {
              setError((ex as Error).message);
            }
          })
        }
      >
        <textarea
          aria-label="Comment"
          placeholder="Write a comment. Ctrl+V pastes an image; Ctrl+Enter sends."
          value={draft}
          onChange={(e) => setDraft(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === "Enter" && e.ctrlKey) {
              e.preventDefault();
              void post();
            }
          }}
        />
        <FilePicker files={files} setFiles={setFiles} error={setError} />
        <label>
          <input
            type="checkbox"
            checked={result}
            onChange={(e) => setResult(e.target.checked)}
          />
          Post as result
        </label>
        <button
          disabled={!draft.trim() && !files.length}
          onClick={() => void post()}
        >
          Send
        </button>
      </fieldset>
      {error && <p role="alert">{error}</p>}
    </div>
  );
}
