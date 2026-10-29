import { useEffect, useRef, useState } from "react";
import { api, type AdminUser, type AdminUsersView } from "../api";
import { ADMIN_ROLE, isAllowed, menusByGroup, withPermission } from "../permissions";

const DEFAULT_ROLE = "Viewer";

/** Port of Components/Pages/UsersPage.razor (Admin only). */
export default function AdminUsersPage() {
  const [view, setView] = useState<AdminUsersView | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [savedFlash, setSavedFlash] = useState(false);
  const flashTimer = useRef<number | undefined>(undefined);

  const [newUser, setNewUser] = useState({ username: "", displayName: "", password: "", role: DEFAULT_ROLE });
  const [addError, setAddError] = useState<string | null>(null);

  const [editNameId, setEditNameId] = useState<number | null>(null);
  const [editNameValue, setEditNameValue] = useState("");

  const [resetTarget, setResetTarget] = useState<AdminUser | null>(null);
  const [resetPassword, setResetPassword] = useState("");
  const [resetError, setResetError] = useState<string | null>(null);

  const load = () =>
    api.admin
      .users()
      .then((v) => {
        setView(v);
        setError(null);
      })
      .catch((e: Error) => setError(e.message));

  useEffect(() => {
    load();
    return () => window.clearTimeout(flashTimer.current);
  }, []);

  /** Runs a change, shows its error, and reloads the list. */
  async function change(action: () => Promise<void>) {
    try {
      await action();
    } catch (e) {
      setError((e as Error).message);
    }
    await load();
  }

  async function togglePermission(role: string, menuId: string, allowed: boolean) {
    if (!view) return;
    setView({ ...view, permissions: withPermission(view.permissions, role, menuId, allowed) });
    try {
      await api.admin.setPermission(role, menuId, allowed);
      setSavedFlash(true);
      window.clearTimeout(flashTimer.current);
      flashTimer.current = window.setTimeout(() => setSavedFlash(false), 1500);
    } catch (e) {
      setError((e as Error).message);
      await load();
    }
  }

  async function addUser() {
    setAddError(null);
    if (!newUser.username.trim()) return setAddError("Username required.");
    if (!newUser.password.trim()) return setAddError("Password required.");
    try {
      await api.admin.addUser(newUser);
      setNewUser({ username: "", displayName: "", password: "", role: DEFAULT_ROLE });
      await load();
    } catch (e) {
      setAddError((e as Error).message);
    }
  }

  async function saveName(id: number) {
    setEditNameId(null);
    await change(() => api.admin.setName(id, editNameValue.trim()));
  }

  async function confirmReset() {
    setResetError(null);
    if (!resetTarget) return;
    if (!resetPassword.trim()) return setResetError("Password required.");
    try {
      await api.admin.setPassword(resetTarget.id, resetPassword);
      setResetTarget(null);
    } catch (e) {
      setResetError((e as Error).message);
    }
  }

  if (!view) return error ? <div className="alert">{error}</div> : <p className="muted">Loading…</p>;

  const isSelf = (u: AdminUser) => u.username.toLowerCase() === view.currentUser.toLowerCase();

  return (
    <div className="usr-page">
      <header className="page-head">
        <h1>User Management</h1>
      </header>
      {error && <div className="alert">{error}</div>}

      <section className="card flush">
        <h2>Users</h2>
        <table className="grid">
          <thead>
            <tr>
              <th>Username</th>
              <th>Name</th>
              <th>Role</th>
              <th>Created</th>
              <th style={{ width: 180 }}>Actions</th>
            </tr>
          </thead>
          <tbody>
            {view.users.map((u) => (
              <tr key={u.id}>
                <td>{u.username}</td>
                <td>
                  {editNameId === u.id ? (
                    <span className="row-edit">
                      <input
                        className="inp sm"
                        autoFocus
                        value={editNameValue}
                        onChange={(e) => setEditNameValue(e.target.value)}
                        onKeyDown={(e) => {
                          if (e.key === "Enter") saveName(u.id);
                          else if (e.key === "Escape") setEditNameId(null);
                        }}
                      />
                      <button className="btn sm primary" onClick={() => saveName(u.id)}>✓</button>
                      <button className="btn sm" onClick={() => setEditNameId(null)}>✕</button>
                    </span>
                  ) : (
                    <>
                      {u.displayName || "—"}{" "}
                      <button
                        className="link muted"
                        title="Edit name"
                        onClick={() => {
                          setEditNameId(u.id);
                          setEditNameValue(u.displayName);
                        }}
                      >
                        ✎
                      </button>
                    </>
                  )}
                </td>
                <td>
                  <select
                    className="inp sm"
                    value={u.role}
                    disabled={isSelf(u)}
                    onChange={(e) => change(() => api.admin.setRole(u.id, e.target.value))}
                  >
                    {view.roles.map((r) => (
                      <option key={r} value={r}>{r}</option>
                    ))}
                  </select>
                </td>
                <td className="muted">{u.createdAt}</td>
                <td>
                  <button
                    className="btn sm"
                    onClick={() => {
                      setResetTarget(u);
                      setResetPassword("");
                      setResetError(null);
                    }}
                  >
                    Reset PW
                  </button>{" "}
                  <button
                    className="btn sm danger"
                    disabled={isSelf(u)}
                    onClick={() => {
                      if (window.confirm(`Remove user ${u.username}?`)) change(() => api.admin.removeUser(u.id));
                    }}
                  >
                    Remove
                  </button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </section>

      <section className="card">
        <h2>Add User</h2>
        <div className="form-row">
          <label>
            Username
            <input className="inp sm" placeholder="username" value={newUser.username}
              onChange={(e) => setNewUser({ ...newUser, username: e.target.value })} />
          </label>
          <label>
            Name
            <input className="inp sm" placeholder="display name" value={newUser.displayName}
              onChange={(e) => setNewUser({ ...newUser, displayName: e.target.value })} />
          </label>
          <label>
            Password
            <input className="inp sm" type="password" placeholder="password" value={newUser.password}
              onChange={(e) => setNewUser({ ...newUser, password: e.target.value })} />
          </label>
          <label>
            Role
            <select className="inp sm" value={newUser.role} onChange={(e) => setNewUser({ ...newUser, role: e.target.value })}>
              {view.roles.map((r) => (
                <option key={r} value={r}>{r}</option>
              ))}
            </select>
          </label>
          <button className="btn sm primary" onClick={addUser}>Add</button>
        </div>
        {addError && <div className="alert small">{addError}</div>}
      </section>

      <section className="card flush">
        <h2 className="split">
          <span>Menu Permissions (by Role)</span>
          <span className="muted small">
            Changes save automatically on checkbox click. The Admin role always has full menu access.
            {savedFlash && <span className="tag ok">✓ Saved</span>}
          </span>
        </h2>
        <div className="scroll-x">
          <table className="grid perms">
            <thead>
              <tr>
                <th style={{ minWidth: 180 }}>Menu</th>
                {view.roles.map((r) => (
                  <th key={r}>{r}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {menusByGroup(view.menus).map(([group, menus]) => [
                <tr key={"g-" + group} className="group-row">
                  <td colSpan={view.roles.length + 1}>{group}</td>
                </tr>,
                ...menus.map((m) => (
                  <tr key={m.id}>
                    <td>
                      {m.label} <span className="muted small">({m.id})</span>
                    </td>
                    {view.roles.map((role) => (
                      <td key={role}>
                        <input
                          type="checkbox"
                          checked={isAllowed(view.permissions, role, m.id)}
                          disabled={role === ADMIN_ROLE}
                          onChange={(e) => togglePermission(role, m.id, e.target.checked)}
                        />
                      </td>
                    ))}
                  </tr>
                )),
              ])}
            </tbody>
          </table>
        </div>
      </section>

      {resetTarget && (
        <div className="modal-back" onClick={() => setResetTarget(null)}>
          <div className="modal" onClick={(e) => e.stopPropagation()}>
            <h2>Reset Password — {resetTarget.username}</h2>
            <input className="inp" type="password" placeholder="New password" autoFocus value={resetPassword}
              onChange={(e) => setResetPassword(e.target.value)}
              onKeyDown={(e) => e.key === "Enter" && confirmReset()} />
            <div className="modal-actions">
              <button className="btn sm primary" onClick={confirmReset}>Save</button>
              <button className="btn sm" onClick={() => setResetTarget(null)}>Cancel</button>
            </div>
            {resetError && <div className="alert small">{resetError}</div>}
          </div>
        </div>
      )}
    </div>
  );
}
