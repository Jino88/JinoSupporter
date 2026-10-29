import { useEffect, useState } from "react";
import type { PortalUser } from "./api";
import { groupActive, ICON_PATHS, isActive, type RailEntry, type RailTile } from "./menu";
import { menuClick, navigate } from "./nav";

function Icon({ name }: { name: string }) {
  return (
    <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.7"
      strokeLinecap="round" strokeLinejoin="round" aria-hidden="true"
      dangerouslySetInnerHTML={{ __html: ICON_PATHS[name] ?? "" }} />
  );
}

function Tile({ tile, pathname, onPick }: { tile: RailTile; pathname: string; onPick?: () => void }) {
  return (
    <a className={`rail-tile${isActive(tile.item, pathname) ? " active" : ""}`} href={tile.item.href}
      title={tile.item.label} onClick={menuClick(tile.item, onPick)}>
      <span className="rail-ico"><Icon name={tile.icon} /></span>
      <span className="rail-lbl">{tile.label}</span>
    </a>
  );
}

function initials(name: string): string {
  const n = name.trim();
  if (!n) return "?";
  const parts = n.split(/[ ._-]+/).filter(Boolean);
  return (parts.length >= 2 ? parts[0][0] + parts[1][0] : n.slice(0, 2)).toUpperCase();
}

export default function Rail({ entries, pathname, user }: { entries: RailEntry[]; pathname: string; user: PortalUser }) {
  const [open, setOpen] = useState<string | null>(null);
  useEffect(() => setOpen(null), [pathname]);

  const flyout = entries.find((e) => e.kind === "group" && e.key === open);
  const close = () => setOpen(null);

  return (
    <aside className="sidebar">
      <div className="top-row">
        <a className="rail-brand" href="/app/" title="GN LAB Supporter"
          onClick={(e) => { e.preventDefault(); navigate("/app/"); }}>
          <span className="rail-brand-mark" />
          <span className="rail-brand-text">GN LAB</span>
        </a>
      </div>

      <div className="rail-nav">
        <nav className="rail-list">
          {entries.map((e, i) => {
            if (e.kind === "sep") return <div key={`sep${i}`} className="rail-sep" />;
            if (e.kind === "tile") return <Tile key={e.item.id} tile={e} pathname={pathname} />;
            return (
              <button key={e.key} type="button" title={e.title} aria-expanded={open === e.key}
                className={`rail-tile rail-group${groupActive(e, pathname) ? " active" : ""}${open === e.key ? " open" : ""}`}
                onClick={() => setOpen(open === e.key ? null : e.key)}>
                <span className="rail-ico"><Icon name={e.icon} /></span>
                <span className="rail-lbl">{e.label}</span>
              </button>
            );
          })}
        </nav>

        <div className="rail-user" title={`Signed in as ${user.name} (${user.role})`}>
          <span className="rail-avatar">{initials(user.name)}</span>
          <span className="rail-role">{user.role}</span>
          <form method="post" action="/auth/logout">
            <button type="submit" className="rail-signout" title="Sign out" aria-label="Sign out">
              <Icon name="signout" />
            </button>
          </form>
        </div>
      </div>

      {flyout && flyout.kind === "group" && (
        <>
          <div className="rail-backdrop" onClick={close} />
          <div className="rail-flyout">
            <div className="rail-flyout-head">
              <span>{flyout.label}</span>
              <button type="button" className="rail-flyout-close" aria-label="Close" onClick={close}>×</button>
            </div>
            {flyout.sections.map((s, i) => (
              <div key={i}>
                {s.title && <div className="rail-flyout-sub">{s.title}</div>}
                <div className="rail-grid">
                  {s.tiles.map((t) => <Tile key={t.item.id} tile={t} pathname={pathname} onPick={close} />)}
                </div>
              </div>
            ))}
          </div>
        </>
      )}
    </aside>
  );
}
