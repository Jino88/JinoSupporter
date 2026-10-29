import { useEffect, useState, type ReactNode } from "react";
import { api, type MenuItem, type PortalUser } from "./api";
import { appPath, groupMenu, isActive } from "./menu";
import { MenuLink, navigate } from "./nav";
import Home from "./pages/Home";
import PcDownloadPage from "./pages/PcDownload";

/** React pages by their path inside /app. A page added here also needs its menu id in
 *  PortalShellEndpoints.ReactRoutes so the menus (React and Blazor) link to it. */
const pages: Record<string, { title: string; render: (menu: MenuItem[]) => ReactNode }> = {
  "/": { title: "Home", render: (menu) => <Home menu={menu} /> },
  "/tools/pc-download": { title: "PC Download", render: () => <PcDownloadPage /> },
};

export default function App() {
  const [pathname, setPathname] = useState(window.location.pathname);
  const [user, setUser] = useState<PortalUser | null>(null);
  const [menu, setMenu] = useState<MenuItem[]>([]);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    const onPop = () => setPathname(window.location.pathname);
    window.addEventListener("popstate", onPop);
    return () => window.removeEventListener("popstate", onPop);
  }, []);

  useEffect(() => {
    Promise.all([api.me(), api.menu()])
      .then(([me, items]) => {
        setUser(me);
        setMenu(items);
      })
      .catch((e: Error) => setError(e.message));
  }, []);

  const page = pages[appPath(pathname)];
  useEffect(() => {
    document.title = `${page?.title ?? "Not found"} - GN LAB Portal`;
  }, [page]);

  if (error) return <div className="boot err">{error}</div>;
  if (!user) return <div className="boot muted">Loading…</div>;

  return (
    <div className="app">
      <aside className="side">
        <a className="side-logo" href="/app/" onClick={(e) => { e.preventDefault(); navigate("/app/"); }}>
          <span className="b1">GN LAB</span><span className="b2">PORTAL</span>
        </a>
        <nav className="side-nav">
          {groupMenu(menu).map((group) => (
            <div key={group.name}>
              <div className="side-group">{group.name}</div>
              {group.items.map((item) => (
                <MenuLink key={item.id} item={item} className={isActive(item, pathname) ? "on" : ""} />
              ))}
            </div>
          ))}
        </nav>
        <div className="side-foot">
          <div className="who">{user.name} <span className="muted">· {user.role}</span></div>
          <a className="muted" href="/">Classic UI</a>
          <form method="post" action="/auth/logout" className="inline">
            <button type="submit" className="link">Sign out</button>
          </form>
        </div>
      </aside>
      <main className="main">
        {page ? page.render(menu) : (
          <section className="card">
            <h1>Page not found</h1>
            <p className="muted">This page has not been moved to the new UI yet.</p>
          </section>
        )}
      </main>
    </div>
  );
}
