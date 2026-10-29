import type { MenuItem } from "../api";
import { MenuLink } from "../nav";
import { groupMenu } from "../menu";

export default function Home({ menu }: { menu: MenuItem[] }) {
  const ported = menu.filter((m) => m.react).length;
  return (
    <>
      <header className="page-head">
        <h1>GN LAB Portal</h1>
        <p className="muted">
          New UI (React). {ported} of {menu.length} menus are on the new UI; the others open the classic page.
        </p>
      </header>
      <div className="tiles">
        {groupMenu(menu).map((group) => (
          <section key={group.name} className="card">
            <h2>{group.name}</h2>
            <ul className="plain">
              {group.items.map((item) => (
                <li key={item.id}>
                  <MenuLink item={item} />
                  {item.react && <span className="tag">new</span>}
                </li>
              ))}
            </ul>
          </section>
        ))}
      </div>
    </>
  );
}
