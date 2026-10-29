import type { MouseEvent } from "react";
import type { MenuItem } from "./api";

export function navigate(href: string) {
  window.history.pushState(null, "", href);
  window.dispatchEvent(new PopStateEvent("popstate"));
}

/** Link that stays inside the React app for /app pages and loads the Blazor page otherwise. */
export function MenuLink({ item, className }: { item: MenuItem; className?: string }) {
  function onClick(e: MouseEvent<HTMLAnchorElement>) {
    if (!item.react || e.ctrlKey || e.metaKey || e.shiftKey || e.button !== 0) return;
    e.preventDefault();
    navigate(item.href);
  }
  return (
    <a className={className} href={item.href} onClick={onClick}>
      {item.label}
    </a>
  );
}
