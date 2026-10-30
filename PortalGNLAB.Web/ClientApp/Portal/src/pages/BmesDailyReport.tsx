import { useEffect, useRef, useState } from "react";
import { bmesReportApi } from "../bmesReportApi";
import "../bmesReport.css";

interface ViewerModule {
  mount(element: HTMLElement, options: { dataUrl: string; refreshUrl: string }): Promise<void>;
  unmount(element: HTMLElement): void;
}

/** Load the viewer's stylesheet once; the bundle's classes are all dr-* so it cannot clash. */
function ensureStylesheet(href: string) {
  const id = "daily-report-viewer-css";
  const existing = document.getElementById(id) as HTMLLinkElement | null;
  if (existing) {
    if (existing.getAttribute("href") !== href) existing.setAttribute("href", href);
    return;
  }
  const link = document.createElement("link");
  link.id = id;
  link.rel = "stylesheet";
  link.href = href;
  document.head.appendChild(link);
}

/**
 * DAILY REPORT (port of BmesDailyReportPage.razor). Like the Blazor page it hosts the standalone
 * viewer built from ClientApp/DailyReport (wwwroot/daily-report) and points it at the same
 * /api/bmes/daily-report data seam; when that viewer is off or fails, it links to the classic page.
 */
export default function BmesDailyReport() {
  const host = useRef<HTMLDivElement>(null);
  const [failure, setFailure] = useState<{ text: string; legacy: boolean } | null>(null);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    let cancelled = false;
    let mounted: { module: ViewerModule; element: HTMLElement } | null = null;

    (async () => {
      try {
        const viewer = await bmesReportApi.daily.viewer();
        if (cancelled) return;
        if (!viewer.reactEnabled || !viewer.assetsExist) {
          setFailure({
            text: viewer.reactEnabled ? "The DAILY REPORT viewer files are not built on this server." : "The DAILY REPORT viewer is turned off on this server.",
            legacy: false,
          });
          return;
        }
        ensureStylesheet(viewer.cssUrl);
        const module = (await import(/* @vite-ignore */ viewer.scriptUrl)) as ViewerModule;
        const element = host.current;
        if (cancelled || !element) return;
        if (typeof module.mount !== "function") throw new Error("The viewer has no mount export.");
        await module.mount(element, { dataUrl: viewer.dataUrl, refreshUrl: viewer.refreshUrl });
        if (cancelled) {
          module.unmount(element);
          return;
        }
        mounted = { module, element };
      } catch (e) {
        if (cancelled) return;
        const message = e instanceof Error ? e.message : String(e);
        setFailure({ text: "The DAILY REPORT viewer could not be loaded: " + message, legacy: true });
        bmesReportApi.daily.viewerFailed(message).catch(() => undefined);
      } finally {
        if (!cancelled) setLoading(false);
      }
    })();

    return () => {
      cancelled = true;
      if (mounted) {
        try {
          mounted.module.unmount(mounted.element);
        } catch {
          // the page is going away anyway
        }
      }
    };
  }, []);

  return (
    <div className="br-page">
      {loading && !failure && <div className="muted small">Loading DAILY REPORT…</div>}
      {failure && (
        <div className="br-alert warn">
          <span>{failure.text}</span>{" "}
          <a href={failure.legacy ? "/bmes/daily-report?legacy=true" : "/bmes/daily-report"}>Open the classic DAILY REPORT page</a>
        </div>
      )}
      <div ref={host} className="dr-react-host" />
    </div>
  );
}
