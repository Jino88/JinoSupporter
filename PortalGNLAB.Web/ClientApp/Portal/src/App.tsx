import { useEffect, useState, type ReactNode } from "react";
import { api, type MenuItem, type PortalUser } from "./api";
import { appPath, buildRail, isChromeHidden } from "./menu";
import Rail from "./Rail";
import Home from "./pages/Home";
import PcDownloadPage from "./pages/PcDownload";
import AdminUsersPage from "./pages/AdminUsers";
import AdminAiUsages from "./pages/AdminAiUsages";
import AdminDbQuery from "./pages/AdminDbQuery";
import TestExcelConverter from "./pages/TestExcelConverter";
import Translate from "./pages/Translate";
import ReviewReport from "./pages/ReviewReport";
import GraphMaker from "./pages/GraphMaker";
import ProductionInput from "./pages/ProductionInput";
import QrSummary from "./pages/QrSummary";
import QrBakoData from "./pages/QrBakoData";
import QrAirleakData from "./pages/QrAirleakData";
import QrGaussData from "./pages/QrGaussData";
import QrGaussScatter from "./pages/QrGaussScatter";
import TeamMemo from "./pages/TeamMemo";
import TeamMemoBoard from "./pages/TeamMemoBoard";
import BmesRoutingTable from "./pages/BmesRoutingTable";
import BmesReasonTable from "./pages/BmesReasonTable";
import BmesWorkerStatus from "./pages/BmesWorkerStatus";
import BmesMakeModelGroup from "./pages/BmesMakeModelGroup";
import Setting from "./pages/Setting";
import BmesDailyReport from "./pages/BmesDailyReport";
import BmesReport from "./pages/BmesReport";
import BmesTest5 from "./pages/BmesTest5";
import BmesFCost from "./pages/BmesFCost";
import NgRate from "./pages/NgRate";
import NgRateAll from "./pages/NgRateAll";
import NgRateByGroup from "./pages/NgRateByGroup";

/** React pages by their path inside /app. A page added here also needs its menu id in
 *  PortalShellEndpoints.ReactRoutes so the menus (React and Blazor) link to it. */
const pages: Record<string, { title: string; render: (menu: MenuItem[]) => ReactNode }> = {
  "/": { title: "Home", render: (menu) => <Home menu={menu} /> },
  "/tools/pc-download": { title: "PC Download", render: () => <PcDownloadPage /> },
  "/admin/users": { title: "User Management", render: () => <AdminUsersPage /> },
  "/admin/ai-usages": { title: "AI Usages", render: () => <AdminAiUsages /> },
  "/admin/db-query": { title: "DB Query", render: () => <AdminDbQuery /> },
  "/admin/test-excel-converter": { title: "Test Excel Converter", render: () => <TestExcelConverter /> },
  "/tools/translate": { title: "Translate", render: () => <Translate /> },
  "/tools/review-report": { title: "Review Report", render: () => <ReviewReport /> },
  "/tools/graph-maker": { title: "Graph Maker", render: () => <GraphMaker /> },
  "/production/input": { title: "Production Input", render: () => <ProductionInput /> },
  "/qr/summary": { title: "QR SUMMARY", render: () => <QrSummary /> },
  "/qr/bako-data": { title: "QR BAKO DATA", render: () => <QrBakoData /> },
  "/qr/airleak-data": { title: "QR AIRLEAK DATA", render: () => <QrAirleakData /> },
  "/qr/gauss-data": { title: "QR GAUSS DATA", render: () => <QrGaussData /> },
  "/qr/gauss-scatter": { title: "GAUSS SCATTER", render: () => <QrGaussScatter /> },
  "/memo": { title: "Work Memo", render: () => <TeamMemo /> },
  "/memo/board": { title: "Memo Board", render: () => <TeamMemoBoard /> },
  "/memo/plan": { title: "Plan", render: () => <TeamMemo /> },
  "/memo/meetings": { title: "Plan", render: () => <TeamMemo /> },
  "/bmes/routing-table": { title: "Routing Table", render: () => <BmesRoutingTable /> },
  "/bmes/reason-table": { title: "Reason Table", render: () => <BmesReasonTable /> },
  "/bmes/worker-status": { title: "Worker Status", render: () => <BmesWorkerStatus /> },
  "/bmes/make-model-group": { title: "Model Group", render: () => <BmesMakeModelGroup /> },
  "/setting": { title: "Setting", render: () => <Setting /> },
  "/bmes/daily-report": { title: "DAILY REPORT", render: () => <BmesDailyReport /> },
  "/report/bmes": { title: "Report", render: () => <BmesReport /> },
  "/bmes/test5": { title: "BOM & Drawing", render: () => <BmesTest5 /> },
  "/bmes/f-cost": { title: "F-Cost", render: () => <BmesFCost /> },
  "/ng-rate": { title: "NG Rate", render: () => <NgRate /> },
  "/ng-rate-all": { title: "NG Rate", render: () => <NgRateAll /> },
  "/ng-rate-by-group": { title: "NG Rate - By Group", render: () => <NgRateByGroup /> },
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
      {!isChromeHidden(window.location.search) && <Rail entries={buildRail(menu)} pathname={pathname} user={user} />}
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
