import { useEffect, useState } from "react";
import { api, type PcDownload } from "../api";
import { formatPublished } from "../format";

/** Port of Components/Pages/PcDownloadPage.razor. */
export default function PcDownloadPage() {
  const [release, setRelease] = useState<PcDownload | null | undefined>(undefined);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    api.pcDownload().then(setRelease).catch((e: Error) => setError(e.message));
  }, []);

  return (
    <div className="pcd-page">
      <header className="page-head">
        <h1>PC Download</h1>
        <p className="muted">
          BMES NG Rate standalone — the desktop build of NG Rate / F-Cost / Worker Status / Routing / Material / LPA.
          Runs without the web server and updates itself.
        </p>
      </header>

      {error && <div className="alert">{error}</div>}
      {release === undefined && !error && <p className="muted">Loading…</p>}
      {release === null && (
        <div className="alert">
          No desktop build has been published yet. Run <code>BmesNgRateStandalone/tools/PublishStandaloneUpdate.ps1</code> to publish one.
        </div>
      )}
      {release && (
        <>
          <section className="card">
            <div className="pcd-name">BMES NG Rate</div>
            <div className="pcd-meta">
              <span className="pcd-ver">v{release.version}</span>
              {release.publishedAt && <> · <span>published {formatPublished(release.publishedAt)}</span></>}
              {" · "}<span>Windows 10/11 64-bit</span>
            </div>
            {release.notes.trim() && <div className="pcd-notes">{release.notes}</div>}
            <div className="pcd-assets">
              {release.setup && (
                <a className="btn primary" href={release.setup.url} download>
                  <span className="main">Download Installer</span>
                  <span className="sub">{release.setup.fileName} · {release.setup.sizeText}</span>
                </a>
              )}
              {release.package && (
                <a className="btn" href={release.package.url} download>
                  <span className="main">{release.setup ? "Portable zip" : "Download (zip)"}</span>
                  <span className="sub">{release.package.fileName} · {release.package.sizeText}</span>
                </a>
              )}
              {!release.setup && !release.package && (
                <div className="alert">
                  The manifest lists v{release.version} but no package file was found in <code>standalone-updates/</code>.
                </div>
              )}
            </div>
            {!release.setup && (
              <p className="muted small">
                Only the portable zip is published. To also produce a one-click installer, install{" "}
                <a href="https://jrsoftware.org/isdl.php" target="_blank" rel="noreferrer">Inno Setup 6</a> on the publishing
                machine and re-run <code>PublishStandaloneUpdate.ps1</code>.
              </p>
            )}
          </section>
          <section className="card">
            <h2>Install</h2>
            <ol>
              {release.setup ? (
                <li>Run the installer and follow the wizard. No admin rights are needed for a per-user install.</li>
              ) : (
                <li>
                  Unzip the package into a folder you can write to (e.g. <code>C:\BmesNgRate</code>), then run{" "}
                  <code>BmesNgRateStandalone.exe</code>.
                </li>
              )}
              <li>
                On first launch the app opens <strong>Setting</strong> and asks for the <strong>Working Folder</strong> — every NG
                Rate / F-Cost / Worker Status file is stored under it.
              </li>
              <li>Enter your <strong>BMES ID and password</strong> in the same Setting page, then restart the app.</li>
              <li>Later versions install themselves: the app checks this server on start and prompts to update.</li>
            </ol>
          </section>
        </>
      )}
    </div>
  );
}
