import { createRoot, type Root } from "react-dom/client";
import { App } from "./App";
import { fetchDailyReport } from "./api";
import "./styles.css";

export interface MountOptions {
  dataUrl?: string;
  refreshUrl?: string;
}

const mountedRoots = new WeakMap<HTMLElement, Root>();

export async function mount(element: HTMLElement, options: MountOptions = {}): Promise<void> {
  unmount(element);
  const dataUrl = options.dataUrl ?? element.dataset.dataUrl;
  const refreshUrl = options.refreshUrl ?? element.dataset.refreshUrl;
  if (!dataUrl || !refreshUrl) {
    throw new Error("DAILY REPORT dataUrl and refreshUrl are required.");
  }

  // Validate the live schema before creating a React root. Any contract/fetch failure
  // rejects this promise so the Razor host can switch to its retained legacy markup.
  const initialEnvelope = await fetchDailyReport(dataUrl);
  const root = createRoot(element);
  mountedRoots.set(element, root);
  root.render(<App initialEnvelope={initialEnvelope} dataUrl={dataUrl} refreshUrl={refreshUrl} />);
}

export function unmount(element: HTMLElement): void {
  const root = mountedRoots.get(element);
  if (!root) return;
  root.unmount();
  mountedRoots.delete(element);
}
