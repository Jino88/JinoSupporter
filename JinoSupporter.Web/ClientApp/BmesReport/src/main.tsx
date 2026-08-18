import { createRoot, type Root } from "react-dom/client";
import { App } from "./App";
import "./styles.css";

export interface MountOptions {
  reportUrl?: string;
}

const mountedRoots = new WeakMap<HTMLElement, Root>();

export function mount(element: HTMLElement, options: MountOptions = {}) {
  unmount(element);
  const reportUrl = options.reportUrl ?? element.dataset.reportUrl;
  if (!reportUrl) {
    throw new Error("BMES reportUrl is required.");
  }
  const root = createRoot(element);
  mountedRoots.set(element, root);
  root.render(<App reportUrl={reportUrl} />);
  return () => unmount(element);
}

export function unmount(element: HTMLElement) {
  const root = mountedRoots.get(element);
  if (!root) return;
  root.unmount();
  mountedRoots.delete(element);
}
