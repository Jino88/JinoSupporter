import { useCallback, useEffect, useLayoutEffect, useRef, useState } from "react";
import { qrApi, type QrPrefs } from "../../qrApi";
import { virtualWindow } from "./logic";

/** How long a date edit has to settle before it reloads (a native date input fires "change" per complete date typed through). */
export const DATE_SETTLE_MS = 900;

/** Loads the QR preferences (connection, remembered caps, model-code table) once. */
export function useQrPrefs() {
  const [prefs, setPrefs] = useState<QrPrefs | null>(null);
  const [error, setError] = useState("");
  useEffect(() => {
    qrApi.prefs().then(setPrefs, (e: Error) => setError(e.message));
  }, []);
  return { prefs, error };
}

/** Fixed-height row windowing for a scroll container: only the rows near the viewport are mounted. */
export function useVirtualRows(count: number, itemSize: number) {
  const ref = useRef<HTMLDivElement | null>(null);
  const [view, setView] = useState({ top: 0, height: 800 });

  const measure = useCallback(() => {
    const el = ref.current;
    if (el) setView((v) => (v.top === el.scrollTop && v.height === el.clientHeight ? v : { top: el.scrollTop, height: el.clientHeight }));
  }, []);

  useLayoutEffect(() => {
    const el = ref.current;
    if (!el) return;
    measure();
    const observer = new ResizeObserver(measure);
    observer.observe(el);
    return () => observer.disconnect();
  });

  return { ref, onScroll: measure, ...virtualWindow(count, itemSize, view.top, view.height) };
}

/** A timer that replaces itself: the last schedule() wins; cancelled on unmount. */
export function useSettleTimer() {
  const timer = useRef<ReturnType<typeof setTimeout> | null>(null);
  const cancel = useCallback(() => {
    if (timer.current) clearTimeout(timer.current);
    timer.current = null;
  }, []);
  useEffect(() => cancel, [cancel]);
  const schedule = useCallback(
    (action: () => void, ms: number) => {
      cancel();
      timer.current = setTimeout(action, ms);
    },
    [cancel],
  );
  return { schedule, cancel };
}
