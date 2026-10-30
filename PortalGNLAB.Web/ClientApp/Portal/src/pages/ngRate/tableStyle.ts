import { useEffect, type RefObject } from "react";
import { maxCellPositions, parseCellNumber } from "./logic";

// Port of window.ngRateTableStyle (wwwroot/js/app.js): marks the largest PPM of every column of each
// .pivot-table with .ppm-max-cell. Total / group-summary rows and label / separator cells do not count.

const SKIP_ROWS = ["total-row", "hier-group-row", "grp-overall-row", "grp-total-row", "wr-mid-row"];
const SKIP_CELLS = ["label-td", "sep-td", "toggle-cell", "row-hide-cell"];

function applyTable(table: HTMLTableElement): void {
  const rows = Array.from(table.tBodies).flatMap((body) => Array.from(body.rows));
  rows.forEach((row) => Array.from(row.cells).forEach((cell) => cell.classList.remove("ppm-max-cell")));

  const eligible = rows.filter((row) => !SKIP_ROWS.some((c) => row.classList.contains(c)));
  const values = eligible.map((row) =>
    Array.from(row.cells).map((cell) =>
      cell.tagName === "TD" && !SKIP_CELLS.some((c) => cell.classList.contains(c)) ? parseCellNumber(cell.textContent) : null,
    ),
  );
  for (const [r, c] of maxCellPositions(values)) eligible[r].cells[c].classList.add("ppm-max-cell");
}

/** Keeps the column-maximum highlight current while anything under `root` changes (rows expand, columns toggle). */
export function usePpmMaxHighlight(root: RefObject<HTMLElement | null>): void {
  useEffect(() => {
    const el = root.current;
    if (!el) return;
    let frame = 0;
    const schedule = () => {
      if (frame) return;
      frame = requestAnimationFrame(() => {
        frame = 0;
        el.querySelectorAll<HTMLTableElement>("table.pivot-table").forEach(applyTable);
      });
    };
    schedule();
    // Class changes made by applyTable are attributes, which are not observed, so this does not loop.
    const observer = new MutationObserver(schedule);
    observer.observe(el, { childList: true, subtree: true, characterData: true });
    return () => {
      observer.disconnect();
      cancelAnimationFrame(frame);
    };
  }, [root]);
}
