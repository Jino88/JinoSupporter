import { ApiError } from "../../api";
import { toolPath, toolSend } from "../../toolsApi";
import { upload } from "./api";

export type LayoutPath = { l: number; c: string | null; p: number[]; z?: 1 };
export type LayoutText = { l: number; c: string | null; x: number; y: number; h: number; r: number; t: string; a?: 0 | 1 | 2; v?: 0 | 1 | 2 | 3 };
export type Layout = {
  fileName: string; size: number; uploadedAt: string; uploadedBy: string;
  bounds: [number, number, number, number];
  layers: string[]; paths: LayoutPath[]; texts: LayoutText[];
};

/** Screen = world * scale + offset, with Y flipped (DXF Y goes up, the screen's goes down). */
export type View = { scale: number; tx: number; ty: number };

const url = (name: string) => `working/models/${encodeURIComponent(name)}/layout`;

/** The model's process layout, or null when none is saved (404). */
export async function loadLayout(name: string): Promise<Layout | null> {
  const response = await fetch(toolPath(url(name)), { credentials: "same-origin", headers: { Accept: "application/json" } });
  if (response.status === 404) return null;
  if (!response.ok) throw new ApiError(response.status, `Request failed (HTTP ${response.status})`);
  return (await response.json()) as Layout;
}

export function uploadLayout(name: string, file: File): Promise<Layout> {
  const form = new FormData();
  form.append("file", file, file.name);
  return upload<Layout>(url(name), form);
}

export const deleteLayout = (name: string) => toolSend<void>(url(name), undefined, "DELETE");

/** A view that shows the whole bounds centred in a width x height box, with a margin (px) around it. */
export function fitView(bounds: Layout["bounds"], width: number, height: number, margin = 16): View {
  const [x0, y0, x1, y1] = bounds;
  const w = Math.max(x1 - x0, 1e-9), h = Math.max(y1 - y0, 1e-9);
  const scale = Math.max(Math.min((width - 2 * margin) / w, (height - 2 * margin) / h), 1e-9);
  return { scale, tx: (width - w * scale) / 2 - x0 * scale, ty: (height + h * scale) / 2 + y0 * scale };
}

export const toScreen = (v: View, x: number, y: number): [number, number] => [x * v.scale + v.tx, -y * v.scale + v.ty];
export const toWorld = (v: View, sx: number, sy: number): [number, number] => [(sx - v.tx) / v.scale, -(sy - v.ty) / v.scale];

/** Zoom by factor around the screen point (sx, sy), which keeps the world point under it. */
export function zoomAt(v: View, sx: number, sy: number, factor: number, min = 1e-6, max = 1e6): View {
  const scale = Math.min(max, Math.max(min, v.scale * factor));
  const k = scale / v.scale;
  return { scale, tx: sx - (sx - v.tx) * k, ty: sy - (sy - v.ty) * k };
}

export const panBy = (v: View, dx: number, dy: number): View => ({ ...v, tx: v.tx + dx, ty: v.ty + dy });

/** Mouse wheel deltaY (negative = zoom in) as a zoom factor. */
export const wheelFactor = (deltaY: number) => Math.exp(-Math.max(-300, Math.min(300, deltaY)) * 0.0015);

/** Paths or texts whose layer is not hidden. */
export const visibleItems = <T extends { l: number }>(items: T[], hidden: ReadonlySet<number>): T[] =>
  hidden.size ? items.filter((i) => !hidden.has(i.l)) : items;

/** Toggled copy of the hidden-layer set. */
export function toggleLayer(hidden: ReadonlySet<number>, layer: number): Set<number> {
  const next = new Set(hidden);
  if (!next.delete(layer)) next.add(layer);
  return next;
}

/** Texts smaller than this on screen (px of height) are skipped. */
export const MIN_TEXT_PX = 3;
export const textShown = (v: View, h: number) => h * v.scale >= MIN_TEXT_PX;
