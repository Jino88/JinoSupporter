import { describe, expect, it } from "vitest";
import { fitView, panBy, textShown, toScreen, toWorld, toggleLayer, visibleItems, wheelFactor, zoomAt } from "./layout";

describe("process layout view maths", () => {
  it("fits the bounds centred in the box, Y flipped", () => {
    const v = fitView([0, 0, 100, 50], 232, 132, 16);
    expect(v.scale).toBeCloseTo(2);
    expect(toScreen(v, 0, 0)).toEqual([16, 116]);
    expect(toScreen(v, 100, 50)).toEqual([216, 16]);
  });

  it("converts screen and world back and forth", () => {
    const v = fitView([-10, -20, 30, 60], 800, 600);
    const [sx, sy] = toScreen(v, 7, -3);
    const [wx, wy] = toWorld(v, sx, sy);
    expect(wx).toBeCloseTo(7);
    expect(wy).toBeCloseTo(-3);
  });

  it("zooms around the cursor, keeping the world point under it", () => {
    const v = fitView([0, 0, 100, 100], 500, 500);
    const [wx, wy] = toWorld(v, 120, 340);
    const z = zoomAt(v, 120, 340, 2.5);
    expect(z.scale).toBeCloseTo(v.scale * 2.5);
    const [sx, sy] = toScreen(z, wx, wy);
    expect(sx).toBeCloseTo(120);
    expect(sy).toBeCloseTo(340);
  });

  it("limits the zoom", () => {
    expect(zoomAt({ scale: 1, tx: 0, ty: 0 }, 0, 0, 1e12).scale).toBe(1e6);
  });

  it("pans by the pixels dragged and zooms in on wheel up", () => {
    expect(panBy({ scale: 2, tx: 5, ty: 6 }, 3, -4)).toEqual({ scale: 2, tx: 8, ty: 2 });
    expect(wheelFactor(-100)).toBeGreaterThan(1);
    expect(wheelFactor(100)).toBeLessThan(1);
  });

  it("hides the items of hidden layers and toggles a layer", () => {
    const items = [{ l: 0 }, { l: 1 }, { l: 2 }];
    expect(visibleItems(items, new Set([1]))).toEqual([{ l: 0 }, { l: 2 }]);
    expect(visibleItems(items, new Set())).toBe(items);
    const hidden = toggleLayer(new Set(), 3);
    expect([...hidden]).toEqual([3]);
    expect(toggleLayer(hidden, 3).size).toBe(0);
  });

  it("skips text under 3px on screen", () => {
    expect(textShown({ scale: 1, tx: 0, ty: 0 }, 2.9)).toBe(false);
    expect(textShown({ scale: 1, tx: 0, ty: 0 }, 3)).toBe(true);
  });
});
