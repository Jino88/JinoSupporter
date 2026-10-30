import { defaultType, newId, newProcess } from "./logic";
import type { ModelSetup, PlanStage, Process, ProcessType, Segment, StageSource } from "./types";

const STAGES: [string, StageSource | null][] = [
  ["SPK Assy (Sub 12h)", "sub"],
  ["Lower Assy (Main 12h)", "main"],
  ["Function OK", "func"],
  ["Visual OK", "vis"],
  ["Send OQC", "oqc"],
  ["Export to KR (S-code)", null],
  ["Export to KR (Free)", null],
  ["Customer ship", null],
];

export const defaultStages = (): PlanStage[] => STAGES.map(([name, source]) => ({ id: newId(), name, source }));

export const emptySetup = (): ModelSetup => ({ processes: [newProcess("A")], materials: [], stages: defaultStages() });

/** Copy of a setup with new ids, for a new model. */
export function copySetup(setup: ModelSetup): ModelSetup {
  return {
    processes: setup.processes.map((p) => ({
      ...p,
      id: newId(),
      ngItems: p.ngItems.map((i) => ({ id: newId(), name: i.name })),
      materials: p.materials.map((m) => ({ ...m })),
    })),
    materials: [...setup.materials],
    stages: setup.stages.map((s) => ({ ...s, id: newId() })),
  };
}

const D3 = ["Glue overflow", "Glue shortage", "Misalign"];
type Row = [string, string, Segment, string[], ProcessType?, [string, number][]?];

// Processes, NG items and material links of the MSM-X316B PV2 "Daily NG Rate" workbook (2026-10).
const X316B: Row[] = [
  ["Bonding Enclosure", "Bôi keo tổ hợp spk", "A", ["NG"], undefined, [["UPPER", 1]]],
  ["AOI 1", "", "A", D3],
  ["SPK Assy", "Tổ hợp spk", "A", ["NG"], undefined, [["SPK 201507", 1]]],
  ["Fix Bonding SPK 1 (UV)", "Bôi keo cố định spk 1", "A", ["NG"]],
  ["AOI 2", "", "A", ["Glue overflow", "Glue shortage", "SPK reversed"]],
  ["Fix Bonding 2 (233)", "Bôi keo cố định 2", "A", ["NG"]],
  ["AOI 3", "", "A", D3],
  ["Fix Bonding 3", "Bôi keo cố định 3", "B", ["NG"]],
  ["AOI 4", "", "B", D3],
  ["F-PCB Adhesion", "Dán F PCB", "B", ["NG"], undefined, [["S-F-PCB", 1]]],
  ["Coating Cream Gluing", "Bôi keo trợ hàn", "B", ["NG"]],
  ["Spot Welding", "Hàn", "B", ["NG"]],
  ["Spot Vision", "", "B", ["Misalign", "Cold weld", "SUS peel"]],
  ["Bonding Lower Assy", "Bôi keo tổ hợp LOWER", "C", D3],
  ["AOI 5", "", "C", ["NG"]],
  ["QR Marking", "", "C", ["QR NG"]],
  ["Lower Assy", "Tổ hợp LOWER", "C", ["NG"], undefined, [["LOWER", 1]]],
  ["Bonding PT Assy", "Bôi keo tổ hợp PT", "C", D3],
  ["AOI 6", "", "C", ["NG"]],
  ["Plate Assy", "Tổ hợp PLATE", "C", ["NG"], undefined, [["S-PLATE", 1]]],
  ["Fix Bonding PT", "Keo cố định PT", "C", ["NG"]],
  ["Vision Inspection", "Kiểm tra keo", "C", D3],
  ["Glue Height Inspection", "High dimension — Bond", "D", ["NG"]],
  ["VS Adhesion", "Dán VS", "D", ["NG"], undefined, [["SCREEN-X516B-TOP-VS", 1]]],
  ["VS Press", "Ấn VS", "D", ["NG"]],
  ["CP Height Inspection", "High dimension — CP", "D", ["NG"]],
  ["FPCB Bonding", "Bôi keo FPCB", "D", ["NG"]],
  ["Air Leak", "", "E", ["NG"]],
  ["SPL Test VS", "Kiểm tra sóng VS", "E", ["Low wave", "High wave"]],
  ["BAKO", "", "E", ["RB+FFT", "RB", "FFT", "SPL", "F0"]],
  ["Hearing", "", "E", ["Noise", "Touch", "No sound"]],
  ["Gauss", "", "E", ["S", "N"], "VISUAL"],
  ["Visual", "", "E", ["Bond", "Particle", "Damage"], "VISUAL"],
  ["Packing", "", "F", ["QR NG"]],
];

/** Setup of the X316B PV2 workbooks, offered as a starting point for a new model. */
export function x316bSetup(): ModelSetup {
  const processes: Process[] = X316B.map(([name, localName, segment, defs, type = defaultType(segment), mats = []]) => ({
    id: newId(), name, localName, segment, type,
    ngItems: defs.map((d) => ({ id: newId(), name: d })),
    materials: mats.map(([material, qtyPerUnit]) => ({ material, qtyPerUnit })),
  }));
  return {
    processes,
    materials: ["UPPER", "LOWER", "SPK 201507", "S-PLATE", "S-F-PCB", "SCREEN-X516B-TOP-VS", "TRAY"],
    stages: defaultStages(),
  };
}
