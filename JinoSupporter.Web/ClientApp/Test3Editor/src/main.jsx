import React, { useMemo, useRef, useState } from "react";
import { createRoot } from "react-dom/client";
import {
  DndContext,
  DragOverlay,
  KeyboardSensor,
  PointerSensor,
  closestCenter,
  useDraggable,
  useDroppable,
  useSensor,
  useSensors
} from "@dnd-kit/core";
import "./style.css";

const roots = new WeakMap();
const SEP = "";
const dragProcessId = (processId) => `process${SEP}${processId}`;
const dragMaterialId = (materialId) => `material${SEP}${materialId}`;
const dragMergeId = (modelName, laneCode) => `merge${SEP}${modelName}${SEP}${laneCode}`;
const dropCellId = (modelName, laneCode, step) => `cell${SEP}${modelName}${SEP}${laneCode}${SEP}${step}`;
const dropPaletteId = (modelName) => `palette${SEP}${modelName}`;

function normalizeSnapshot(value) {
  const snapshot = value ?? {};
  return {
    modelNames: snapshot.modelNames ?? [],
    selectedModel: snapshot.selectedModel ?? "",
    sides: (snapshot.sides ?? []).map((side) => ({
      ...side,
      lanes: (side.lanes ?? []).map((lane) => ({ ...lane, processes: lane.processes ?? [] })),
      availableProcesses: side.availableProcesses ?? []
    })),
    materials: snapshot.materials ?? [],
    bomBusy: Boolean(snapshot.bomBusy),
    bomSource: snapshot.bomSource ?? "",
    bomError: snapshot.bomError ?? "",
    statusMessage: snapshot.statusMessage ?? "",
    statusError: Boolean(snapshot.statusError)
  };
}

function sideClass(sideLabel) {
  return sideLabel === "L" ? "is-left" : sideLabel === "R" ? "is-right" : "";
}

function sideTitle(side) {
  return side.sideLabel === "L" ? "LEFT" : side.sideLabel === "R" ? "RIGHT" : "MODEL";
}

function laneNumber(laneCode) {
  if (!laneCode || laneCode.toUpperCase() === "MAIN") return 0;
  const parsed = Number.parseInt(laneCode.slice(3), 10);
  return Number.isFinite(parsed) ? parsed : 999;
}

function orderedLanes(side) {
  return [...side.lanes].sort((left, right) => laneNumber(left.laneCode) - laneNumber(right.laneCode));
}

function nextSubLane(lanes) {
  const used = new Set(lanes.map((lane) => lane.laneCode.toUpperCase()));
  for (let index = 1; index <= 99; index += 1) {
    if (!used.has(`SUB${index}`)) return `SUB${index}`;
  }
  return "SUB99";
}

function materialLabel(material) {
  return (material.materialName || "").trim() || (material.materialCode || "").trim() || "자재";
}

function buildLayoutRequest(sides) {
  return {
    lanes: sides.flatMap((side) =>
      side.lanes.map((lane) => ({
        modelName: side.modelName,
        laneCode: lane.laneCode,
        mergeTargetProcessId: lane.mergeTargetProcessId ?? "",
        processIds: lane.processes.map((process) => process.id)
      }))
    )
  };
}

function cloneSides(sides) {
  return sides.map((side) => ({
    ...side,
    lanes: side.lanes.map((lane) => ({ ...lane, processes: [...lane.processes] })),
    availableProcesses: [...side.availableProcesses]
  }));
}

// 레인의 스텝별 누적 표기를 만든다. 저장하지 않고 항상 계산한다.
// 예: MTR1+MTR2+MTR3 +SUB1
function buildLaneSteps(lane, firstInputByProcess, mergeSourcesByProcess) {
  const carried = [];
  return lane.processes.map((process) => {
    const firstInputs = firstInputByProcess.get(process.id) ?? [];
    for (const material of firstInputs) carried.push(materialLabel(material));
    const merges = mergeSourcesByProcess.get(process.id) ?? [];
    return { process, firstInputs, merges, formula: [...carried] };
  });
}

function formulaText(step) {
  const base = step.formula.join("+");
  const merged = step.merges.map((laneCode) => ` +${laneCode}`).join("");
  return `${base}${merged}`.trim();
}

function Test3Editor({ dotnet, initialSnapshot }) {
  const initial = normalizeSnapshot(initialSnapshot);
  const [snapshot, setSnapshot] = useState(initial);
  const [sides, setSides] = useState(initial.sides);
  const [modelInput, setModelInput] = useState(initial.selectedModel);
  const [processSearch, setProcessSearch] = useState("");
  const [materialSearch, setMaterialSearch] = useState("");
  const [busy, setBusy] = useState(false);
  const [activeDrag, setActiveDrag] = useState(null);
  const [paletteTab, setPaletteTab] = useState("process");
  const [paletteOpen, setPaletteOpen] = useState(true);
  const [mergePick, setMergePick] = useState(null);
  const [localError, setLocalError] = useState("");
  const requestNumber = useRef(0);
  const extraLanes = useRef(new Map());
  const sensors = useSensors(
    useSensor(PointerSensor, { activationConstraint: { distance: 6 } }),
    useSensor(KeyboardSensor)
  );

  const firstInputByProcess = useMemo(() => {
    const result = new Map();
    for (const material of snapshot.materials) {
      if (!material.assignedProcessId) continue;
      const values = result.get(material.assignedProcessId) ?? [];
      values.push(material);
      result.set(material.assignedProcessId, values);
    }
    return result;
  }, [snapshot.materials]);

  const mergeSourcesByProcess = useMemo(() => {
    const result = new Map();
    for (const side of sides) {
      for (const lane of side.lanes) {
        if (lane.laneCode === "MAIN" || !lane.mergeTargetProcessId) continue;
        const values = result.get(lane.mergeTargetProcessId) ?? [];
        values.push(lane.laneCode);
        result.set(lane.mergeTargetProcessId, values);
      }
    }
    return result;
  }, [sides]);

  // 서버 스냅샷에는 공정이 없는 빈 SUB 레인이 존재하지 않으므로 화면에서 만든 레인은 다시 붙여 준다.
  const withExtraLanes = (nextSides) =>
    nextSides.map((side) => {
      const extra = extraLanes.current.get(side.modelName);
      if (!extra || extra.size === 0) return side;
      const existing = new Set(side.lanes.map((lane) => lane.laneCode));
      const added = [...extra]
        .filter((laneCode) => !existing.has(laneCode))
        .map((laneCode) => ({ laneCode, mergeTargetProcessId: "", mergeTargetLabel: "", processes: [] }));
      return added.length === 0 ? side : { ...side, lanes: [...side.lanes, ...added] };
    });

  const applySnapshot = (value) => {
    const next = normalizeSnapshot(value);
    setSnapshot(next);
    setSides(withExtraLanes(next.sides));
    setModelInput(next.selectedModel);
  };

  const invoke = async (method, ...args) => {
    const currentRequest = ++requestNumber.current;
    setBusy(true);
    setLocalError("");
    try {
      const value = await dotnet.invokeMethodAsync(method, ...args);
      if (currentRequest === requestNumber.current) applySnapshot(value);
      return value;
    } catch (error) {
      setSnapshot((current) => ({
        ...current,
        statusMessage: error?.message ?? String(error),
        statusError: true
      }));
      return null;
    } finally {
      if (currentRequest === requestNumber.current) setBusy(false);
    }
  };

  const loadModel = async (force = false) => {
    extraLanes.current.clear();
    await invoke("ReactSelectModelAsync", modelInput.trim(), force);
  };

  const changeModelInput = (value) => {
    setModelInput(value);
    const exactModel = snapshot.modelNames.find(
      (model) => model.toLocaleLowerCase() === value.trim().toLocaleLowerCase()
    );
    if (exactModel) {
      extraLanes.current.clear();
      void invoke("ReactSelectModelAsync", exactModel, false);
    }
  };

  const persistLayout = async (nextSides) => {
    setSides(nextSides);
    await invoke("ReactSaveLayoutAsync", buildLayoutRequest(nextSides));
  };

  const addSubLane = (modelName) => {
    setSides((current) =>
      current.map((side) => {
        if (side.modelName !== modelName) return side;
        const laneCode = nextSubLane(side.lanes);
        const extra = extraLanes.current.get(modelName) ?? new Set();
        extra.add(laneCode);
        extraLanes.current.set(modelName, extra);
        return {
          ...side,
          lanes: [...side.lanes, { laneCode, mergeTargetProcessId: "", mergeTargetLabel: "", processes: [] }]
        };
      })
    );
  };

  const removeEmptyLane = (modelName, laneCode) => {
    extraLanes.current.get(modelName)?.delete(laneCode);
    if (mergePick?.modelName === modelName && mergePick?.laneCode === laneCode) setMergePick(null);
    setSides((current) =>
      current.map((side) =>
        side.modelName === modelName
          ? { ...side, lanes: side.lanes.filter((lane) => lane.laneCode !== laneCode || lane.processes.length > 0) }
          : side
      )
    );
  };

  const updateMergeTarget = (modelName, laneCode, processId) => {
    const next = cloneSides(sides);
    const side = next.find((item) => item.modelName === modelName);
    const lane = side?.lanes.find((item) => item.laneCode === laneCode);
    if (!lane) return;
    const mainLane = side.lanes.find((item) => item.laneCode === "MAIN");
    const targetIndex = mainLane?.processes.findIndex((item) => item.id === processId) ?? -1;
    lane.mergeTargetProcessId = processId;
    lane.mergeTargetLabel =
      targetIndex >= 0 ? `MAIN #${targetIndex + 1} · ${mainLane.processes[targetIndex].processName}` : "";
    setMergePick(null);
    void persistLayout(next);
  };

  // 드롭한 행이 그대로 스텝이 된다. 배열에서 빼낸 뒤 같은 행 번호에 끼워 넣는다.
  const moveProcessToCell = (activeData, cellData) => {
    const next = cloneSides(sides);
    const side = next.find((item) => item.modelName === activeData.modelName);
    if (!side) return;
    let process = side.availableProcesses.find((item) => item.id === activeData.processId);
    side.availableProcesses = side.availableProcesses.filter((item) => item.id !== activeData.processId);
    for (const lane of side.lanes) {
      const match = lane.processes.find((item) => item.id === activeData.processId);
      if (match) process = match;
      lane.processes = lane.processes.filter((item) => item.id !== activeData.processId);
    }
    if (!process) return;

    const targetLane = side.lanes.find((lane) => lane.laneCode === cellData.laneCode);
    if (!targetLane) return;
    const insertIndex = Math.min(Math.max(cellData.step, 0), targetLane.processes.length);
    targetLane.processes.splice(insertIndex, 0, { ...process, laneCode: cellData.laneCode });
    void persistLayout(next);
  };

  const unassignProcess = (activeData) => {
    const next = cloneSides(sides);
    const side = next.find((item) => item.modelName === activeData.modelName);
    if (!side) return;
    let process = null;
    for (const lane of side.lanes) {
      const match = lane.processes.find((item) => item.id === activeData.processId);
      if (match) process = match;
      lane.processes = lane.processes.filter((item) => item.id !== activeData.processId);
    }
    if (!process) return;
    side.availableProcesses = [
      { ...process, laneCode: "", processNo: "", order: 0 },
      ...side.availableProcesses.filter((item) => item.id !== process.id)
    ];
    void persistLayout(next);
  };

  const assignMaterial = (materialId, processId, usageQty, usageUnit) =>
    invoke("ReactSaveMaterialAsync", {
      materialId,
      processId,
      usageQty: Number(usageQty) || 1,
      usageUnit: usageUnit || "PC"
    });

  const handleDragEnd = ({ active, over }) => {
    const activeData = active?.data?.current;
    const overData = over?.data?.current;
    setActiveDrag(null);
    if (!activeData || !overData) return;

    if (overData.kind === "palette") {
      if (activeData.kind !== "process" || activeData.modelName !== overData.modelName) return;
      if (!activeData.laneCode) return;
      unassignProcess(activeData);
      return;
    }

    if (overData.kind !== "cell") return;

    // 자재는 L/R 짝 자재가 서버에서 치환되므로 모델이 달라도 받는다.
    if (activeData.kind === "material") {
      if (!overData.processId) {
        setLocalError("자재는 공정이 배치된 셀에만 투입할 수 있습니다.");
        return;
      }
      void assignMaterial(activeData.materialId, overData.processId, activeData.usageQty, activeData.usageUnit);
      return;
    }

    if (activeData.modelName !== overData.modelName) {
      setLocalError("같은 보드 안에서만 배치할 수 있습니다.");
      return;
    }

    if (activeData.kind === "process") {
      moveProcessToCell(activeData, overData);
      return;
    }

    if (activeData.kind === "merge") {
      if (overData.laneCode !== "MAIN" || !overData.processId) {
        setLocalError("SUB 산출물은 MAIN의 공정 셀에만 합류시킬 수 있습니다.");
        return;
      }
      updateMergeTarget(activeData.modelName, activeData.laneCode, overData.processId);
    }
  };

  const handleCellPick = (side, lane, process) => {
    if (!mergePick || mergePick.modelName !== side.modelName) return;
    if (lane.laneCode !== "MAIN" || !process) {
      setLocalError("SUB 산출물은 MAIN의 공정 셀에만 합류시킬 수 있습니다.");
      return;
    }
    updateMergeTarget(side.modelName, mergePick.laneCode, process.id);
  };

  const materialQuery = materialSearch.trim().toLocaleLowerCase();
  const filteredMaterials = snapshot.materials.filter(
    (material) =>
      !materialQuery ||
      material.materialCode.toLocaleLowerCase().includes(materialQuery) ||
      material.materialName.toLocaleLowerCase().includes(materialQuery)
  );
  const processQuery = processSearch.trim().toLocaleLowerCase();
  const availableCount = sides.reduce((sum, side) => sum + side.availableProcesses.length, 0);

  return (
    <DndContext
      sensors={sensors}
      collisionDetection={closestCenter}
      onDragStart={({ active }) => {
        setLocalError("");
        setActiveDrag(active.data.current ?? null);
      }}
      onDragCancel={() => setActiveDrag(null)}
      onDragEnd={handleDragEnd}
    >
      <section className="t3r-shell" aria-busy={busy}>
        <header className="t3r-toolbar">
          <div className="t3r-model-search">
            <label htmlFor="t3r-model">모델</label>
            <input
              id="t3r-model"
              list="t3r-model-options"
              value={modelInput}
              placeholder="기준 모델 또는 L/R 모델 검색"
              onChange={(event) => changeModelInput(event.target.value)}
              onKeyDown={(event) => {
                if (event.key === "Enter") void loadModel(false);
              }}
            />
            <datalist id="t3r-model-options">
              {snapshot.modelNames.map((model) => (
                <option key={model} value={model} />
              ))}
            </datalist>
            <button type="button" className="t3r-button primary" onClick={() => void loadModel(false)} disabled={busy}>
              검색
            </button>
          </div>
          <div className="t3r-toolbar-actions">
            <button
              type="button"
              className="t3r-button"
              onClick={() => void invoke("ReactReloadBomAsync")}
              disabled={busy || sides.length === 0}
            >
              BOM 새로고침
            </button>
            <button
              type="button"
              className="t3r-button danger"
              onClick={() => {
                if (window.confirm("현재 L/R 공정 배치를 모두 초기화할까요?")) {
                  extraLanes.current.clear();
                  void invoke("ReactResetLayoutAsync");
                }
              }}
              disabled={busy || sides.length === 0}
            >
              공정 초기화
            </button>
          </div>
        </header>

        {(snapshot.statusMessage || snapshot.bomError || localError) && (
          <div className="t3r-status-row">
            {localError && <div className="t3r-status error">{localError}</div>}
            {snapshot.statusMessage && (
              <div className={`t3r-status ${snapshot.statusError ? "error" : "success"}`}>{snapshot.statusMessage}</div>
            )}
            {snapshot.bomError && <div className="t3r-status error">{snapshot.bomError}</div>}
          </div>
        )}

        {mergePick && (
          <div className="t3r-pick-banner">
            <strong>{mergePick.laneCode}</strong> 산출물을 합류시킬 <strong>MAIN 공정 셀</strong>을 클릭하세요.
            <button type="button" className="t3r-link-button" onClick={() => setMergePick(null)}>
              취소
            </button>
          </div>
        )}

        {sides.length === 0 ? (
          <div className="t3r-welcome">
            <strong>모델을 검색하면 레인 매트릭스가 열립니다.</strong>
            <span>기준 모델을 선택하면 L/R 두 보드를 한 화면에서 함께 편집할 수 있습니다.</span>
          </div>
        ) : (
          <div className={`t3r-workspace ${paletteOpen ? "" : "palette-collapsed"}`}>
            <div className="t3r-boards">
              {sides.map((side) => (
                <SideBoard
                  key={side.modelName}
                  side={side}
                  activeDrag={activeDrag}
                  mergePick={mergePick}
                  firstInputByProcess={firstInputByProcess}
                  mergeSourcesByProcess={mergeSourcesByProcess}
                  onAddSub={() => addSubLane(side.modelName)}
                  onRemoveLane={(laneCode) => removeEmptyLane(side.modelName, laneCode)}
                  onStartPick={(laneCode) => setMergePick({ modelName: side.modelName, laneCode })}
                  onClearMerge={(laneCode) => updateMergeTarget(side.modelName, laneCode, "")}
                  onCellPick={(lane, process) => handleCellPick(side, lane, process)}
                  onUnassign={(processId) => unassignProcess({ modelName: side.modelName, processId })}
                />
              ))}
            </div>

            <aside className={`t3r-palette ${paletteOpen ? "" : "is-collapsed"}`}>
              {paletteOpen ? (
                <>
                  <div className="t3r-palette-head">
                    <button
                      type="button"
                      className={`t3r-tab ${paletteTab === "process" ? "is-active" : ""}`}
                      onClick={() => setPaletteTab("process")}
                    >
                      미배치 공정 <span>{availableCount}</span>
                    </button>
                    <button
                      type="button"
                      className={`t3r-tab ${paletteTab === "material" ? "is-active" : ""}`}
                      onClick={() => setPaletteTab("material")}
                    >
                      BOM 자재 <span>{filteredMaterials.length}</span>
                    </button>
                    <button
                      type="button"
                      className="t3r-icon-button plain"
                      title="팔레트 접기"
                      onClick={() => setPaletteOpen(false)}
                    >
                      ›
                    </button>
                  </div>
                  <div className="t3r-palette-search">
                    {paletteTab === "process" ? (
                      <input
                        value={processSearch}
                        onChange={(event) => setProcessSearch(event.target.value)}
                        placeholder="공정 검색"
                      />
                    ) : (
                      <input
                        value={materialSearch}
                        onChange={(event) => setMaterialSearch(event.target.value)}
                        placeholder="BOM 자재 검색"
                      />
                    )}
                  </div>
                  <div className="t3r-palette-body">
                    {paletteTab === "process" ? (
                      sides.map((side) => <ProcessPalette key={side.modelName} side={side} query={processQuery} />)
                    ) : (
                      <MaterialPalette
                        materials={filteredMaterials}
                        onSave={(material, usageQty, usageUnit) =>
                          void assignMaterial(material.id, material.assignedProcessId, usageQty, usageUnit)
                        }
                        onClear={(materialId) => void invoke("ReactClearMaterialAsync", materialId)}
                      />
                    )}
                    {paletteTab === "material" && snapshot.bomSource && (
                      <div className="t3r-palette-note">{snapshot.bomSource}</div>
                    )}
                  </div>
                </>
              ) : (
                <button type="button" className="t3r-palette-reopen" onClick={() => setPaletteOpen(true)}>
                  ‹ 팔레트
                </button>
              )}
            </aside>
          </div>
        )}
        {busy && (
          <div className="t3r-busy">
            <span />
            저장 중…
          </div>
        )}
      </section>

      <DragOverlay dropAnimation={null}>{activeDrag ? <DragPreview data={activeDrag} /> : null}</DragOverlay>
    </DndContext>
  );
}

function SideBoard({
  side,
  activeDrag,
  mergePick,
  firstInputByProcess,
  mergeSourcesByProcess,
  onAddSub,
  onRemoveLane,
  onStartPick,
  onClearMerge,
  onCellPick,
  onUnassign
}) {
  const lanes = orderedLanes(side);
  const stepCount = lanes.reduce((max, lane) => Math.max(max, lane.processes.length), 0);
  const rowCount = stepCount + 1;
  const laneSteps = lanes.map((lane) => buildLaneSteps(lane, firstInputByProcess, mergeSourcesByProcess));
  const mirrored = Boolean(side.mirrorModelName);
  const gridStyle = { gridTemplateColumns: `34px repeat(${Math.max(lanes.length, 1)}, minmax(178px, 1fr))` };
  const isPicking = mergePick?.modelName === side.modelName;

  return (
    <article className={`t3r-board ${sideClass(side.sideLabel)}`}>
      <header className="t3r-board-head">
        <span className="t3r-side-badge">{mirrored ? `${side.sideLabel}+${side.mirrorSideLabel}` : side.sideLabel || "-"}</span>
        <span className="t3r-board-title">{mirrored ? "L/R 동시 편집" : sideTitle(side)}</span>
        <strong title={mirrored ? `${side.modelName} · ${side.mirrorModelName}` : side.modelName}>
          {side.modelName}
          {mirrored && <em className="t3r-mirror-name"> · {side.mirrorModelName}</em>}
        </strong>
        {mirrored && <span className="t3r-mirror-note">한쪽만 배치하면 반대쪽에 같이 저장됩니다</span>}
        <button type="button" className="t3r-link-button" onClick={onAddSub}>
          + SUB 레인
        </button>
      </header>

      <div className="t3r-matrix-scroll">
        <div className="t3r-matrix" style={gridStyle}>
          <div className="t3r-matrix-corner">스텝</div>
          {lanes.map((lane) => (
            <LaneHeader
              key={lane.laneCode}
              side={side}
              lane={lane}
              isPicking={isPicking && mergePick.laneCode === lane.laneCode}
              onRemove={() => onRemoveLane(lane.laneCode)}
              onStartPick={() => onStartPick(lane.laneCode)}
              onClearMerge={() => onClearMerge(lane.laneCode)}
            />
          ))}

          {Array.from({ length: rowCount }, (unused, rowIndex) => (
            <React.Fragment key={`row-${rowIndex}`}>
              <div className={`t3r-step-gutter ${rowIndex === stepCount ? "is-tail" : ""}`}>
                {rowIndex === stepCount ? "+" : rowIndex + 1}
              </div>
              {lanes.map((lane, laneIndex) => (
                <MatrixCell
                  key={`${lane.laneCode}-${rowIndex}`}
                  side={side}
                  lane={lane}
                  step={rowIndex}
                  cellStep={laneSteps[laneIndex][rowIndex] ?? null}
                  activeDrag={activeDrag}
                  isPicking={isPicking}
                  onPick={() => onCellPick(lane, laneSteps[laneIndex][rowIndex]?.process ?? null)}
                  onUnassign={onUnassign}
                />
              ))}
            </React.Fragment>
          ))}
        </div>
      </div>
    </article>
  );
}

function LaneHeader({ side, lane, isPicking, onRemove, onStartPick, onClearMerge }) {
  const isSub = lane.laneCode !== "MAIN";
  const { attributes, listeners, setNodeRef, isDragging } = useDraggable({
    id: dragMergeId(side.modelName, lane.laneCode),
    data: { kind: "merge", modelName: side.modelName, laneCode: lane.laneCode, label: `${lane.laneCode} 산출물` },
    disabled: !isSub
  });

  return (
    <div className={`t3r-lane-head ${isSub ? "sub" : "main"}`}>
      <div className="t3r-lane-head-top">
        <span className="t3r-lane-name">{lane.laneCode}</span>
        <span className="t3r-lane-count">{lane.processes.length}</span>
        {isSub && lane.processes.length === 0 && (
          <button type="button" className="t3r-icon-button" title="빈 SUB 레인 삭제" onClick={onRemove}>
            ×
          </button>
        )}
      </div>
      {isSub && (
        <div className="t3r-lane-merge">
          <button
            ref={setNodeRef}
            type="button"
            className={`t3r-merge-handle ${isDragging ? "is-dragging" : ""} ${isPicking ? "is-picking" : ""}`}
            title="MAIN 공정 셀로 끌어 놓거나 클릭한 뒤 MAIN 셀을 선택하세요"
            onClick={onStartPick}
            {...attributes}
            {...listeners}
          >
            {lane.laneCode} 산출물 → MAIN
          </button>
          <span className="t3r-merge-label" title={lane.mergeTargetLabel || "합류 위치 미지정"}>
            {lane.mergeTargetLabel || "합류 위치 미지정"}
          </span>
          {lane.mergeTargetProcessId && (
            <button type="button" className="t3r-link-button danger-text" onClick={onClearMerge}>
              해제
            </button>
          )}
        </div>
      )}
    </div>
  );
}

function MatrixCell({ side, lane, step, cellStep, activeDrag, isPicking, onPick, onUnassign }) {
  const process = cellStep?.process ?? null;
  const { setNodeRef, isOver } = useDroppable({
    id: dropCellId(side.modelName, lane.laneCode, step),
    data: {
      kind: "cell",
      modelName: side.modelName,
      laneCode: lane.laneCode,
      step,
      processId: process?.id ?? ""
    }
  });

  let dropState = "";
  if (activeDrag && activeDrag.modelName === side.modelName) {
    if (activeDrag.kind === "process") dropState = "can-drop";
    else if (activeDrag.kind === "material") dropState = process ? "can-drop" : "no-drop";
    else if (activeDrag.kind === "merge") dropState = lane.laneCode === "MAIN" && process ? "can-drop" : "no-drop";
  }
  const pickTarget = isPicking && lane.laneCode === "MAIN" && Boolean(process);
  const formula = cellStep ? formulaText(cellStep) : "";

  return (
    <div
      ref={setNodeRef}
      className={`t3r-cell ${lane.laneCode === "MAIN" ? "main" : "sub"} ${dropState} ${isOver ? "is-over" : ""} ${
        pickTarget ? "is-pick-target" : ""
      }`}
      onClick={pickTarget ? onPick : undefined}
    >
      {process ? (
        <>
          <CellProcess
            side={side}
            lane={lane}
            step={step}
            process={process}
            onUnassign={() => onUnassign(process.id)}
          />
          {cellStep.firstInputs.length > 0 && (
            <div className="t3r-chips">
              {cellStep.firstInputs.map((material) => (
                <span key={material.id} className="t3r-chip" title={`${material.materialCode} · ${material.materialName}`}>
                  {materialLabel(material)}
                </span>
              ))}
            </div>
          )}
          {cellStep.merges.map((laneCode) => (
            <span key={laneCode} className="t3r-merge-chip">
              + {laneCode} 합류
            </span>
          ))}
          {formula && (
            <div className="t3r-formula" title={formula}>
              {formula}
            </div>
          )}
        </>
      ) : (
        <span className="t3r-cell-empty">{step === 0 ? "공정을 끌어 놓으세요" : ""}</span>
      )}
    </div>
  );
}

function CellProcess({ side, lane, step, process, onUnassign }) {
  const { attributes, listeners, setNodeRef, isDragging } = useDraggable({
    id: dragProcessId(process.id),
    data: {
      kind: "process",
      processId: process.id,
      modelName: side.modelName,
      laneCode: lane.laneCode,
      step,
      process,
      label: process.processName || process.processCode
    }
  });

  return (
    <div ref={setNodeRef} className={`t3r-cell-process ${isDragging ? "is-dragging" : ""}`}>
      <button type="button" className="t3r-drag-handle" aria-label="공정 이동" {...attributes} {...listeners}>
        ⋮⋮
      </button>
      <div className="t3r-card-content">
        <div className="t3r-card-title">
          <code>{process.processCode}</code>
          <strong title={process.processName}>{process.processName}</strong>
        </div>
        <small>
          {process.processType}
          {process.mirrorProcessCode && process.mirrorProcessCode !== process.processCode && (
            <em className="t3r-mirror-code" title="반대쪽 모델에서 같이 배치되는 공정">
              ↔ {process.mirrorProcessCode}
            </em>
          )}
        </small>
      </div>
      <button
        type="button"
        className="t3r-icon-button"
        title="배치 해제"
        onClick={(event) => {
          event.stopPropagation();
          onUnassign();
        }}
      >
        ×
      </button>
    </div>
  );
}

function ProcessPalette({ side, query }) {
  const { setNodeRef, isOver } = useDroppable({
    id: dropPaletteId(side.modelName),
    data: { kind: "palette", modelName: side.modelName }
  });
  const processes = side.availableProcesses.filter(
    (process) =>
      !query ||
      process.processCode.toLocaleLowerCase().includes(query) ||
      process.processName.toLocaleLowerCase().includes(query)
  );

  return (
    <section ref={setNodeRef} className={`t3r-palette-group ${isOver ? "is-over" : ""}`}>
      <header className={`t3r-palette-group-head ${sideClass(side.sideLabel)}`}>
        <span className="t3r-side-badge">{side.sideLabel || "-"}</span>
        <strong title={side.modelName}>{side.modelName}</strong>
        <span className="t3r-lane-count">{processes.length}</span>
      </header>
      <div className="t3r-palette-list">
        {processes.map((process) => (
          <PaletteProcess key={process.id} process={process} side={side} />
        ))}
        {processes.length === 0 && <div className="t3r-empty">미배치 공정이 없습니다.</div>}
      </div>
    </section>
  );
}

function PaletteProcess({ process, side }) {
  const { attributes, listeners, setNodeRef, isDragging } = useDraggable({
    id: dragProcessId(process.id),
    data: {
      kind: "process",
      processId: process.id,
      modelName: side.modelName,
      laneCode: "",
      step: -1,
      process,
      label: process.processName || process.processCode
    }
  });

  return (
    <div ref={setNodeRef} className={`t3r-palette-card ${isDragging ? "is-dragging" : ""}`}>
      <button type="button" className="t3r-drag-handle" aria-label="공정 배치" {...attributes} {...listeners}>
        ⋮⋮
      </button>
      <div className="t3r-card-content">
        <div className="t3r-card-title">
          <code>{process.processCode}</code>
          <strong title={process.processName}>{process.processName}</strong>
        </div>
        {process.processType && <small>{process.processType}</small>}
      </div>
    </div>
  );
}

function MaterialPalette({ materials, onSave, onClear }) {
  return (
    <section className="t3r-palette-group">
      <div className="t3r-palette-list">
        {materials.map((material) => (
          <MaterialCard
            key={material.id}
            material={material}
            onSave={(usageQty, usageUnit) => onSave(material, usageQty, usageUnit)}
            onClear={() => onClear(material.id)}
          />
        ))}
        {materials.length === 0 && <div className="t3r-empty">표시할 BOM 자재가 없습니다.</div>}
      </div>
    </section>
  );
}

function MaterialCard({ material, onSave, onClear }) {
  const [usageQty, setUsageQty] = useState(material.usageQty || 1);
  const [usageUnit, setUsageUnit] = useState(material.usageUnit || "PC");
  const { attributes, listeners, setNodeRef, isDragging } = useDraggable({
    id: dragMaterialId(material.id),
    data: {
      kind: "material",
      materialId: material.id,
      modelName: material.modelName,
      usageQty,
      usageUnit,
      label: materialLabel(material)
    }
  });

  return (
    <div
      ref={setNodeRef}
      className={`t3r-material-card ${material.assignedProcessId ? "configured" : ""} ${isDragging ? "is-dragging" : ""}`}
    >
      <div className="t3r-material-head">
        <button
          type="button"
          className="t3r-drag-handle material"
          aria-label="자재를 최초 투입 공정 셀로 이동"
          {...attributes}
          {...listeners}
        >
          ⠿
        </button>
        <div className="t3r-card-content">
          <code>{material.materialCode}</code>
          <strong title={material.materialName}>{material.materialName}</strong>
          {material.mirrorMaterialName && material.mirrorMaterialName !== material.materialName && (
            <em className="t3r-mirror-code" title={`반대쪽 모델에는 ${material.mirrorMaterialCode} 로 투입됩니다`}>
              ↔ {material.mirrorMaterialName}
            </em>
          )}
        </div>
      </div>
      <div className="t3r-assignment">
        <span>최초 투입</span>
        <strong title={material.assignedProcessLabel}>
          {material.assignedProcessLabel || "공정 셀 위로 끌어 놓으세요"}
        </strong>
        <small>{material.scopeLabel}</small>
      </div>
      <div className="t3r-material-controls">
        <input
          type="number"
          min="0.000001"
          step="any"
          value={usageQty}
          onChange={(event) => setUsageQty(event.target.value)}
          aria-label="사용량"
        />
        <input value={usageUnit} onChange={(event) => setUsageUnit(event.target.value)} aria-label="사용 단위" />
        {material.assignedProcessId && (
          <button type="button" className="t3r-link-button" onClick={() => onSave(usageQty, usageUnit)}>
            저장
          </button>
        )}
        {material.assignedProcessId && (
          <button type="button" className="t3r-link-button danger-text" onClick={onClear}>
            해제
          </button>
        )}
      </div>
    </div>
  );
}

function DragPreview({ data }) {
  const icon = data.kind === "material" ? "⠿" : data.kind === "merge" ? "↘" : "↕";
  return (
    <div className={`t3r-drag-preview ${data.kind}`}>
      <span>{icon}</span>
      <strong>{data.label || data.process?.processName || data.process?.processCode || "이동"}</strong>
    </div>
  );
}

export function mount(element, dotnet, initialSnapshot) {
  const previous = roots.get(element);
  if (previous) previous.unmount();
  const root = createRoot(element);
  roots.set(element, root);
  root.render(<Test3Editor dotnet={dotnet} initialSnapshot={initialSnapshot} />);
}

export function unmount(element) {
  const root = roots.get(element);
  if (!root) return;
  root.unmount();
  roots.delete(element);
}
