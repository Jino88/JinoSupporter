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
import {
  SortableContext,
  sortableKeyboardCoordinates,
  useSortable,
  verticalListSortingStrategy
} from "@dnd-kit/sortable";
import { CSS } from "@dnd-kit/utilities";
import "./style.css";

const roots = new WeakMap();
const processDragId = (id) => `process:${id}`;
const laneDropId = (modelName, laneCode) => `lane:${modelName}\u001f${laneCode}`;
const availableDropId = (modelName) => `available:${modelName}`;
const materialDragId = (id) => `material:${id}`;
const transferDragId = (modelName, laneCode) => `transfer:${modelName}\u001f${laneCode}`;

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

function nextSubLane(lanes) {
  const used = new Set(lanes.map((lane) => lane.laneCode.toUpperCase()));
  for (let index = 1; index <= 99; index += 1) {
    if (!used.has(`SUB${index}`)) return `SUB${index}`;
  }
  return "SUB99";
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

function Test3Editor({ dotnet, initialSnapshot }) {
  const initial = normalizeSnapshot(initialSnapshot);
  const [snapshot, setSnapshot] = useState(initial);
  const [sides, setSides] = useState(initial.sides);
  const [modelInput, setModelInput] = useState(initial.selectedModel);
  const [processSearch, setProcessSearch] = useState("");
  const [materialSearch, setMaterialSearch] = useState("");
  const [busy, setBusy] = useState(false);
  const [activeDrag, setActiveDrag] = useState(null);
  const requestNumber = useRef(0);
  const sensors = useSensors(
    useSensor(PointerSensor, { activationConstraint: { distance: 6 } }),
    useSensor(KeyboardSensor, { coordinateGetter: sortableKeyboardCoordinates })
  );

  const materialsByProcess = useMemo(() => {
    const result = new Map();
    for (const material of snapshot.materials) {
      if (!material.assignedProcessId) continue;
      const side = sides.find((item) => item.modelName === material.modelName);
      const sourceLane = side?.lanes.find((lane) => lane.processes.some((process) => process.id === material.assignedProcessId));
      if (!side || !sourceLane) continue;
      const sourceIndex = sourceLane.processes.findIndex((process) => process.id === material.assignedProcessId);
      const effectiveProcesses = [...sourceLane.processes.slice(sourceIndex)];
      if (sourceLane.laneCode !== "MAIN" && sourceLane.mergeTargetProcessId) {
        const mainLane = side.lanes.find((lane) => lane.laneCode === "MAIN");
        const mergeIndex = mainLane?.processes.findIndex((process) => process.id === sourceLane.mergeTargetProcessId) ?? -1;
        if (mainLane && mergeIndex >= 0) effectiveProcesses.push(...mainLane.processes.slice(mergeIndex));
      }
      for (const process of effectiveProcesses) {
        const values = result.get(process.id) ?? [];
        values.push({ ...material, isFirstInput: process.id === material.assignedProcessId });
        result.set(process.id, values);
      }
    }
    return result;
  }, [snapshot.materials, sides]);

  const applySnapshot = (value) => {
    const next = normalizeSnapshot(value);
    setSnapshot(next);
    setSides(next.sides);
    setModelInput(next.selectedModel);
  };

  const invoke = async (method, ...args) => {
    const currentRequest = ++requestNumber.current;
    setBusy(true);
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
    await invoke("ReactSelectModelAsync", modelInput.trim(), force);
  };

  const changeModelInput = (value) => {
    setModelInput(value);
    const exactModel = snapshot.modelNames.find((model) => model.toLocaleLowerCase() === value.trim().toLocaleLowerCase());
    if (exactModel) void invoke("ReactSelectModelAsync", exactModel, false);
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
        return { ...side, lanes: [...side.lanes, { laneCode, mergeTargetProcessId: "", mergeTargetLabel: "", processes: [] }] };
      })
    );
  };

  const removeEmptyLane = (modelName, laneCode) => {
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
    const mainProcess = side.lanes.find((item) => item.laneCode === "MAIN")?.processes.find((item) => item.id === processId);
    lane.mergeTargetProcessId = processId;
    lane.mergeTargetLabel = mainProcess ? `MAIN #${mainProcess.order || "?"} · ${mainProcess.processName}` : "";
    void persistLayout(next);
  };

  const moveProcess = (activeData, overData) => {
    if (overData.kind === "process" && overData.processId === activeData.processId) return;
    const next = cloneSides(sides);
    const sourceSide = next.find((side) => side.modelName === activeData.modelName);
    if (!sourceSide) return;
    const originalSourceLane = sourceSide.lanes.find((lane) => lane.processes.some((item) => item.id === activeData.processId));
    const originalSourceIndex = originalSourceLane?.processes.findIndex((item) => item.id === activeData.processId) ?? -1;
    const originalTargetLane = overData.laneCode
      ? sourceSide.lanes.find((lane) => lane.laneCode === overData.laneCode)
      : null;
    const originalTargetIndex = overData.kind === "process"
      ? originalTargetLane?.processes.findIndex((item) => item.id === overData.processId) ?? -1
      : -1;
    let process = sourceSide.availableProcesses.find((item) => item.id === activeData.processId);
    sourceSide.availableProcesses = sourceSide.availableProcesses.filter((item) => item.id !== activeData.processId);
    for (const lane of sourceSide.lanes) {
      const match = lane.processes.find((item) => item.id === activeData.processId);
      if (match) process = match;
      lane.processes = lane.processes.filter((item) => item.id !== activeData.processId);
    }
    if (!process) return;

    if (overData.kind === "available") {
      if (overData.modelName !== sourceSide.modelName) return;
      sourceSide.availableProcesses.push({ ...process, laneCode: "", processNo: "", order: 0 });
      void persistLayout(next);
      return;
    }

    const targetModel = overData.modelName;
    const targetLaneCode = overData.laneCode;
    if (targetModel !== sourceSide.modelName || !targetLaneCode) return;
    const targetSide = next.find((side) => side.modelName === targetModel);
    const targetLane = targetSide?.lanes.find((lane) => lane.laneCode === targetLaneCode);
    if (!targetLane) return;
    const targetIndex = overData.kind === "process"
      ? targetLane.processes.findIndex((item) => item.id === overData.processId)
      : targetLane.processes.length;
    let insertIndex = targetIndex < 0 ? targetLane.processes.length : targetIndex;
    if (originalSourceLane?.laneCode === targetLaneCode &&
        originalSourceIndex >= 0 &&
        originalTargetIndex > originalSourceIndex) {
      insertIndex += 1;
    }
    targetLane.processes.splice(insertIndex, 0, { ...process, laneCode: targetLaneCode });
    void persistLayout(next);
  };

  const handleDragEnd = ({ active, over }) => {
    const activeData = active.data.current;
    const overData = over?.data.current;
    setActiveDrag(null);
    if (!activeData || !overData) return;

    if (activeData.kind === "material" && overData.kind === "process") {
      if (activeData.modelName !== overData.modelName) {
        setSnapshot((current) => ({ ...current, statusMessage: "L/R가 다른 공정에는 자재를 투입할 수 없습니다.", statusError: true }));
        return;
      }
      void invoke("ReactSaveMaterialAsync", {
        materialId: activeData.materialId,
        processId: overData.processId,
        usageQty: Number(activeData.usageQty) || 1,
        usageUnit: activeData.usageUnit || "PC"
      });
      return;
    }

    if (activeData.kind === "transfer" && overData.kind === "process") {
      if (activeData.modelName === overData.modelName && overData.laneCode === "MAIN") {
        updateMergeTarget(activeData.modelName, activeData.laneCode, overData.processId);
      }
      return;
    }

    if (activeData.kind === "process") moveProcess(activeData, overData);
  };

  const filteredMaterials = snapshot.materials.filter((material) => {
    const query = materialSearch.trim().toLocaleLowerCase();
    return !query || material.materialCode.toLocaleLowerCase().includes(query) || material.materialName.toLocaleLowerCase().includes(query);
  });

  return (
    <DndContext
      sensors={sensors}
      collisionDetection={closestCenter}
      onDragStart={({ active }) => setActiveDrag(active.data.current ?? null)}
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
              onKeyDown={(event) => { if (event.key === "Enter") void loadModel(false); }}
            />
            <datalist id="t3r-model-options">
              {snapshot.modelNames.map((model) => <option key={model} value={model} />)}
            </datalist>
            <button type="button" className="t3r-button primary" onClick={() => void loadModel(false)} disabled={busy}>검색</button>
          </div>
          <div className="t3r-toolbar-actions">
            <input value={processSearch} onChange={(event) => setProcessSearch(event.target.value)} placeholder="공정 검색" />
            <input value={materialSearch} onChange={(event) => setMaterialSearch(event.target.value)} placeholder="BOM 자재 검색" />
            <button type="button" className="t3r-button" onClick={() => void invoke("ReactReloadBomAsync")} disabled={busy || sides.length === 0}>BOM 새로고침</button>
            <button type="button" className="t3r-button danger" onClick={() => {
              if (window.confirm("현재 L/R 공정 배치를 모두 초기화할까요?")) void invoke("ReactResetLayoutAsync");
            }} disabled={busy || sides.length === 0}>공정 초기화</button>
          </div>
        </header>

        {snapshot.statusMessage && (
          <div className={`t3r-status ${snapshot.statusError ? "error" : "success"}`}>{snapshot.statusMessage}</div>
        )}
        {snapshot.bomError && <div className="t3r-status error">{snapshot.bomError}</div>}

        {sides.length === 0 ? (
          <div className="t3r-welcome">
            <strong>모델을 검색하면 편집 영역이 열립니다.</strong>
            <span>기준 모델을 선택하면 연결된 L/R 모델을 한 화면에서 함께 편집할 수 있습니다.</span>
          </div>
        ) : (
          <div className="t3r-grid">
            <section className="t3r-column route-column">
              <ColumnTitle index="1" title="저장된 공정 순서" count={sides.reduce((sum, side) => sum + side.lanes.reduce((laneSum, lane) => laneSum + lane.processes.length, 0), 0)} />
              <div className="t3r-column-body">
                {sides.map((side) => (
                  <SideRoute
                    key={side.modelName}
                    side={side}
                    materialsByProcess={materialsByProcess}
                    onAddSub={() => addSubLane(side.modelName)}
                    onRemoveLane={(laneCode) => removeEmptyLane(side.modelName, laneCode)}
                    onMergeTarget={(laneCode, processId) => updateMergeTarget(side.modelName, laneCode, processId)}
                  />
                ))}
              </div>
            </section>

            <section className="t3r-column available-column">
              <ColumnTitle index="2" title="검색된 모델의 공정" count={sides.reduce((sum, side) => sum + side.availableProcesses.length, 0)} />
              <div className="t3r-column-body">
                {sides.map((side) => (
                  <AvailableProcesses key={side.modelName} side={side} search={processSearch} />
                ))}
              </div>
            </section>

            <section className="t3r-column material-column">
              <ColumnTitle index="3" title="BOM 자재 최초 투입 공정" count={filteredMaterials.length} detail={snapshot.bomSource} />
              <div className="t3r-column-body">
                {sides.map((side) => (
                  <MaterialSide
                    key={side.modelName}
                    side={side}
                    materials={filteredMaterials.filter((material) => material.modelName === side.modelName)}
                    onSave={(material, usageQty, usageUnit) => void invoke("ReactSaveMaterialAsync", {
                      materialId: material.id,
                      processId: material.assignedProcessId,
                      usageQty: Number(usageQty) || 1,
                      usageUnit: usageUnit || "PC"
                    })}
                    onClear={(materialId) => void invoke("ReactClearMaterialAsync", materialId)}
                  />
                ))}
              </div>
            </section>
          </div>
        )}
        {busy && <div className="t3r-busy"><span />저장 중…</div>}
      </section>

      <DragOverlay dropAnimation={null}>
        {activeDrag ? <DragPreview data={activeDrag} /> : null}
      </DragOverlay>
    </DndContext>
  );
}

function ColumnTitle({ index, title, count, detail }) {
  return (
    <div className="t3r-column-title">
      <span className="t3r-step">{index}</span>
      <strong>{title}</strong>
      <span className="t3r-count">{count}</span>
      {detail && <small>{detail}</small>}
    </div>
  );
}

function SideHeader({ side, actions }) {
  return (
    <div className={`t3r-side-header ${sideClass(side.sideLabel)}`}>
      <span className="t3r-side-badge">{side.sideLabel || "-"}</span>
      <strong>{side.modelName}</strong>
      {actions}
    </div>
  );
}

function SideRoute({ side, materialsByProcess, onAddSub, onRemoveLane, onMergeTarget }) {
  const mainProcesses = side.lanes.find((lane) => lane.laneCode === "MAIN")?.processes ?? [];
  return (
    <article className="t3r-side-block">
      <SideHeader side={side} actions={<button type="button" className="t3r-link-button" onClick={onAddSub}>+ SUB 추가</button>} />
      <div className="t3r-lanes">
        {side.lanes.map((lane) => (
          <ProcessLane
            key={`${side.modelName}:${lane.laneCode}`}
            side={side}
            lane={lane}
            mainProcesses={mainProcesses}
            materialsByProcess={materialsByProcess}
            onRemove={() => onRemoveLane(lane.laneCode)}
            onMergeTarget={(processId) => onMergeTarget(lane.laneCode, processId)}
          />
        ))}
      </div>
    </article>
  );
}

function ProcessLane({ side, lane, mainProcesses, materialsByProcess, onRemove, onMergeTarget }) {
  const { setNodeRef, isOver } = useDroppable({
    id: laneDropId(side.modelName, lane.laneCode),
    data: { kind: "lane", modelName: side.modelName, laneCode: lane.laneCode }
  });
  const isSub = lane.laneCode !== "MAIN";
  return (
    <section ref={setNodeRef} className={`t3r-lane ${isSub ? "sub" : "main"} ${isOver ? "is-over" : ""}`}>
      <div className="t3r-lane-header">
        <span className="t3r-lane-name">{lane.laneCode}</span>
        <span>{lane.processes.length} 공정</span>
        {isSub && <TransferHandle side={side} lane={lane} />}
        {isSub && lane.processes.length === 0 && <button type="button" className="t3r-icon-button" title="빈 SUB 삭제" onClick={onRemove}>×</button>}
      </div>
      {isSub && (
        <div className="t3r-merge-row">
          <span>MAIN 합류</span>
          <select value={lane.mergeTargetProcessId ?? ""} onChange={(event) => onMergeTarget(event.target.value)} disabled={mainProcesses.length === 0}>
            <option value="">합류 공정 선택</option>
            {mainProcesses.map((process, index) => <option key={process.id} value={process.id}>MAIN #{index + 1} · {process.processName}</option>)}
          </select>
        </div>
      )}
      <SortableContext items={lane.processes.map((process) => processDragId(process.id))} strategy={verticalListSortingStrategy}>
        <div className="t3r-process-list">
          {lane.processes.map((process, index) => (
            <SortableProcess
              key={process.id}
              process={process}
              index={index}
              modelName={side.modelName}
              laneCode={lane.laneCode}
              materials={materialsByProcess.get(process.id) ?? []}
            />
          ))}
          {lane.processes.length === 0 && <div className="t3r-drop-hint">공정을 이 라인으로 끌어오세요</div>}
        </div>
      </SortableContext>
    </section>
  );
}

function SortableProcess({ process, index, modelName, laneCode, materials }) {
  const { attributes, listeners, setNodeRef, transform, transition, isDragging, isOver } = useSortable({
    id: processDragId(process.id),
    data: { kind: "process", processId: process.id, modelName, laneCode, process }
  });
  return (
    <div ref={setNodeRef} style={{ transform: CSS.Transform.toString(transform), transition }} className={`t3r-process-card ${isDragging ? "is-dragging" : ""} ${isOver ? "is-over" : ""}`}>
      <button type="button" className="t3r-drag-handle" aria-label="공정 이동" {...attributes} {...listeners}>⋮⋮</button>
      <span className="t3r-process-no">{index + 1}</span>
      <div className="t3r-card-content">
        <div className="t3r-card-title"><code>{process.processCode}</code><strong>{process.processName}</strong></div>
        {process.processType && <small>{process.processType}</small>}
        <div className={`t3r-material-chips ${materials.length ? "configured" : ""}`}>
          {materials.length === 0 ? <span>투입 자재 없음</span> : materials.slice(0, 3).map((material) => (
            <span key={material.id} className={material.isFirstInput ? "first-input" : "carried"} title={material.isFirstInput ? "이 공정에서 최초 투입" : "이전 공정에서 계속 사용"}>
              {material.isFirstInput ? "+ " : "↳ "}{material.materialName || material.materialCode}
            </span>
          ))}
          {materials.length > 3 && <span>+{materials.length - 3}</span>}
        </div>
      </div>
    </div>
  );
}

function TransferHandle({ side, lane }) {
  const { attributes, listeners, setNodeRef, isDragging } = useDraggable({
    id: transferDragId(side.modelName, lane.laneCode),
    data: { kind: "transfer", modelName: side.modelName, laneCode: lane.laneCode, label: `${lane.laneCode} 완성품` }
  });
  return (
    <button ref={setNodeRef} type="button" className={`t3r-transfer-handle ${isDragging ? "is-dragging" : ""}`} title="MAIN 공정으로 끌어 합류 위치 지정" {...attributes} {...listeners}>
      {lane.laneCode} 출력 ↗
    </button>
  );
}

function AvailableProcesses({ side, search }) {
  const { setNodeRef, isOver } = useDroppable({
    id: availableDropId(side.modelName),
    data: { kind: "available", modelName: side.modelName }
  });
  const query = search.trim().toLocaleLowerCase();
  const processes = side.availableProcesses.filter((process) =>
    !query || process.processCode.toLocaleLowerCase().includes(query) || process.processName.toLocaleLowerCase().includes(query)
  );
  return (
    <article ref={setNodeRef} className={`t3r-side-block t3r-available ${isOver ? "is-over" : ""}`}>
      <SideHeader side={side} actions={<span>{processes.length}</span>} />
      <div className="t3r-available-list">
        {processes.map((process) => <AvailableProcess key={process.id} process={process} side={side} />)}
        {processes.length === 0 && <div className="t3r-empty">미배치 공정이 없습니다.</div>}
      </div>
    </article>
  );
}

function AvailableProcess({ process, side }) {
  const { attributes, listeners, setNodeRef, transform, isDragging } = useDraggable({
    id: processDragId(process.id),
    data: { kind: "process", processId: process.id, modelName: side.modelName, laneCode: "", process }
  });
  return (
    <div ref={setNodeRef} style={{ transform: CSS.Translate.toString(transform) }} className={`t3r-process-card compact ${isDragging ? "is-dragging" : ""}`}>
      <button type="button" className="t3r-drag-handle" aria-label="공정 이동" {...attributes} {...listeners}>⋮⋮</button>
      <div className="t3r-card-content">
        <div className="t3r-card-title"><code>{process.processCode}</code><strong>{process.processName}</strong></div>
        {process.processType && <small>{process.processType}</small>}
      </div>
    </div>
  );
}

function MaterialSide({ side, materials, onSave, onClear }) {
  return (
    <article className="t3r-side-block">
      <SideHeader side={side} actions={<span>{materials.length}</span>} />
      <div className="t3r-material-list">
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
    </article>
  );
}

function MaterialCard({ material, onSave, onClear }) {
  const [usageQty, setUsageQty] = useState(material.usageQty || 1);
  const [usageUnit, setUsageUnit] = useState(material.usageUnit || "PC");
  const { attributes, listeners, setNodeRef, transform, isDragging } = useDraggable({
    id: materialDragId(material.id),
    data: {
      kind: "material",
      materialId: material.id,
      modelName: material.modelName,
      usageQty,
      usageUnit,
      label: material.materialName || material.materialCode
    }
  });
  return (
    <div ref={setNodeRef} style={{ transform: CSS.Translate.toString(transform) }} className={`t3r-material-card ${material.assignedProcessId ? "configured" : ""} ${isDragging ? "is-dragging" : ""}`}>
      <div className="t3r-material-head">
        <button type="button" className="t3r-drag-handle material" aria-label="자재를 최초 투입 공정으로 이동" {...attributes} {...listeners}>⠿</button>
        <div className="t3r-card-content">
          <code>{material.materialCode}</code>
          <strong>{material.materialName}</strong>
        </div>
      </div>
      <div className="t3r-assignment">
        <span>최초 투입</span>
        <strong>{material.assignedProcessLabel || "공정 카드 위로 끌어 놓으세요"}</strong>
        <small>{material.scopeLabel}</small>
      </div>
      <div className="t3r-material-controls">
        <input type="number" min="0.000001" step="any" value={usageQty} onChange={(event) => setUsageQty(event.target.value)} aria-label="사용량" />
        <input value={usageUnit} onChange={(event) => setUsageUnit(event.target.value)} aria-label="사용 단위" />
        {material.assignedProcessId && <button type="button" className="t3r-link-button" onClick={() => onSave(usageQty, usageUnit)}>저장</button>}
        {material.assignedProcessId && <button type="button" className="t3r-link-button danger-text" onClick={onClear}>해제</button>}
      </div>
    </div>
  );
}

function DragPreview({ data }) {
  return (
    <div className="t3r-drag-preview">
      <span>↕</span>
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
