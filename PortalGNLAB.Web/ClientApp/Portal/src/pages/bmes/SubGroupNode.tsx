import { totalLineShifts, type SubVm } from "./modelGroupLogic";

export interface SubGroupHandlers {
  activeSub: string | null;
  draggingId: string | null;
  canDrop: (id: string) => boolean;
  onSelect: (id: string) => void;
  onAddNested: (id: string) => void;
  onRemove: (id: string) => void;
  onRename: (id: string, name: string) => void;
  onRemoveLineShift: (id: string, lineShift: string) => void;
  onDragStart: (id: string) => void;
  onDragEnd: () => void;
  onDrop: (targetId: string) => void;
}

interface Props extends SubGroupHandlers {
  node: SubVm;
  depth: number;
  hideHeader: boolean;
}

/** One sub group; a sub group holds Shift Groups and further sub groups (recursive). */
export default function SubGroupNode({ node, depth, hideHeader, ...h }: Props) {
  const active = h.activeSub === node.id;
  const classes = ["sub-grp", active && "active", h.canDrop(node.id) && "drop", h.draggingId === node.id && "dragging"].filter(Boolean).join(" ");
  return (
    <div
      className={classes}
      style={{ marginLeft: depth * 10 }}
      role="treeitem"
      aria-level={depth + 1}
      onClick={(e) => { e.stopPropagation(); h.onSelect(node.id); }}
      onDragOver={(e) => e.preventDefault()}
      onDrop={(e) => { e.preventDefault(); e.stopPropagation(); h.onDrop(node.id); }}
    >
      {!hideHeader && (
        <div className="sub-head">
          {active && <span className="grp-dot" title="Shift Groups go into this sub group">●</span>}
          <span
            className="tag-badge sub drag-handle"
            title="Drag to move this sub group under another sub group"
            draggable
            onDragStart={(e) => { e.dataTransfer.setData("text/plain", node.id); h.onDragStart(node.id); }}
            onDragEnd={h.onDragEnd}
            onClick={(e) => e.stopPropagation()}
          >Sub Group {depth + 1}</span>
          <input className="mg-name sub" value={node.name} placeholder="Sub Group name"
            onChange={(e) => h.onRename(node.id, e.target.value)} onClick={(e) => e.stopPropagation()} />
          <span className="bm-sub" style={{ fontFamily: "Consolas, monospace" }}>{totalLineShifts(node)} shifts</span>
          <button className="btn sm mg-mini" style={{ marginLeft: "auto" }} title="Add Sub Group"
            onClick={(e) => { e.stopPropagation(); h.onAddNested(node.id); }}>+ Sub Group</button>
          <button className="btn sm danger mg-mini" title="Remove sub group"
            onClick={(e) => { e.stopPropagation(); h.onRemove(node.id); }}>✕</button>
        </div>
      )}
      {node.lineShifts.length === 0 && node.subGroups.length === 0 && <div className="mg-empty">(empty sub group)</div>}
      {node.lineShifts.length > 0 && (
        <ul className="shift-list">
          {node.lineShifts.map((ls) => (
            <li key={ls}>
              <span className="tag-badge shift" title="Shift Group">Shift Group</span>
              <span>{ls}</span>
              <button className="shift-remove" title="Remove" onClick={(e) => { e.stopPropagation(); h.onRemoveLineShift(node.id, ls); }}>✕</button>
            </li>
          ))}
        </ul>
      )}
      {node.subGroups.length > 0 && (
        <div className="sub-list" role="group">
          {node.subGroups.map((child) => <SubGroupNode key={child.id} node={child} depth={depth + 1} hideHeader={false} {...h} />)}
        </div>
      )}
    </div>
  );
}
