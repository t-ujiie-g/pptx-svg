/** TopBar — app title, undo/redo, open, export, status. */
import { DropZone } from './DropZone';

interface Props {
  status: string;
  canUndo: boolean;
  canRedo: boolean;
  onUndo: () => void;
  onRedo: () => void;
  onExport: () => void;
  onOpen: (file: File) => void;
}

export function TopBar({ status, canUndo, canRedo, onUndo, onRedo, onExport, onOpen }: Props) {
  return (
    <header className="topbar">
      <div className="brand">pptx-svg <span>Editor</span></div>
      <div className="tb-group">
        <button className="icon" title="Undo (⌘Z)" disabled={!canUndo} onClick={onUndo}>↩</button>
        <button className="icon" title="Redo (⌘⇧Z)" disabled={!canRedo} onClick={onRedo}>↪</button>
      </div>
      <div className="status" title={status}>{status}</div>
      <div className="tb-group right">
        <DropZone onFile={onOpen} compact />
        <button className="primary" onClick={onExport}>Export .pptx</button>
      </div>
    </header>
  );
}
