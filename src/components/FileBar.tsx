import { FilePlus, FolderOpen, Printer, Save, SaveAll, type LucideIcon } from 'lucide-react';

export interface FileActions {
  newDocument: () => void;
  open: () => void;
  save: () => void;
  saveAs: () => void;
  print: () => void;
}

function FileButton(props: { icon: LucideIcon; label: string; shortcut?: string; onClick: () => void }) {
  const Icon = props.icon;
  return (
    <button
      type="button"
      className="file-button"
      title={props.shortcut ? `${props.label} (${props.shortcut})` : props.label}
      onClick={props.onClick}
    >
      <Icon size={16} />
      <span>{props.label}</span>
    </button>
  );
}

export function FileBar(props: { title: string; dirty: boolean; actions: FileActions }) {
  const { actions } = props;
  return (
    <header className="filebar">
      <div className="brand">
        <img src="/icon.svg" alt="" width={22} height={22} />
        <span className="doc-title" title={props.title}>
          {props.title}
          {props.dirty && <span className="dirty-dot" title="Alterações não salvas"> •</span>}
        </span>
      </div>
      <nav className="file-actions" aria-label="Arquivo">
        <FileButton icon={FilePlus} label="Novo" shortcut="Ctrl+Alt+N" onClick={actions.newDocument} />
        <FileButton icon={FolderOpen} label="Abrir" shortcut="Ctrl+O" onClick={actions.open} />
        <FileButton icon={Save} label="Salvar" shortcut="Ctrl+S" onClick={actions.save} />
        <FileButton icon={SaveAll} label="Salvar como" shortcut="Ctrl+Shift+S" onClick={actions.saveAs} />
        <FileButton icon={Printer} label="Imprimir / PDF" shortcut="Ctrl+P" onClick={actions.print} />
      </nav>
    </header>
  );
}
