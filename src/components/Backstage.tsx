import { useEffect } from 'react';
import { ArrowLeft, FilePlus, FileText, FolderOpen, Printer, Save, SaveAll, type LucideIcon } from 'lucide-react';
import type { FileActions } from './file-actions';

/**
 * Modo de exibição Backstage (guia "Arquivo" do Word): coluna de comandos à esquerda e, na
 * página inicial, o "Documento em branco" e o botão de abrir.
 */
export function Backstage(props: { actions: FileActions; onClose: () => void }) {
  const { actions, onClose } = props;
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onClose();
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, [onClose]);

  const then = (action: () => void) => () => {
    onClose();
    action();
  };

  return (
    <div className="backstage" role="dialog" aria-label="Arquivo">
      <nav className="backstage-nav">
        <button type="button" className="backstage-back" title="Voltar (Esc)" onClick={onClose}>
          <ArrowLeft size={20} />
        </button>
        <NavItem icon={FilePlus} label="Novo" onClick={then(actions.newDocument)} />
        <NavItem icon={FolderOpen} label="Abrir" onClick={then(actions.open)} />
        <div className="backstage-divider" />
        <NavItem icon={Save} label="Salvar" onClick={then(actions.save)} />
        <NavItem icon={SaveAll} label="Salvar como" onClick={then(actions.saveAs)} />
        <NavItem icon={Printer} label="Imprimir" onClick={then(actions.print)} />
      </nav>
      <div className="backstage-main">
        <h1>{greeting()}</h1>
        <h2>Novo</h2>
        <button type="button" className="new-blank" onClick={then(actions.newDocument)}>
          <span className="new-blank-page" />
          <span>Documento em branco</span>
        </button>
        <h2>Abrir</h2>
        <button type="button" className="backstage-open" onClick={then(actions.open)}>
          <FileText size={18} />
          Procurar um documento do Word (.docx) no computador
        </button>
        <p className="backstage-about">
          Textos — editor de textos de código aberto, compatível com o Microsoft Word. Os documentos
          nunca saem do seu computador.
        </p>
      </div>
    </div>
  );
}

/** Como a página inicial do Word: "Bom dia", "Boa tarde" ou "Boa noite". */
function greeting() {
  const h = new Date().getHours();
  return h >= 5 && h < 12 ? 'Bom dia' : h >= 12 && h < 18 ? 'Boa tarde' : 'Boa noite';
}

function NavItem(props: { icon: LucideIcon; label: string; onClick: () => void }) {
  const Icon = props.icon;
  return (
    <button type="button" className="backstage-item" onClick={props.onClick}>
      <Icon size={18} strokeWidth={1.6} />
      <span>{props.label}</span>
    </button>
  );
}
