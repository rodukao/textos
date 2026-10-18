import { useState } from 'react';
import { Redo2, Save, Search, Undo2 } from 'lucide-react';
import { withShortcut } from '../shortcuts';

/**
 * Barra de título do Word: ícone, Barra de Ferramentas de Acesso Rápido (Salvar, Desfazer,
 * Refazer), nome do documento ao centro e a caixa "Pesquisar".
 */
export function TitleBar(props: {
  title: string;
  dirty: boolean;
  canUndo: boolean;
  canRedo: boolean;
  onSave: () => void;
  onUndo: () => void;
  onRedo: () => void;
  onSearch: (text: string, backwards: boolean) => void;
}) {
  const [query, setQuery] = useState('');
  const keep = (e: React.MouseEvent) => e.preventDefault();
  return (
    <header className="titlebar">
      <div className="qat">
        <img className="app-icon" src="/icon.svg" alt="" width={18} height={18} />
        <button type="button" className="qat-button" title={withShortcut('Salvar', 'save')} onMouseDown={keep} onClick={props.onSave}>
          <Save size={15} />
        </button>
        <button
          type="button"
          className="qat-button"
          title={withShortcut('Desfazer', 'undo')}
          disabled={!props.canUndo}
          onMouseDown={keep}
          onClick={props.onUndo}
        >
          <Undo2 size={15} />
        </button>
        <button
          type="button"
          className="qat-button"
          title={withShortcut('Refazer', 'redo')}
          disabled={!props.canRedo}
          onMouseDown={keep}
          onClick={props.onRedo}
        >
          <Redo2 size={15} />
        </button>
      </div>
      <div className="title-text" title={props.title}>
        {props.title}
        {props.dirty && <span className="title-dirty"> • Não salvo</span>}
        <span className="title-app"> - Textos</span>
      </div>
      <label className="title-search">
        <Search size={14} />
        <input
          type="search"
          placeholder="Pesquisar no documento"
          value={query}
          onChange={(e) => setQuery(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === 'Enter' && query.trim()) {
              e.preventDefault();
              props.onSearch(query, e.shiftKey);
            }
          }}
        />
      </label>
    </header>
  );
}
