import { useEffect, useRef, useState } from 'react';
import { X } from 'lucide-react';
import type { EditorActions, FindResult } from './ribbon/editor-actions';

export type FindTab = 'find' | 'replace';

/**
 * Caixa "Localizar e substituir" do Word (Ctrl+L / Ctrl+U): fica aberta sobre o documento
 * enquanto se edita, com "Localizar Próxima", "Substituir" e "Substituir Tudo".
 */
export function FindReplace(props: {
  tab: FindTab;
  initialQuery: string;
  editor: EditorActions;
  onTab: (tab: FindTab) => void;
  onClose: () => void;
}) {
  const { editor } = props;
  const [query, setQuery] = useState(props.initialQuery);
  const [replacement, setReplacement] = useState('');
  const [caseSensitive, setCaseSensitive] = useState(false);
  const [result, setResult] = useState<FindResult | null>(null);
  const [message, setMessage] = useState('');
  // A busca é refeita quando o texto ou a opção mudam; senão "Localizar Próxima" só avança.
  const searched = useRef<string | null>(null);
  const queryRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    queryRef.current?.focus();
    queryRef.current?.select();
  }, [props.tab]);
  useEffect(() => () => editor.clearFind(), [editor]);

  const key = `${caseSensitive ? 1 : 0}:${query}`;
  const startOrNext = (backwards = false) => {
    if (!query) return null;
    let r: FindResult;
    if (searched.current !== key) {
      r = editor.find(query, { caseSensitive });
      searched.current = key;
    } else {
      r = editor.findNext(backwards);
    }
    setResult(r);
    setMessage(r.count ? '' : 'O Textos concluiu a pesquisa do documento. O item não foi encontrado.');
    return r;
  };

  const replaceOne = () => {
    // Como no Word: o primeiro clique só localiza; os seguintes substituem e vão para a próxima.
    if (searched.current !== key) {
      startOrNext();
      return;
    }
    const r = editor.replaceCurrent(replacement);
    setResult(r);
    if (!r.count) setMessage('O Textos concluiu a pesquisa do documento.');
  };

  const replaceAll = () => {
    if (!query) return;
    editor.find(query, { caseSensitive });
    const n = editor.replaceAll(replacement);
    searched.current = null;
    setResult(null);
    setMessage(n ? `Concluído. Foram feitas ${n} ${n === 1 ? 'substituição' : 'substituições'}.` : 'O item não foi encontrado.');
  };

  const isReplace = props.tab === 'replace';
  return (
    <div
      className="find-dialog"
      role="dialog"
      aria-label="Localizar e substituir"
      onKeyDown={(e) => {
        if (e.key === 'Escape') props.onClose();
      }}
    >
      <div className="find-header">
        <span>Localizar e substituir</span>
        <button type="button" className="find-close" title="Fechar (Esc)" onClick={props.onClose}>
          <X size={14} />
        </button>
      </div>
      <div className="find-tabs" role="tablist">
        {(['find', 'replace'] as const).map((t) => (
          <button
            key={t}
            type="button"
            role="tab"
            aria-selected={props.tab === t}
            className={`find-tab${props.tab === t ? ' is-selected' : ''}`}
            onClick={() => props.onTab(t)}
          >
            {t === 'find' ? 'Localizar' : 'Substituir'}
          </button>
        ))}
      </div>
      <form
        className="find-body"
        onSubmit={(e) => {
          e.preventDefault();
          startOrNext();
        }}
        onKeyDown={(e) => {
          if (e.key === 'Enter' && e.shiftKey) {
            e.preventDefault();
            startOrNext(true);
          }
        }}
      >
        <label className="find-field">
          <span>Localizar:</span>
          <input ref={queryRef} value={query} onChange={(e) => setQuery(e.target.value)} />
        </label>
        {isReplace && (
          <label className="find-field">
            <span>Substituir por:</span>
            <input value={replacement} onChange={(e) => setReplacement(e.target.value)} />
          </label>
        )}
        <label className="find-check">
          <input type="checkbox" checked={caseSensitive} onChange={(e) => setCaseSensitive(e.target.checked)} />
          Diferenciar maiúsc./minúsc.
        </label>
        <div className="find-status" aria-live="polite">
          {message || (result && result.count ? `Resultado ${result.index + 1} de ${result.count}` : '')}
        </div>
        <div className="find-buttons">
          {isReplace && (
            <>
              <button type="button" onClick={replaceOne} disabled={!query}>
                Substituir
              </button>
              <button type="button" onClick={replaceAll} disabled={!query}>
                Substituir Tudo
              </button>
            </>
          )}
          <button type="button" onClick={() => startOrNext(true)} disabled={!query} title="Shift+Enter">
            Anterior
          </button>
          <button type="submit" className="dialog-ok" disabled={!query}>
            Localizar Próxima
          </button>
          <button type="button" onClick={props.onClose}>
            Cancelar
          </button>
        </div>
      </form>
    </div>
  );
}
