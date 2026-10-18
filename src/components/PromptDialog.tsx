import { useEffect, useRef, useState } from 'react';

export interface PromptRequest {
  title: string;
  label: string;
  initial: string;
  resolve: (value: string | null) => void;
}

/** Caixa de diálogo de uma linha (ex.: "Inserir Hiperlink"), no lugar do prompt do navegador. */
export function PromptDialog({ request, onDone }: { request: PromptRequest; onDone: () => void }) {
  const [value, setValue] = useState(request.initial);
  const inputRef = useRef<HTMLInputElement>(null);
  useEffect(() => inputRef.current?.select(), []);
  const finish = (result: string | null) => {
    request.resolve(result);
    onDone();
  };
  return (
    <div className="dialog-backdrop" onMouseDown={() => finish(null)}>
      <form
        className="dialog"
        role="dialog"
        aria-label={request.title}
        onMouseDown={(e) => e.stopPropagation()}
        onSubmit={(e) => {
          e.preventDefault();
          finish(value);
        }}
        onKeyDown={(e) => {
          if (e.key === 'Escape') finish(null);
        }}
      >
        <div className="dialog-title">{request.title}</div>
        <label className="dialog-field">
          <span>{request.label}</span>
          <input ref={inputRef} value={value} onChange={(e) => setValue(e.target.value)} />
        </label>
        <div className="dialog-buttons">
          <button type="submit" className="dialog-ok">
            OK
          </button>
          <button type="button" onClick={() => finish(null)}>
            Cancelar
          </button>
        </div>
      </form>
    </div>
  );
}
