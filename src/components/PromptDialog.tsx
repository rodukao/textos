import { useEffect, useRef, useState } from 'react';

export interface PromptField {
  label: string;
  initial: string;
}

export interface PromptRequest {
  title: string;
  fields: PromptField[];
  resolve: (values: string[] | null) => void;
}

/**
 * Caixa de diálogo com um ou mais campos (ex.: "Inserir Hiperlink", "Margens Personalizadas"),
 * no lugar do prompt do navegador.
 */
export function PromptDialog({ request, onDone }: { request: PromptRequest; onDone: () => void }) {
  const [values, setValues] = useState(request.fields.map((f) => f.initial));
  const firstRef = useRef<HTMLInputElement>(null);
  useEffect(() => firstRef.current?.select(), []);
  const finish = (result: string[] | null) => {
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
          finish(values);
        }}
        onKeyDown={(e) => {
          if (e.key === 'Escape') finish(null);
        }}
      >
        <div className="dialog-title">{request.title}</div>
        <div className={request.fields.length > 1 ? 'dialog-grid' : undefined}>
          {request.fields.map((field, i) => (
            <label key={field.label} className="dialog-field">
              <span>{field.label}</span>
              <input
                ref={i === 0 ? firstRef : undefined}
                value={values[i]}
                onChange={(e) => setValues(values.map((v, j) => (j === i ? e.target.value : v)))}
              />
            </label>
          ))}
        </div>
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
