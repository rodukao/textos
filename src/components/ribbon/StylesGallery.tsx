import { useRef, useState } from 'react';
import { ChevronDown } from 'lucide-react';
import { popupPosition } from './controls';
import type { QuickStyle } from './editor-actions';
import { styleDisplayName } from './word-data';

/**
 * Galeria "Estilos" da Página Inicial: amostras "AaBbCcDd" com o nome do estilo embaixo, como no
 * Word. A seta abre a galeria inteira.
 */
export function StylesGallery(props: {
  styles: QuickStyle[];
  activeId: string;
  disabled: boolean;
  onApply: (source: Record<string, unknown>) => void;
}) {
  const [expanded, setExpanded] = useState(false);
  const galleryRef = useRef<HTMLDivElement>(null);
  const tile = (s: QuickStyle) => (
    <button
      key={s.id}
      type="button"
      className={`style-tile${s.id === props.activeId ? ' is-active' : ''}`}
      title={styleDisplayName(s.name)}
      disabled={props.disabled}
      onMouseDown={(e) => e.preventDefault()}
      onClick={() => {
        setExpanded(false);
        props.onApply(s.source);
      }}
    >
      <span className="style-sample" style={s.preview}>
        AaBbCcDd
      </span>
      <span className="style-name">{styleDisplayName(s.name)}</span>
    </button>
  );

  if (!props.styles.length) return <div className="styles-empty">Sem estilos</div>;
  return (
    <div className="styles-gallery" ref={galleryRef}>
      <div className="styles-strip">{props.styles.map(tile)}</div>
      <button
        type="button"
        className="styles-more"
        title="Mais estilos"
        aria-expanded={expanded}
        onMouseDown={(e) => e.preventDefault()}
        onClick={() => setExpanded(!expanded)}
      >
        <ChevronDown size={12} />
      </button>
      {expanded && (
        <div className="rb-popup styles-popup" style={popupPosition(galleryRef.current, 450)} onMouseLeave={() => setExpanded(false)}>
          {props.styles.map(tile)}
        </div>
      )}
    </div>
  );
}
