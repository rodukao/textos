import { Minus, Plus } from 'lucide-react';

/** Limites do zoom do Word; no controle deslizante, 100% fica no meio. */
export const ZOOM_MIN = 10;
export const ZOOM_MAX = 500;

const toSlider = (zoom: number) =>
  zoom <= 100 ? ((zoom - ZOOM_MIN) / (100 - ZOOM_MIN)) * 50 : 50 + ((zoom - 100) / (ZOOM_MAX - 100)) * 50;
const fromSlider = (pos: number) =>
  Math.round(pos <= 50 ? ZOOM_MIN + (pos / 50) * (100 - ZOOM_MIN) : 100 + ((pos - 50) / 50) * (ZOOM_MAX - 100));

/** Os botões − e + do Word andam de 10 em 10%. */
const stepZoom = (zoom: number, up: boolean) =>
  Math.min(ZOOM_MAX, Math.max(ZOOM_MIN, up ? Math.floor(zoom / 10) * 10 + 10 : Math.ceil(zoom / 10) * 10 - 10));

const formatNumber = (n: number) => n.toLocaleString('pt-BR');

/** Barra de status do Word: "Página 1 de 3", palavras, idioma e zoom. */
export function StatusBar(props: {
  page: number;
  pages: number | null;
  words: number | null;
  zoom: number;
  onZoom: (zoom: number) => void;
}) {
  const { zoom, onZoom } = props;
  return (
    <footer className="statusbar">
      <span>{props.pages == null ? 'Carregando…' : `Página ${props.page} de ${props.pages}`}</span>
      {props.words != null && (
        <span>
          {formatNumber(props.words)} {props.words === 1 ? 'palavra' : 'palavras'}
        </span>
      )}
      <span>Português (Brasil)</span>
      <span className="spacer" />
      <div className="zoom-control">
        <button type="button" className="zoom-step" title="Reduzir" onClick={() => onZoom(stepZoom(zoom, false))}>
          <Minus size={12} />
        </button>
        <input
          type="range"
          className="zoom-slider"
          min={0}
          max={100}
          step={0.5}
          value={toSlider(zoom)}
          aria-label="Zoom"
          onChange={(e) => onZoom(fromSlider(Number(e.target.value)))}
        />
        <button type="button" className="zoom-step" title="Ampliar" onClick={() => onZoom(stepZoom(zoom, true))}>
          <Plus size={12} />
        </button>
        <button type="button" className="zoom-value" title="Nível de zoom: clique para voltar a 100%" onClick={() => onZoom(100)}>
          {Math.round(zoom)}%
        </button>
      </div>
    </footer>
  );
}
