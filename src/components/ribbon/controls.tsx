import { useEffect, useRef, useState, type ReactNode } from 'react';
import { ChevronDown, type LucideIcon } from 'lucide-react';

/**
 * Peças da faixa de opções, no desenho do Word: grupos com o nome embaixo, botões grandes
 * (ícone sobre o texto) e pequenos, botões divididos com menu e caixas de combinação.
 *
 * Os botões não tiram o foco do documento (mousedown sem efeito padrão): a seleção continua
 * onde estava e o comando vale para ela, como no Word.
 */

const keepFocus = (e: React.MouseEvent) => e.preventDefault();

/**
 * Posição do menu na janela, logo abaixo do controle: a faixa rola na horizontal em telas
 * estreitas e cortaria um menu posicionado dentro dela.
 */
export function popupPosition(anchor: HTMLElement | null, width = 320): React.CSSProperties {
  if (!anchor) return {};
  const r = anchor.getBoundingClientRect();
  return { position: 'fixed', top: r.bottom + 2, left: Math.max(4, Math.min(r.left, window.innerWidth - width - 4)) };
}

export function RibbonGroup(props: { label: string; children: ReactNode; className?: string }) {
  return (
    <section className={`ribbon-group ${props.className ?? ''}`} aria-label={props.label}>
      <div className="ribbon-group-body">{props.children}</div>
      <div className="ribbon-group-label">{props.label}</div>
    </section>
  );
}

/** Coluna de linhas de botões pequenos dentro de um grupo. */
export function Rows(props: { children: ReactNode }) {
  return <div className="ribbon-rows">{props.children}</div>;
}

export function Row(props: { children: ReactNode }) {
  return <div className="ribbon-row">{props.children}</div>;
}

export interface ButtonProps {
  icon?: LucideIcon;
  /** Conteúdo próprio no lugar do ícone (ex.: o "A" com a barra de cor). */
  glyph?: ReactNode;
  label?: string;
  /** Dica (nome e atalho), como a do Word. */
  title: string;
  active?: boolean;
  disabled?: boolean;
  large?: boolean;
  onClick: () => void;
}

export function RibbonButton(props: ButtonProps) {
  const Icon = props.icon;
  return (
    <button
      type="button"
      className={`rb ${props.large ? 'rb-large' : 'rb-small'}${props.active ? ' is-active' : ''}`}
      title={props.title}
      aria-label={props.label ?? props.title}
      aria-pressed={props.active}
      disabled={props.disabled}
      onMouseDown={keepFocus}
      onClick={props.onClick}
    >
      {props.glyph ?? (Icon && <Icon size={props.large ? 30 : 16} strokeWidth={props.large ? 1.4 : 1.8} />)}
      {props.label && <span className="rb-label">{props.label}</span>}
    </button>
  );
}

/** Fecha o menu ao clicar fora ou apertar Esc. */
function useDismiss(open: boolean, close: () => void) {
  const ref = useRef<HTMLDivElement>(null);
  useEffect(() => {
    if (!open) return;
    const onDown = (e: MouseEvent) => {
      if (!ref.current?.contains(e.target as Node)) close();
    };
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') close();
    };
    document.addEventListener('mousedown', onDown, true);
    document.addEventListener('keydown', onKey, true);
    return () => {
      document.removeEventListener('mousedown', onDown, true);
      document.removeEventListener('keydown', onKey, true);
    };
  }, [open, close]);
  return ref;
}

/**
 * Botão com menu. Com `onClick`, é dividido como no Word: a parte principal repete a última
 * escolha (ex.: a cor do realce) e a seta abre o menu.
 */
export function MenuButton(
  props: Omit<ButtonProps, 'onClick'> & {
    menu: (close: () => void) => ReactNode;
    onClick?: () => void;
  },
) {
  const [open, setOpen] = useState(false);
  const close = () => setOpen(false);
  const ref = useDismiss(open, close);
  const toggle = () => setOpen(!open);
  const Icon = props.icon;
  const face = props.glyph ?? (Icon && <Icon size={props.large ? 30 : 16} strokeWidth={props.large ? 1.4 : 1.8} />);
  const split = typeof props.onClick === 'function';
  return (
    <div className={`rb-menu${split ? ' is-split' : ''}${props.large ? ' rb-menu-large' : ''}`} ref={ref}>
      {split ? (
        <>
          <button
            type="button"
            className={`rb rb-small rb-main${props.active ? ' is-active' : ''}`}
            title={props.title}
            aria-label={props.label ?? props.title}
            disabled={props.disabled}
            onMouseDown={keepFocus}
            onClick={props.onClick}
          >
            {face}
            {props.label && <span className="rb-label">{props.label}</span>}
          </button>
          <button
            type="button"
            className="rb rb-small rb-arrow"
            title={`${props.title}: mais opções`}
            aria-haspopup="menu"
            aria-expanded={open}
            disabled={props.disabled}
            onMouseDown={keepFocus}
            onClick={toggle}
          >
            <ChevronDown size={11} />
          </button>
        </>
      ) : (
        <button
          type="button"
          className={`rb ${props.large ? 'rb-large' : 'rb-small'}${props.active ? ' is-active' : ''}`}
          title={props.title}
          aria-haspopup="menu"
          aria-expanded={open}
          disabled={props.disabled}
          onMouseDown={keepFocus}
          onClick={toggle}
        >
          {face}
          {props.label && <span className="rb-label">{props.label}</span>}
          <ChevronDown size={11} className="rb-caret" />
        </button>
      )}
      {open && (
        <div className="rb-popup" role="menu" style={popupPosition(ref.current)}>
          {props.menu(close)}
        </div>
      )}
    </div>
  );
}

export function MenuItem(props: {
  label: string;
  checked?: boolean;
  onSelect: () => void;
  close: () => void;
  style?: React.CSSProperties;
}) {
  return (
    <button
      type="button"
      role="menuitemradio"
      aria-checked={props.checked}
      className={`menu-item${props.checked ? ' is-checked' : ''}`}
      style={props.style}
      onMouseDown={keepFocus}
      onClick={() => {
        props.close();
        props.onSelect();
      }}
    >
      {props.label}
    </button>
  );
}

/**
 * Caixa de combinação do Word (fonte, tamanho): dá para digitar e confirmar com Enter, ou
 * escolher na lista.
 */
export function ComboBox(props: {
  value: string;
  options: { label: string; value: string; style?: React.CSSProperties }[];
  onCommit: (value: string) => void;
  title: string;
  width: number;
}) {
  const [text, setText] = useState(props.value);
  const [open, setOpen] = useState(false);
  const close = () => setOpen(false);
  const ref = useDismiss(open, close);
  const listRef = useRef<HTMLDivElement>(null);
  useEffect(() => setText(props.value), [props.value]);
  useEffect(() => {
    if (!open) return;
    listRef.current?.querySelector('.is-checked')?.scrollIntoView({ block: 'center' });
  }, [open]);
  const commit = (value: string) => {
    const v = value.trim();
    if (v) props.onCommit(v);
    else setText(props.value);
  };
  return (
    <div className="combo" ref={ref} style={{ width: props.width }}>
      <input
        className="combo-input"
        value={text}
        title={props.title}
        aria-label={props.title}
        spellCheck={false}
        onChange={(e) => setText(e.target.value)}
        onFocus={(e) => e.target.select()}
        onKeyDown={(e) => {
          if (e.key === 'Enter') {
            e.preventDefault();
            commit(text);
          } else if (e.key === 'Escape') {
            setText(props.value);
          }
        }}
      />
      <button
        type="button"
        className="combo-arrow"
        title={props.title}
        aria-haspopup="listbox"
        aria-expanded={open}
        onMouseDown={keepFocus}
        onClick={() => setOpen(!open)}
      >
        <ChevronDown size={11} />
      </button>
      {open && (
        <div className="rb-popup combo-list" role="listbox" ref={listRef} style={popupPosition(ref.current, 220)}>
          {props.options.map((o) => (
            <MenuItem
              key={o.value}
              label={o.label}
              style={o.style}
              checked={o.label.toLowerCase() === props.value.toLowerCase()}
              onSelect={() => commit(o.value)}
              close={close}
            />
          ))}
        </div>
      )}
    </div>
  );
}

/** Grade de cores, como a "Cores do Tema" + "Cores Padrão" do Word. */
export function ColorGrid(props: {
  rows: string[][];
  standard?: string[];
  noneLabel: string;
  onPick: (color: string | null) => void;
  close: () => void;
}) {
  const swatch = (c: string) => (
    <button
      key={c}
      type="button"
      className="swatch"
      style={{ background: c }}
      title={c}
      onMouseDown={keepFocus}
      onClick={() => {
        props.close();
        props.onPick(c);
      }}
    />
  );
  return (
    <div className="color-grid">
      <button
        type="button"
        className="menu-item"
        onMouseDown={keepFocus}
        onClick={() => {
          props.close();
          props.onPick(null);
        }}
      >
        {props.noneLabel}
      </button>
      {props.rows.map((row, i) => (
        <div key={i} className={`swatch-row${i === 0 ? ' first' : ''}`}>
          {row.map(swatch)}
        </div>
      ))}
      {props.standard && (
        <>
          <div className="color-grid-title">Cores Padrão</div>
          <div className="swatch-row">{props.standard.map(swatch)}</div>
        </>
      )}
    </div>
  );
}

/** A letra (ou ícone) com a barra da cor atual embaixo: "A" da cor da fonte, marcador do realce. */
export function ColorGlyph(props: { children: ReactNode; color: string }) {
  return (
    <span className="color-glyph">
      {props.children}
      <span className="color-bar" style={{ background: props.color }} />
    </span>
  );
}
