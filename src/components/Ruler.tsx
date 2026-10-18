import { useEffect, useState } from 'react';
import { CM_PER_INCH, TWIPS_PER_INCH, type PageActions } from './ribbon/page-actions';
import { formatDecimal } from './ribbon/word-data';

/**
 * Régua horizontal do Word, em centímetros: margens em cinza, área do texto em branco, números
 * contados a partir da margem esquerda e os marcadores de recuo (primeira linha, deslocamento,
 * recuo à esquerda e à direita). Arrastar a divisa cinza/branco muda a margem; arrastar um
 * marcador muda o recuo dos parágrafos selecionados. Alt ao arrastar desliga o ajuste à grade.
 */

type Handle = 'marginLeft' | 'marginRight' | 'firstLine' | 'hanging' | 'leftIndent' | 'rightIndent';

interface Drag {
  handle: Handle;
  /** Posição atual, em polegadas a partir da borda esquerda da folha. */
  x: number;
}

/** Ajuste à grade do Word ao arrastar na régua: 0,25 cm. */
const SNAP_IN = 0.25 / CM_PER_INCH;
const twipsToIn = (t: number) => t / TWIPS_PER_INCH;
const inToTwips = (i: number) => i * TWIPS_PER_INCH;

const HANDLE_TITLES: Record<Handle, string> = {
  marginLeft: 'Margem Esquerda',
  marginRight: 'Margem Direita',
  firstLine: 'Recuo da Primeira Linha',
  hanging: 'Recuo Deslocado',
  leftIndent: 'Recuo à Esquerda',
  rightIndent: 'Recuo à Direita',
};

export function Ruler(props: {
  page: PageActions;
  workspace: HTMLElement | null;
  pageIndex: number;
  /** Muda quando a seleção ou o documento mudam (snapshot do motor). */
  revision: unknown;
}) {
  const { page, workspace } = props;
  // Redesenha ao rolar, redimensionar ou mudar o zoom (a folha muda de lugar e de tamanho).
  const [, setFrame] = useState(0);
  useEffect(() => {
    if (!workspace) return;
    const bump = () => setFrame((n) => n + 1);
    const observer = new ResizeObserver(bump);
    observer.observe(workspace);
    workspace.addEventListener('scroll', bump, { passive: true });
    window.addEventListener('resize', bump);
    return () => {
      observer.disconnect();
      workspace.removeEventListener('scroll', bump);
      window.removeEventListener('resize', bump);
    };
  }, [workspace]);

  const [drag, setDrag] = useState<Drag | null>(null);

  const section = page.section();
  const paragraph = page.paragraph();
  const pages = workspace?.querySelectorAll<HTMLElement>('.superdoc-page');
  const sheet = pages?.[Math.min(props.pageIndex, pages.length) - 1] ?? pages?.[0];
  if (!section || !sheet || !workspace) return <div className="ruler-bar" />;

  const bar = workspace.getBoundingClientRect();
  const rect = sheet.getBoundingClientRect();
  const pxPerIn = rect.width / section.width;
  const W = section.width;
  const m = section.margins;

  // Posições (polegadas a partir da borda da folha), já com o arrasto em andamento.
  const at = (h: Handle, base: number) => (drag?.handle === h ? drag.x : base);
  const marginLeft = at('marginLeft', m.left);
  const marginRight = at('marginRight', W - m.right);
  const p = paragraph ?? { left: 0, right: 0, firstLine: 0, before: 0, after: 0, isList: false };
  let leftIndent = marginLeft + twipsToIn(p.left);
  let firstLine = leftIndent + twipsToIn(p.firstLine);
  if (drag?.handle === 'leftIndent') {
    const delta = drag.x - leftIndent;
    leftIndent += delta;
    firstLine += delta;
  } else if (drag?.handle === 'hanging') {
    leftIndent = drag.x;
  } else if (drag?.handle === 'firstLine') {
    firstLine = drag.x;
  }
  const rightIndent = at('rightIndent', marginRight - twipsToIn(p.right));

  const toPx = (inches: number) => inches * pxPerIn;

  const startDrag = (handle: Handle, base: number) => (e: React.PointerEvent) => {
    e.preventDefault();
    e.stopPropagation();
    (e.target as Element).setPointerCapture(e.pointerId);
    const offset = (e.clientX - rect.left) / pxPerIn - base;
    setDrag({ handle, x: base });
    const move = (ev: PointerEvent) => {
      let x = (ev.clientX - rect.left) / pxPerIn - offset;
      if (!ev.altKey) x = Math.round((x - m.left) / SNAP_IN) * SNAP_IN + m.left;
      setDrag({ handle, x: Math.min(W, Math.max(0, x)) });
    };
    const up = (ev: PointerEvent) => {
      window.removeEventListener('pointermove', move);
      window.removeEventListener('pointerup', up);
      let x = (ev.clientX - rect.left) / pxPerIn - offset;
      if (!ev.altKey) x = Math.round((x - m.left) / SNAP_IN) * SNAP_IN + m.left;
      x = Math.min(W, Math.max(0, x));
      setDrag(null);
      commit(handle, x);
    };
    window.addEventListener('pointermove', move);
    window.addEventListener('pointerup', up);
  };

  const commit = (handle: Handle, x: number) => {
    const minText = 0.5; // a área do texto não some
    switch (handle) {
      case 'marginLeft':
        page.setMargins({ left: Math.min(x, W - m.right - minText) });
        break;
      case 'marginRight':
        page.setMargins({ right: Math.min(W - x, W - m.left - minText) });
        break;
      case 'leftIndent': {
        page.setIndentation({ left: inToTwips(x - m.left), firstLine: p.firstLine });
        break;
      }
      case 'hanging': {
        // Como no Word: a primeira linha fica onde estava e o resto do parágrafo se move.
        const firstAbs = m.left + twipsToIn(p.left + p.firstLine);
        page.setIndentation({ left: inToTwips(x - m.left), firstLine: inToTwips(firstAbs - x) });
        break;
      }
      case 'firstLine':
        page.setIndentation({ firstLine: inToTwips(x - (m.left + twipsToIn(p.left))) });
        break;
      case 'rightIndent':
        page.setIndentation({ right: inToTwips(W - m.right - x) });
        break;
    }
  };

  // Marcas a cada 0,25 cm, contadas da margem esquerda para os dois lados (como no Word).
  const quarter = 0.25 / CM_PER_INCH;
  const ticks: { x: number; kind: 'number' | 'half' | 'quarter'; label?: string }[] = [];
  for (let i = -Math.floor(marginLeft / quarter); marginLeft + i * quarter <= W + 1e-6; i++) {
    if (i === 0) continue;
    const x = marginLeft + i * quarter;
    const kind = i % 4 === 0 ? 'number' : i % 2 === 0 ? 'half' : 'quarter';
    ticks.push({ x, kind, label: kind === 'number' ? String(Math.abs(i / 4)) : undefined });
  }

  const guide = drag ? toPx(drag.x) : null;
  const dragLabel =
    drag &&
    (drag.handle === 'marginLeft' || drag.handle === 'leftIndent' || drag.handle === 'hanging' || drag.handle === 'firstLine'
      ? `${HANDLE_TITLES[drag.handle]}: ${formatDecimal((drag.x - (drag.handle === 'marginLeft' ? 0 : m.left)) * CM_PER_INCH)} cm`
      : `${HANDLE_TITLES[drag.handle]}: ${formatDecimal((W - drag.x - (drag.handle === 'marginRight' ? 0 : m.right)) * CM_PER_INCH)} cm`);

  return (
    <div className="ruler-bar" aria-label="Régua">
      <div className="ruler" style={{ left: rect.left - bar.left, width: rect.width }}>
        <div className="ruler-margin" style={{ left: 0, width: toPx(marginLeft) }} />
        <div className="ruler-margin" style={{ left: toPx(marginRight), right: 0 }} />
        {ticks.map((t) =>
          t.kind === 'number' ? (
            <span key={t.x} className="ruler-number" style={{ left: toPx(t.x) }}>
              {t.label}
            </span>
          ) : (
            <span key={t.x} className={`ruler-tick ruler-tick-${t.kind}`} style={{ left: toPx(t.x) }} />
          ),
        )}
        <span
          className="ruler-edge"
          style={{ left: toPx(marginLeft) }}
          title={HANDLE_TITLES.marginLeft}
          onPointerDown={startDrag('marginLeft', marginLeft)}
        />
        <span
          className="ruler-edge"
          style={{ left: toPx(marginRight) }}
          title={HANDLE_TITLES.marginRight}
          onPointerDown={startDrag('marginRight', marginRight)}
        />
        {paragraph && (
          <>
            <span
              className="ruler-marker ruler-first"
              style={{ left: toPx(firstLine) }}
              title={HANDLE_TITLES.firstLine}
              onPointerDown={startDrag('firstLine', firstLine)}
            />
            <span
              className="ruler-marker ruler-hanging"
              style={{ left: toPx(leftIndent) }}
              title={HANDLE_TITLES.hanging}
              onPointerDown={startDrag('hanging', leftIndent)}
            />
            <span
              className="ruler-marker ruler-left"
              style={{ left: toPx(leftIndent) }}
              title={HANDLE_TITLES.leftIndent}
              onPointerDown={startDrag('leftIndent', leftIndent)}
            />
            <span
              className="ruler-marker ruler-right"
              style={{ left: toPx(rightIndent) }}
              title={HANDLE_TITLES.rightIndent}
              onPointerDown={startDrag('rightIndent', rightIndent)}
            />
          </>
        )}
        {guide != null && (
          <>
            <span className="ruler-guide" style={{ left: guide }} />
            <span className="ruler-drag-label" style={{ left: guide }}>
              {dragLabel}
            </span>
          </>
        )}
      </div>
    </div>
  );
}
