import { useEffect, useRef, useState } from 'react';
import {
  CM_PER_INCH,
  TWIPS_PER_INCH,
  TWIPS_PER_PT,
  type Alignment,
  type LineRule,
  type PageActions,
  type ParagraphLayout,
} from './ribbon/page-actions';
import { formatDecimal, parseDecimal } from './ribbon/word-data';

/**
 * Caixa "Parágrafo" do Word (guia Recuos e espaçamento): alinhamento, recuos com "Especial"
 * (primeira linha ou deslocamento), espaçamento antes/depois e entre linhas (Simples, 1,5 linha,
 * Duplo, Pelo menos, Exatamente, Múltiplos), e "Não adicionar espaço entre parágrafos do mesmo
 * estilo". Vale para todos os parágrafos selecionados.
 */

type LineKind = 'single' | 'onePointFive' | 'double' | 'atLeast' | 'exact' | 'multiple';

const LINE_KINDS: { value: LineKind; label: string }[] = [
  { value: 'single', label: 'Simples' },
  { value: 'onePointFive', label: '1,5 linha' },
  { value: 'double', label: 'Duplo' },
  { value: 'atLeast', label: 'Pelo menos' },
  { value: 'exact', label: 'Exatamente' },
  { value: 'multiple', label: 'Múltiplos' },
];

const ALIGNMENTS: { value: Alignment; label: string }[] = [
  { value: 'left', label: 'Esquerda' },
  { value: 'center', label: 'Centralizado' },
  { value: 'right', label: 'Direita' },
  { value: 'justify', label: 'Justificado' },
];

const twipsToCm = (t: number) => (t / TWIPS_PER_INCH) * CM_PER_INCH;
const cmToTwips = (cm: number) => (cm / CM_PER_INCH) * TWIPS_PER_INCH;
const cmText = (t: number) => `${formatDecimal(twipsToCm(t))} cm`;
const ptText = (t: number) => `${formatDecimal(t / TWIPS_PER_PT, 1)} pt`;

/** Como o Word mostra a entrelinha guardada no documento. */
function lineKindOf(p: ParagraphLayout): { kind: LineKind; at: string } {
  if (p.lineRule === 'exact') return { kind: 'exact', at: ptText(p.line) };
  if (p.lineRule === 'atLeast') return { kind: 'atLeast', at: ptText(p.line) };
  if (p.line === 240) return { kind: 'single', at: '' };
  if (p.line === 360) return { kind: 'onePointFive', at: '' };
  if (p.line === 480) return { kind: 'double', at: '' };
  return { kind: 'multiple', at: formatDecimal(p.line / 240) };
}

export function ParagraphDialog(props: { page: PageActions; onClose: () => void }) {
  const { page } = props;
  const initial = useRef(page.paragraph()).current;
  const firstRef = useRef<HTMLSelectElement>(null);
  useEffect(() => firstRef.current?.focus(), []);

  const p = initial ?? {
    left: 0,
    right: 0,
    firstLine: 0,
    before: 0,
    after: 0,
    line: 240,
    lineRule: 'auto' as LineRule,
    contextualSpacing: false,
    alignment: 'left' as Alignment,
    isList: false,
  };
  const startLine = lineKindOf(p);
  const [alignment, setAlignment] = useState<Alignment>(p.alignment);
  const [left, setLeft] = useState(cmText(p.left));
  const [right, setRight] = useState(cmText(p.right));
  const [special, setSpecial] = useState<'none' | 'firstLine' | 'hanging'>(
    p.firstLine > 0 ? 'firstLine' : p.firstLine < 0 ? 'hanging' : 'none',
  );
  const [by, setBy] = useState(p.firstLine ? cmText(Math.abs(p.firstLine)) : '');
  const [before, setBefore] = useState(ptText(p.before));
  const [after, setAfter] = useState(ptText(p.after));
  const [lineKind, setLineKind] = useState<LineKind>(startLine.kind);
  const [lineAt, setLineAt] = useState(startLine.at);
  const [contextual, setContextual] = useState(p.contextualSpacing);

  const changeLineKind = (kind: LineKind) => {
    setLineKind(kind);
    // Como no Word, a caixa "Em" ganha um valor padrão ao trocar o tipo.
    if (kind === 'multiple') setLineAt(lineKind === 'multiple' ? lineAt : '3');
    else if (kind === 'atLeast' || kind === 'exact') setLineAt(lineKind === kind ? lineAt : '12 pt');
    else setLineAt('');
  };

  const apply = () => {
    const num = (text: string) => parseDecimal(text) ?? 0;
    // Como no Word: só o que mudou vira formatação direta; o resto continua vindo do estilo.
    const changed = (a: number, b: number) => Math.abs(a - b) > 1;
    const newLeft = cmToTwips(num(left));
    const newRight = cmToTwips(num(right));
    const newFirst = special === 'none' ? 0 : (special === 'hanging' ? -1 : 1) * cmToTwips(Math.abs(num(by)));
    const indent: { left?: number; right?: number; firstLine?: number } = {};
    if (changed(newLeft, p.left)) indent.left = newLeft;
    if (changed(newRight, p.right)) indent.right = newRight;
    if (changed(newFirst, p.firstLine)) indent.firstLine = newFirst;
    if (Object.keys(indent).length) page.setIndentation(indent);

    let line = 240;
    let lineRule: LineRule = 'auto';
    if (lineKind === 'onePointFive') line = 360;
    else if (lineKind === 'double') line = 480;
    else if (lineKind === 'multiple') line = Math.max(0.06, num(lineAt)) * 240;
    else if (lineKind === 'atLeast' || lineKind === 'exact') {
      line = Math.max(0.7, num(lineAt)) * TWIPS_PER_PT;
      lineRule = lineKind;
    }
    const newBefore = Math.max(0, num(before)) * TWIPS_PER_PT;
    const newAfter = Math.max(0, num(after)) * TWIPS_PER_PT;
    const spacing: { before?: number; after?: number; line?: number; lineRule?: LineRule } = {};
    if (changed(newBefore, p.before)) spacing.before = newBefore;
    if (changed(newAfter, p.after)) spacing.after = newAfter;
    if (changed(line, p.line) || lineRule !== p.lineRule) Object.assign(spacing, { line, lineRule });
    if (Object.keys(spacing).length) page.setSpacing(spacing);

    if (alignment !== p.alignment) page.setAlignment(alignment);
    if (contextual !== p.contextualSpacing) page.setContextualSpacing(contextual);
    props.onClose();
  };

  const lineAtEnabled = lineKind === 'multiple' || lineKind === 'atLeast' || lineKind === 'exact';

  return (
    <div className="dialog-backdrop" onMouseDown={props.onClose}>
      <form
        className="dialog paragraph-dialog"
        role="dialog"
        aria-label="Parágrafo"
        onMouseDown={(e) => e.stopPropagation()}
        onSubmit={(e) => {
          e.preventDefault();
          apply();
        }}
        onKeyDown={(e) => {
          if (e.key === 'Escape') props.onClose();
        }}
      >
        <div className="dialog-title">Parágrafo</div>
        <div className="dialog-tab">Recuos e espaçamento</div>

        <fieldset className="dialog-section">
          <legend>Geral</legend>
          <label className="pd-field">
            <span>Alinhamento:</span>
            <select ref={firstRef} value={alignment} onChange={(e) => setAlignment(e.target.value as Alignment)}>
              {ALIGNMENTS.map((a) => (
                <option key={a.value} value={a.value}>
                  {a.label}
                </option>
              ))}
            </select>
          </label>
        </fieldset>

        <fieldset className="dialog-section pd-grid">
          <legend>Recuo</legend>
          <label className="pd-field">
            <span>Esquerdo:</span>
            <input value={left} onChange={(e) => setLeft(e.target.value)} />
          </label>
          <label className="pd-field">
            <span>Especial:</span>
            <select
              value={special}
              onChange={(e) => {
                const v = e.target.value as typeof special;
                setSpecial(v);
                if (v === 'none') setBy('');
                else if (!by) setBy('1,25 cm');
              }}
            >
              <option value="none">(nenhum)</option>
              <option value="firstLine">Primeira linha</option>
              <option value="hanging">Deslocamento</option>
            </select>
          </label>
          <label className="pd-field">
            <span>Direito:</span>
            <input value={right} onChange={(e) => setRight(e.target.value)} />
          </label>
          <label className="pd-field">
            <span>Por:</span>
            <input value={by} disabled={special === 'none'} onChange={(e) => setBy(e.target.value)} />
          </label>
        </fieldset>

        <fieldset className="dialog-section pd-grid">
          <legend>Espaçamento</legend>
          <label className="pd-field">
            <span>Antes:</span>
            <input value={before} onChange={(e) => setBefore(e.target.value)} />
          </label>
          <label className="pd-field">
            <span>Entre linhas:</span>
            <select value={lineKind} onChange={(e) => changeLineKind(e.target.value as LineKind)}>
              {LINE_KINDS.map((k) => (
                <option key={k.value} value={k.value}>
                  {k.label}
                </option>
              ))}
            </select>
          </label>
          <label className="pd-field">
            <span>Depois:</span>
            <input value={after} onChange={(e) => setAfter(e.target.value)} />
          </label>
          <label className="pd-field">
            <span>Em:</span>
            <input value={lineAt} disabled={!lineAtEnabled} onChange={(e) => setLineAt(e.target.value)} />
          </label>
          <label className="pd-check">
            <input type="checkbox" checked={contextual} onChange={(e) => setContextual(e.target.checked)} />
            Não adicionar espaço entre parágrafos do mesmo estilo
          </label>
        </fieldset>

        <div className="dialog-buttons">
          <button type="submit" className="dialog-ok">
            OK
          </button>
          <button type="button" onClick={props.onClose}>
            Cancelar
          </button>
        </div>
      </form>
    </div>
  );
}
