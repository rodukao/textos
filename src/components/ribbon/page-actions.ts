import type { SuperDoc } from 'superdoc';

/**
 * Configurar Página e Parágrafo (guia Layout e régua), pela API de documento do motor
 * (`editor.doc`): seções em polegadas, recuos e espaçamentos em twips (1/20 pt).
 */

export const CM_PER_INCH = 2.54;
export const TWIPS_PER_INCH = 1440;
export const TWIPS_PER_PT = 20;

export interface Margins {
  top: number;
  right: number;
  bottom: number;
  left: number;
}

export interface SectionState {
  id: string;
  /** Largura e altura da folha, em polegadas. */
  width: number;
  height: number;
  orientation: 'portrait' | 'landscape';
  /** Margens em polegadas. */
  margins: Margins;
  columns: number;
}

/** Recuos e espaçamentos do parágrafo do cursor, já somando o que vem do estilo (twips). */
export interface ParagraphLayout {
  left: number;
  right: number;
  /** Recuo da primeira linha; negativo = deslocamento (hanging). */
  firstLine: number;
  before: number;
  after: number;
  /** Entrelinha: em 240 avos de linha (regra "auto") ou em twips ("exact"/"atLeast"). */
  line: number;
  lineRule: LineRule;
  /** "Não adicionar espaço entre parágrafos do mesmo estilo". */
  contextualSpacing: boolean;
  alignment: Alignment;
  /** Item de lista: o recuo anda por nível da lista, não pela tabulação. */
  isList: boolean;
}

export type LineRule = 'auto' | 'exact' | 'atLeast';
export type Alignment = 'left' | 'center' | 'right' | 'justify';

/** Entrelinha simples do Word quando nada a define (docDefaults sem w:spacing). */
const SINGLE_LINE = 240;

export interface PageActions {
  section(): SectionState | null;
  setMargins(margins: Partial<Margins>): void;
  setOrientation(orientation: 'portrait' | 'landscape'): void;
  setPaperSize(widthIn: number, heightIn: number): void;
  setColumns(count: number): void;
  insertSectionBreak(): void;
  paragraph(): ParagraphLayout | null;
  setIndentation(values: { left?: number; right?: number; firstLine?: number }): void;
  setSpacing(values: { before?: number; after?: number; line?: number; lineRule?: LineRule }): void;
  setAlignment(alignment: Alignment): void;
  setContextualSpacing(value: boolean): void;
  /** "Aumentar/Diminuir Recuo" do Word: até a próxima parada de tabulação padrão do documento. */
  stepIndent(direction: 1 | -1): void;
}

type Xml = { name?: string; attributes?: Record<string, string>; elements?: Xml[] };

interface PmNode {
  type: { name: string };
  attrs: Record<string, unknown>;
}

interface DocApiEditor {
  doc?: {
    sections: {
      list(): { items: SectionItem[] };
      setPageMargins(input: object): unknown;
      setPageSetup(input: object): unknown;
      setColumns(input: object): unknown;
    };
    selection: { current(): { target?: { segments?: { blockId: string }[] } } };
    format: {
      paragraph: {
        setIndentation(input: object): unknown;
        setSpacing(input: object): unknown;
        setAlignment(input: object): unknown;
        setFlowOptions(input: object): unknown;
      };
    };
  };
  state: {
    doc: { descendants(fn: (node: PmNode, pos: number) => boolean | void): void };
    selection: { from: number; $from: { depth: number; node(depth: number): PmNode } };
  };
  commands: Record<string, (...args: unknown[]) => unknown>;
  converter?: { convertedXml?: Record<string, Xml> };
}

interface SectionItem {
  address: { kind: 'section'; sectionId: string };
  range: { startParagraphIndex: number; endParagraphIndex: number };
  pageSetup?: { width?: number; height?: number; orientation?: 'portrait' | 'landscape' };
  margins?: Partial<Margins>;
  columns?: { count?: number };
}

const child = (el: Xml | undefined, name: string) => el?.elements?.find((e) => e.name === name);
const num = (v: string | undefined) => (v == null || v === '' ? undefined : Number(v));

/** Recuos/espaçamentos de um w:pPr (atributos em twips). */
function readPPr(pPr: Xml | undefined): Partial<ParagraphLayout> {
  const out: Partial<ParagraphLayout> = {};
  const spacing = child(pPr, 'w:spacing')?.attributes;
  const ind = child(pPr, 'w:ind')?.attributes;
  if (spacing) {
    if (num(spacing['w:before']) != null) out.before = num(spacing['w:before']);
    if (num(spacing['w:after']) != null) out.after = num(spacing['w:after']);
    if (num(spacing['w:line']) != null) {
      out.line = num(spacing['w:line']);
      out.lineRule = (spacing['w:lineRule'] as LineRule | undefined) ?? 'auto';
    }
  }
  const contextual = child(pPr, 'w:contextualSpacing');
  if (contextual) out.contextualSpacing = contextual.attributes?.['w:val'] !== '0' && contextual.attributes?.['w:val'] !== 'false';
  const jc = child(pPr, 'w:jc')?.attributes?.['w:val'];
  if (jc) out.alignment = jc === 'both' || jc === 'distribute' ? 'justify' : jc === 'end' ? 'right' : jc === 'start' ? 'left' : (jc as Alignment);
  if (ind) {
    const left = num(ind['w:left'] ?? ind['w:start']);
    const right = num(ind['w:right'] ?? ind['w:end']);
    if (left != null) out.left = left;
    if (right != null) out.right = right;
    if (num(ind['w:hanging']) != null) out.firstLine = -num(ind['w:hanging'])!;
    else if (num(ind['w:firstLine']) != null) out.firstLine = num(ind['w:firstLine']);
  }
  return out;
}

/** Valores do estilo (seguindo "baseado em") e dos padrões do documento, como o Word resolve. */
function styleLayout(editor: DocApiEditor, styleId: string | undefined): Partial<ParagraphLayout> {
  const root = editor.converter?.convertedXml?.['word/styles.xml']?.elements?.[0];
  const styles = root?.elements ?? [];
  const defaults = readPPr(child(child(child(root, 'w:docDefaults'), 'w:pPrDefault'), 'w:pPr'));
  const byId = (id: string) => styles.find((s) => s.name === 'w:style' && s.attributes?.['w:styleId'] === id);
  let style =
    (styleId && byId(styleId)) ||
    styles.find((s) => s.name === 'w:style' && s.attributes?.['w:type'] === 'paragraph' && s.attributes?.['w:default'] === '1');
  const chain: Partial<ParagraphLayout>[] = [];
  for (let guard = 0; style && guard < 20; guard++) {
    chain.unshift(readPPr(child(style, 'w:pPr')));
    const basedOn = child(style, 'w:basedOn')?.attributes?.['w:val'];
    style = basedOn ? byId(basedOn) : undefined;
  }
  return Object.assign({}, defaults, ...chain);
}

const inches = (v: number | undefined, fallback: number) => (typeof v === 'number' && Number.isFinite(v) ? v : fallback);

export function createPageActions(superdoc: SuperDoc | null): PageActions {
  const editor = () => (superdoc?.activeEditor ?? null) as unknown as DocApiEditor | null;

  /** Parágrafo do cursor e seu índice na ordem do documento (é o índice que as seções usam). */
  const cursorParagraph = (e: DocApiEditor) => {
    const { $from, from } = e.state.selection;
    let node: PmNode | null = null;
    for (let d = $from.depth; d >= 0; d--) {
      if ($from.node(d).type.name === 'paragraph') {
        node = $from.node(d);
        break;
      }
    }
    let index = -1;
    e.state.doc.descendants((n, pos) => {
      if (pos > from) return false;
      if (n.type.name === 'paragraph') index++;
    });
    return { node, index: Math.max(0, index) };
  };

  const currentSection = (e: DocApiEditor): SectionItem | null => {
    const items = e.doc?.sections.list().items ?? [];
    const { index } = cursorParagraph(e);
    return items.find((s) => index >= s.range.startParagraphIndex && index <= s.range.endParagraphIndex) ?? items[0] ?? null;
  };

  /** Parágrafos da seleção (ou o do cursor), como alvos da API de documento. */
  const selectedParagraphs = (e: DocApiEditor) => {
    const ids = new Set<string>();
    for (const s of e.doc?.selection.current().target?.segments ?? []) ids.add(s.blockId);
    if (!ids.size) {
      const id = cursorParagraph(e).node?.attrs.sdBlockId;
      if (typeof id === 'string') ids.add(id);
    }
    return [...ids].map((nodeId) => ({ kind: 'block', nodeType: 'paragraph', nodeId }));
  };

  const withSection = (fn: (e: DocApiEditor, target: SectionItem['address']) => void) => {
    const e = editor();
    const section = e && currentSection(e);
    if (e && section) fn(e, section.address);
  };

  // Cada alvo é aplicado separadamente: um parágrafo que já tem o valor devolve "sem mudança" e não
  // deve impedir os outros.
  const eachParagraph = (fn: (e: DocApiEditor, target: object) => void) => {
    const e = editor();
    if (!e) return;
    for (const target of selectedParagraphs(e)) {
      try {
        fn(e, target);
      } catch {
        // Bloco que não é parágrafo comum (ex.: dentro de um campo protegido): fica como está.
      }
    }
  };

  return {
    section: () => {
      const e = editor();
      const s = e && currentSection(e);
      if (!s) return null;
      const m = s.margins ?? {};
      return {
        id: s.address.sectionId,
        width: inches(s.pageSetup?.width, 8.5),
        height: inches(s.pageSetup?.height, 11),
        orientation: s.pageSetup?.orientation ?? 'portrait',
        margins: { top: inches(m.top, 1), right: inches(m.right, 1), bottom: inches(m.bottom, 1), left: inches(m.left, 1) },
        columns: s.columns?.count ?? 1,
      };
    },
    setMargins: (margins) => withSection((e, target) => e.doc!.sections.setPageMargins({ target, ...margins })),
    setOrientation: (orientation) =>
      withSection((e, target) => {
        const s = currentSection(e);
        const w = inches(s?.pageSetup?.width, 8.5);
        const h = inches(s?.pageSetup?.height, 11);
        // Como no Word, virar a folha troca largura e altura.
        const [width, height] = (orientation === 'landscape') === w > h ? [w, h] : [h, w];
        e.doc!.sections.setPageSetup({ target, orientation, width, height });
      }),
    setPaperSize: (widthIn, heightIn) =>
      withSection((e, target) => {
        const landscape = currentSection(e)?.pageSetup?.orientation === 'landscape';
        const [width, height] = landscape ? [Math.max(widthIn, heightIn), Math.min(widthIn, heightIn)] : [widthIn, heightIn];
        e.doc!.sections.setPageSetup({ target, width, height });
      }),
    setColumns: (count) => withSection((e, target) => e.doc!.sections.setColumns({ target, count })),
    insertSectionBreak: () => {
      editor()?.commands.insertSectionBreakAtSelection?.();
    },
    paragraph: () => {
      const e = editor();
      const node = e && cursorParagraph(e).node;
      if (!e || !node) return null;
      const props = (node.attrs.paragraphProperties ?? {}) as {
        styleId?: string;
        spacing?: { before?: number; after?: number; line?: number; lineRule?: LineRule };
        indent?: { left?: number; right?: number; firstLine?: number; hanging?: number };
        contextualSpacing?: boolean;
        justification?: string;
        numberingProperties?: unknown;
      };
      const fromStyle = styleLayout(e, props.styleId);
      const ind = props.indent ?? {};
      const firstLine = ind.hanging != null ? -ind.hanging : ind.firstLine;
      return {
        left: ind.left ?? fromStyle.left ?? 0,
        right: ind.right ?? fromStyle.right ?? 0,
        firstLine: firstLine ?? fromStyle.firstLine ?? 0,
        before: props.spacing?.before ?? fromStyle.before ?? 0,
        after: props.spacing?.after ?? fromStyle.after ?? 0,
        line: props.spacing?.line ?? fromStyle.line ?? SINGLE_LINE,
        lineRule: props.spacing?.line != null ? (props.spacing.lineRule ?? 'auto') : (fromStyle.lineRule ?? 'auto'),
        contextualSpacing: props.contextualSpacing ?? fromStyle.contextualSpacing ?? false,
        alignment: (props.justification === 'both' ? 'justify' : (props.justification as Alignment | undefined)) ?? fromStyle.alignment ?? 'left',
        isList: Boolean(props.numberingProperties),
      };
    },
    stepIndent(direction) {
      const e = editor();
      const current = this.paragraph();
      if (!e || !current) return;
      // Parada de tabulação padrão (w:defaultTabStop; 1,25 cm no Word em português).
      const settings = e.converter?.convertedXml?.['word/settings.xml'];
      const root = settings?.name === 'w:settings' ? settings : settings?.elements?.find((x) => x.name === 'w:settings');
      const tab = num(child(root, 'w:defaultTabStop')?.attributes?.['w:val']) || 709;
      const left = current.left;
      const next = direction > 0 ? (Math.floor(left / tab + 1e-6) + 1) * tab : Math.max(0, (Math.ceil(left / tab - 1e-6) - 1) * tab);
      this.setIndentation({ left: next });
    },
    setIndentation: ({ left, right, firstLine }) =>
      eachParagraph((e, target) => {
        const input: Record<string, unknown> = { target };
        if (left != null) input.left = Math.round(left);
        if (right != null) input.right = Math.round(right);
        if (firstLine != null) {
          if (firstLine < 0) input.hanging = Math.round(-firstLine);
          else input.firstLine = Math.round(firstLine);
        }
        e.doc!.format.paragraph.setIndentation(input);
      }),
    setSpacing: ({ before, after, line, lineRule }) =>
      eachParagraph((e, target) => {
        const input: Record<string, unknown> = { target };
        if (before != null) input.before = Math.max(0, Math.round(before));
        if (after != null) input.after = Math.max(0, Math.round(after));
        if (line != null) {
          input.line = Math.max(1, Math.round(line));
          input.lineRule = lineRule ?? 'auto';
        }
        e.doc!.format.paragraph.setSpacing(input);
      }),
    setAlignment: (alignment) => eachParagraph((e, target) => e.doc!.format.paragraph.setAlignment({ target, alignment })),
    setContextualSpacing: (value) =>
      eachParagraph((e, target) => e.doc!.format.paragraph.setFlowOptions({ target, contextualSpacing: value })),
  };
}
