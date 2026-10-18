import type { SuperDoc } from 'superdoc';
import { platform } from '../../platform';

/**
 * Ações da faixa que a barra "headless" do motor não cobre (área de transferência, subscrito,
 * estilos rápidos, quebra de página, link), feitas direto no editor ativo do SuperDoc.
 */
export interface EditorActions {
  focus(): void;
  copy(): void;
  cut(): void;
  paste(): void;
  selectAll(): void;
  insertPageBreak(): void;
  insertLink(): void;
  vertAlign(): string | null;
  toggleVertAlign(value: 'subscript' | 'superscript'): void;
  documentFonts(): string[];
  /** Fonte e tamanho padrão do documento (docDefaults + estilo Normal), para parágrafos vazios. */
  defaultFont(): { family: string | null; size: number | null };
  quickStyles(): QuickStyle[];
  /** Aplica um estilo pelo nome interno do Word ("heading 1", "Normal"). */
  applyStyleByName(name: string): void;
  wordCount(): number;
  /** Texto selecionado, se for um trecho curto de uma linha (o Word o põe na caixa "Localizar"). */
  selectedText(): string;
  /** Localizar e substituir, como no Word. Os resultados ficam realçados no documento. */
  find(query: string, options: { caseSensitive: boolean }): FindResult;
  findNext(backwards: boolean): FindResult;
  replaceCurrent(text: string): FindResult;
  replaceAll(text: string): number;
  clearFind(): void;
}

/** Posição na lista de ocorrências: "Resultado index+1 de count". */
export interface FindResult {
  index: number;
  count: number;
}

export interface QuickStyle {
  id: string;
  name: string;
  /** O objeto do motor, passado de volta ao aplicar o estilo. */
  source: Record<string, unknown>;
  /** Aparência resumida do estilo, para a amostra da galeria. */
  preview: React.CSSProperties;
}

/** Só o que usamos do editor do motor (a tipagem pública dele não expõe tudo). */
interface EngineEditor {
  commands: Record<string, (...args: unknown[]) => unknown>;
  state: {
    doc: { content: { size: number }; textBetween(from: number, to: number, block: string, leaf: string): string };
    selection?: { from: number; to: number; empty: boolean };
  };
  view?: { focus(): void };
  focus?: () => void;
  getAttributes(name: string): Record<string, unknown>;
  converter?: {
    linkedStyles?: EngineStyle[];
    getDocumentFonts?: () => string[];
    getDocumentDefaultStyles?: () => { typeface?: string; fontSizePt?: number };
  };
}

interface EngineStyle {
  id: string;
  type: string;
  definition?: { attrs?: Record<string, unknown>; styles?: Record<string, unknown> };
}

export type AskText = (title: string, label: string, initial?: string) => Promise<string | null>;

/** Primeiros estilos da galeria, na ordem do Word; os demais rápidos vêm depois em ordem alfabética. */
const QUICK_ORDER = ['normal', 'no spacing', 'heading 1', 'heading 2', 'heading 3', 'title', 'subtitle', 'quote', 'intense quote'];

function previewOf(styles: Record<string, unknown> | undefined): React.CSSProperties {
  if (!styles) return {};
  const pick = (k: string) => (typeof styles[k] === 'string' || typeof styles[k] === 'number' ? String(styles[k]) : undefined);
  const size = parseFloat(pick('font-size') ?? '');
  return {
    fontFamily: pick('font-family'),
    // A amostra da galeria do Word também limita o tamanho para caber no quadro.
    fontSize: Number.isFinite(size) ? `${Math.min(Math.max(size, 9), 16)}pt` : undefined,
    fontWeight: pick('font-weight'),
    fontStyle: pick('font-style'),
    color: pick('color'),
    textDecoration: pick('text-decoration'),
    letterSpacing: pick('letter-spacing'),
  };
}

export function createEditorActions(superdoc: SuperDoc | null, ask: AskText): EditorActions {
  const editor = () => (superdoc?.activeEditor ?? null) as unknown as EngineEditor | null;
  const cmd = (name: string, ...args: unknown[]) => editor()?.commands?.[name]?.(...args);
  // Ocorrências da última busca (o motor só devolve a contagem ao buscar e ao substituir).
  let findCount = 0;
  // Ir até a ocorrência põe o foco no documento; no Word ele fica na caixa de pesquisa ou no diálogo.
  const keepingFocus = <T,>(fn: () => T): T => {
    const had = document.activeElement as HTMLElement | null;
    const result = fn();
    if (had && had !== document.activeElement && had.closest('.find-dialog, .titlebar')) had.focus();
    return result;
  };

  const actions: EditorActions = {
    focus: () => {
      const e = editor();
      if (e?.focus) e.focus();
      else e?.view?.focus();
    },
    copy: () => {
      document.execCommand('copy');
    },
    cut: () => {
      document.execCommand('cut');
    },
    paste: () => {
      // O navegador não deixa um botão disparar o "colar" nativo; lemos a área de transferência
      // (pede permissão na primeira vez) e inserimos como o Ctrl+V faria.
      void (async () => {
        try {
          const items = await navigator.clipboard.read();
          for (const item of items) {
            if (item.types.includes('text/html')) {
              const html = await (await item.getType('text/html')).text();
              cmd('insertContent', html, { contentType: 'html' });
              return;
            }
          }
          const text = await navigator.clipboard.readText();
          if (text) cmd('insertContent', text);
        } catch {
          void platform.alert('Para colar, use Ctrl+V (o navegador não permitiu ler a área de transferência).');
        }
      })();
    },
    selectAll: () => {
      cmd('selectAll');
    },
    insertPageBreak: () => {
      cmd('insertPageBreak');
    },
    insertLink: () => {
      const current = editor()?.getAttributes('link')?.href;
      void ask('Inserir Hiperlink', 'Endereço:', typeof current === 'string' ? current : 'https://').then((href) => {
        if (href == null) return;
        if (href.trim()) cmd('setLink', { href: href.trim() });
        else cmd('unsetLink');
      });
    },
    vertAlign: () => {
      const v = editor()?.getAttributes('textStyle')?.vertAlign;
      return typeof v === 'string' ? v : null;
    },
    toggleVertAlign: (value) => {
      const e = editor();
      if (!e) return;
      const current = e.getAttributes('textStyle')?.vertAlign;
      cmd('setMark', 'textStyle', { vertAlign: current === value ? null : value });
    },
    documentFonts: () => {
      const fonts = editor()?.converter?.getDocumentFonts?.();
      return Array.isArray(fonts) ? fonts.filter((f) => typeof f === 'string' && f.trim()) : [];
    },
    defaultFont: () => {
      const d = editor()?.converter?.getDocumentDefaultStyles?.() ?? {};
      return { family: d.typeface ?? null, size: typeof d.fontSizePt === 'number' ? d.fontSizePt : null };
    },
    quickStyles: () => {
      const all = editor()?.converter?.linkedStyles ?? [];
      const quick = all.filter((s) => {
        if (s.type !== 'paragraph' || !s.definition?.attrs) return false;
        const name = String(s.definition.attrs.name ?? s.id).toLowerCase();
        return QUICK_ORDER.includes(name) || s.definition.attrs.qFormat === true;
      });
      const rank = (s: EngineStyle) => {
        const i = QUICK_ORDER.indexOf(String(s.definition?.attrs?.name ?? s.id).toLowerCase());
        return i < 0 ? QUICK_ORDER.length : i;
      };
      return quick
        .sort((a, b) => rank(a) - rank(b) || String(a.definition?.attrs?.name).localeCompare(String(b.definition?.attrs?.name)))
        .map((s) => ({
          id: s.id,
          name: String(s.definition?.attrs?.name ?? s.id),
          source: s as unknown as Record<string, unknown>,
          preview: previewOf(s.definition?.styles),
        }));
    },
    applyStyleByName(name) {
      const style = actions.quickStyles().find((s) => s.name.toLowerCase() === name.toLowerCase());
      if (style) cmd('setLinkedStyle', style.source);
    },
    selectedText: () => {
      const e = editor();
      const sel = e?.state?.selection;
      if (!e || !sel || sel.empty) return '';
      const text = e.state.doc.textBetween(sel.from, sel.to, '\n', ' ');
      return text.length <= 255 && !text.includes('\n') ? text : '';
    },
    find: (query, options) => keepingFocus(() => {
      const r = cmd('setSearchSession', query, { caseSensitive: options.caseSensitive, highlight: true }) as
        | { matches?: unknown[]; activeMatchIndex?: number }
        | undefined;
      const count = r?.matches?.length ?? 0;
      findCount = count;
      if (count) cmd('goToSearchResult', r!.matches![0]);
      return { index: count ? (r?.activeMatchIndex ?? 0) : -1, count };
    }),
    findNext: (backwards) => keepingFocus(() => {
      const r = cmd(backwards ? 'previousSearchMatch' : 'nextSearchMatch') as { activeMatchIndex?: number } | undefined;
      return { index: r?.activeMatchIndex ?? -1, count: findCount };
    }),
    replaceCurrent: (text) => keepingFocus(() => {
      const r = cmd('replaceSearchMatch', text) as { matches?: unknown[]; activeMatchIndex?: number } | undefined;
      findCount = r?.matches?.length ?? 0;
      return { index: r?.activeMatchIndex ?? -1, count: findCount };
    }),
    replaceAll: (text) => {
      const r = cmd('replaceAllSearchMatches', text) as { replacedCount?: number } | undefined;
      return r?.replacedCount ?? 0;
    },
    clearFind: () => {
      findCount = 0;
      cmd('clearSearchSession');
    },
    wordCount: () => {
      const doc = editor()?.state?.doc;
      if (!doc) return 0;
      const text = doc.textBetween(0, doc.content.size, ' ', ' ');
      return text.split(/\s+/).filter((w) => /[\p{L}\p{N}]/u.test(w)).length;
    },
  };
  return actions;
}
