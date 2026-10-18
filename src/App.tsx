import { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react';
import type { PublicToolbarItemId } from 'superdoc/headless-toolbar';
import { useHeadlessToolbar } from 'superdoc/headless-toolbar/react';
import { Backstage } from './components/Backstage';
import type { FileActions } from './components/file-actions';
import { FindReplace, type FindTab } from './components/FindReplace';
import { PromptDialog, type PromptRequest } from './components/PromptDialog';
import { Ribbon } from './components/ribbon/Ribbon';
import { createEditorActions } from './components/ribbon/editor-actions';
import { createPageActions } from './components/ribbon/page-actions';
import { Ruler } from './components/Ruler';
import { stepFontSize } from './components/ribbon/word-data';
import { StatusBar, ZOOM_MAX, ZOOM_MIN } from './components/StatusBar';
import { TitleBar } from './components/TitleBar';
import { DOCX_FILTER, baseName } from './document/format';
import { DocumentEditor, type EditorHandle } from './editor/DocumentEditor';
import { DOCX_MIME } from './engine/superdoc-config';
import { platform, type FileRef } from './platform';
import { matchShortcut, type ShortcutAction } from './shortcuts';

const UNTITLED = 'Documento sem título';

/** Comandos do motor cujo estado (ativo, desativado, valor) a faixa e a barra de título mostram. */
const RIBBON_COMMANDS: PublicToolbarItemId[] = [
  'bold',
  'italic',
  'underline',
  'strikethrough',
  'font-size',
  'font-family',
  'text-color',
  'highlight-color',
  'link',
  'text-align',
  'line-height',
  'linked-style',
  'bullet-list',
  'numbered-list',
  'indent-increase',
  'indent-decrease',
  'undo',
  'redo',
  'ruler',
  'formatting-marks',
  'zoom',
  'zoom-fit-width',
  'document-mode',
  'clear-formatting',
  'copy-format',
  'track-changes-accept-selection',
  'track-changes-reject-selection',
  'image',
  'table-of-contents-insert',
  'table-insert',
];
const errorMessage = (e: unknown) => (e instanceof Error ? e.message : String(e));

interface DocState {
  /** Muda a cada documento novo/aberto e remonta o DocumentEditor. */
  key: number;
  file: File | null;
  name: string;
  ref: FileRef | null;
}

/** Número (1…) da folha que ocupa o alto da área visível, como a "Página X de Y" do Word. */
function visiblePage(workspace: HTMLElement | null): number {
  if (!workspace) return 1;
  const pages = workspace.querySelectorAll('.superdoc-page');
  const probe = workspace.getBoundingClientRect().top + workspace.clientHeight / 3;
  for (let i = 0; i < pages.length; i++) {
    if (pages[i].getBoundingClientRect().bottom >= probe) return i + 1;
  }
  return Math.max(1, pages.length);
}

export default function App() {
  const [doc, setDoc] = useState<DocState>({ key: 0, file: null, name: UNTITLED, ref: null });
  const [dirty, setDirtyState] = useState(false);
  const [pages, setPages] = useState<number | null>(null);
  const [page, setPage] = useState(1);
  const [words, setWords] = useState<number | null>(null);
  const [handle, setHandle] = useState<EditorHandle | null>(null);
  const [backstage, setBackstage] = useState(false);
  const [prompt, setPrompt] = useState<PromptRequest | null>(null);
  const [find, setFind] = useState<{ tab: FindTab; query: string } | null>(null);
  const dirtyRef = useRef(false);
  const workspaceRef = useRef<HTMLElement | null>(null);
  const [workspaceEl, setWorkspaceEl] = useState<HTMLElement | null>(null);
  const setWorkspace = useCallback((el: HTMLElement | null) => {
    workspaceRef.current = el;
    setWorkspaceEl(el);
  }, []);

  const superdoc = handle?.superdoc ?? null;
  const { snapshot, execute } = useHeadlessToolbar(superdoc, RIBBON_COMMANDS);
  if (import.meta.env.DEV) (window as unknown as { toolbarSnapshot?: unknown }).toolbarSnapshot = snapshot;

  const askFields = useCallback(
    (title: string, fields: { label: string; initial: string }[]) =>
      new Promise<string[] | null>((resolve) => setPrompt({ title, fields, resolve })),
    [],
  );
  const ask = useCallback(
    (title: string, label: string, initial = '') =>
      askFields(title, [{ label, initial }]).then((values) => values?.[0] ?? null),
    [askFields],
  );
  const editor = useMemo(() => createEditorActions(superdoc, ask), [superdoc, ask]);
  const pageActions = useMemo(() => createPageActions(superdoc), [superdoc]);

  // Régua: lembrada entre sessões, como no Word (só neste navegador/computador).
  const [showRuler, setShowRuler] = useState(() => {
    try {
      return localStorage.getItem('textos.regua') === '1';
    } catch {
      return false;
    }
  });
  const toggleRuler = () => {
    setShowRuler((v) => {
      try {
        localStorage.setItem('textos.regua', v ? '0' : '1');
      } catch {
        // Sem armazenamento local: vale só nesta sessão.
      }
      return !v;
    });
  };

  const setDirty = useCallback((value: boolean) => {
    dirtyRef.current = value;
    setDirtyState(value);
  }, []);

  // Contagem de palavras: refeita pouco depois de cada edição, sem pesar na digitação.
  const wordTimer = useRef<number | undefined>(undefined);
  const countWords = useCallback(() => {
    window.clearTimeout(wordTimer.current);
    wordTimer.current = window.setTimeout(() => setWords(editor.wordCount()), 400);
  }, [editor]);
  useEffect(() => {
    if (superdoc) countWords();
  }, [superdoc, countWords]);

  const onChange = useCallback(() => {
    if (!dirtyRef.current) setDirty(true);
    countWords();
  }, [setDirty, countWords]);
  const onPages = useCallback((total: number) => {
    setPages(total);
    setPage(visiblePage(workspaceRef.current));
  }, []);
  const showError = useCallback((msg: string) => {
    void platform.alert(`O documento não pôde ser aberto corretamente.\n\n${msg}`);
  }, []);

  useEffect(() => {
    const ws = workspaceRef.current;
    if (!ws) return;
    const onScroll = () => setPage(visiblePage(ws));
    ws.addEventListener('scroll', onScroll, { passive: true });
    return () => ws.removeEventListener('scroll', onScroll);
  }, [doc.key]);

  const replaceDocument = (next: Omit<DocState, 'key'>) => {
    setDoc((prev) => ({ ...next, key: prev.key + 1 }));
    setPages(null);
    setPage(1);
    setWords(null);
    setDirty(false);
  };

  const confirmDiscard = async () =>
    !dirtyRef.current || platform.confirm('Há alterações não salvas. Deseja descartá-las?');

  const run = (action: () => Promise<unknown>) => () => {
    action().catch((e) => platform.alert(`Não foi possível concluir a operação.\n\n${errorMessage(e)}`));
  };

  const save = async (asNew: boolean) => {
    if (!handle) return;
    const ref = await platform.saveFile(await handle.exportDocx(), {
      ref: asNew ? null : doc.ref,
      suggestedName: `${doc.name}.docx`,
      filters: [DOCX_FILTER],
    });
    if (!ref) return;
    setDoc((prev) => ({ ...prev, ref, name: baseName(ref.name) }));
    setDirty(false);
  };

  const actions: FileActions = {
    newDocument: run(async () => {
      if (!(await confirmDiscard())) return;
      replaceDocument({ file: null, name: UNTITLED, ref: null });
    }),
    open: run(async () => {
      if (!(await confirmDiscard())) return;
      const opened = await platform.openFile([DOCX_FILTER]);
      if (!opened) return;
      const file = new File([opened.bytes as Uint8Array<ArrayBuffer>], opened.ref.name, { type: DOCX_MIME });
      replaceDocument({ file, name: baseName(opened.ref.name), ref: opened.ref });
    }),
    save: run(() => save(false)),
    saveAs: run(() => save(true)),
    print: () => window.print(),
  };

  // Pesquisa da barra de título: Enter vai para a próxima ocorrência, Shift+Enter para a anterior.
  const titleSearch = useRef('');
  const search = (query: string, backwards: boolean) => {
    const r = titleSearch.current === query ? editor.findNext(backwards) : editor.find(query, { caseSensitive: false });
    titleSearch.current = query;
    if (!r.count) {
      titleSearch.current = '';
      void platform.alert(`O Textos concluiu a pesquisa do documento. "${query}" não foi encontrado.`);
    }
  };
  // Uma edição pode mudar as ocorrências: a próxima pesquisa recomeça.
  const onEdit = useCallback(() => {
    titleSearch.current = '';
    onChange();
  }, [onChange]);

  const openFind = (tab: FindTab) => setFind({ tab, query: editor.selectedText() });

  // Atalhos do Word em português (src/shortcuts.ts).
  const fontSize = () => parseFloat(String(snapshot.commands['font-size']?.value ?? '')) || editor.defaultFont().size || 11;
  const runShortcut = (action: ShortcutAction) => {
    const lineSpacing = (v: number) => execute('line-height', v);
    const handlers: Record<ShortcutAction, () => unknown> = {
      new: actions.newDocument,
      open: actions.open,
      save: actions.save,
      saveAs: actions.saveAs,
      print: actions.print,
      find: () => openFind('find'),
      replace: () => openFind('replace'),
      undo: () => execute('undo'),
      redo: () => execute('redo'),
      selectAll: editor.selectAll,
      bold: () => execute('bold'),
      italic: () => execute('italic'),
      underline: () => execute('underline'),
      subscript: () => editor.toggleVertAlign('subscript'),
      superscript: () => editor.toggleVertAlign('superscript'),
      growFont: () => execute('font-size', `${stepFontSize(fontSize(), true)}pt`),
      shrinkFont: () => execute('font-size', `${stepFontSize(fontSize(), false)}pt`),
      growFontPoint: () => execute('font-size', `${Math.min(1638, Math.floor(fontSize()) + 1)}pt`),
      shrinkFontPoint: () => execute('font-size', `${Math.max(1, Math.ceil(fontSize()) - 1)}pt`),
      clearFormatting: () => execute('clear-formatting'),
      copyFormat: () => execute('copy-format'),
      alignLeft: () => execute('text-align', 'left'),
      alignCenter: () => execute('text-align', 'center'),
      alignRight: () => execute('text-align', 'right'),
      justify: () => execute('text-align', 'justify'),
      indent: () => (pageActions.paragraph()?.isList ? execute('indent-increase') : pageActions.stepIndent(1)),
      outdent: () => (pageActions.paragraph()?.isList ? execute('indent-decrease') : pageActions.stepIndent(-1)),
      lineSpacing1: () => lineSpacing(1),
      lineSpacing15: () => lineSpacing(1.5),
      lineSpacing2: () => lineSpacing(2),
      showAll: () => execute('formatting-marks'),
      heading1: () => editor.applyStyleByName('heading 1'),
      heading2: () => editor.applyStyleByName('heading 2'),
      heading3: () => editor.applyStyleByName('heading 3'),
      normalStyle: () => editor.applyStyleByName('Normal'),
      link: editor.insertLink,
      pageBreak: editor.insertPageBreak,
    };
    handlers[action]();
  };

  // O ref evita registrar o listener de novo a cada render.
  const shortcutRef = useRef(runShortcut);
  useLayoutEffect(() => {
    shortcutRef.current = runShortcut;
  });
  useEffect(() => {
    const onKeyDown = (e: KeyboardEvent) => {
      const target = e.target as HTMLElement | null;
      const inTextField = Boolean(target?.closest?.('input, textarea, select'));
      const action = matchShortcut(e, inTextField);
      if (!action) return;
      e.preventDefault();
      e.stopPropagation();
      shortcutRef.current(action);
    };
    // Captura: o atalho do Word vence o padrão do motor (que segue o Word em inglês: Ctrl+B negrito…).
    window.addEventListener('keydown', onKeyDown, true);
    return () => window.removeEventListener('keydown', onKeyDown, true);
  }, []);

  useEffect(() => platform.guardClose(() => dirtyRef.current), []);

  // Só em desenvolvimento: ?abrir=/bench/corpus/01-formatacao.docx abre um documento servido pelo Vite.
  useEffect(() => {
    const path = import.meta.env.DEV ? new URLSearchParams(location.search).get('abrir') : null;
    if (!path) return;
    void fetch(path)
      .then((r) => (r.ok ? r.blob() : Promise.reject(new Error(r.statusText))))
      .then((blob) => {
        const name = decodeURIComponent(path.split('/').pop() ?? 'documento.docx');
        replaceDocument({ file: new File([blob], name, { type: DOCX_MIME }), name: baseName(name), ref: null });
      })
      .catch((e) => platform.alert(`Não foi possível abrir ${path}: ${errorMessage(e)}`));
    // Uma vez, ao abrir o app.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  useEffect(() => {
    platform.setTitle(`${dirty ? '• ' : ''}${doc.name} — Textos`);
  }, [dirty, doc.name]);

  const zoom = Number(snapshot.commands.zoom?.value ?? 100);

  // "Largura da Página" do Word: zoom para a folha ocupar a largura da área do documento.
  const fitWidth = () => {
    const ws = workspaceRef.current;
    const section = pageActions.section();
    if (!ws || !section) return;
    const style = getComputedStyle(ws);
    const available = ws.clientWidth - parseFloat(style.paddingLeft) - parseFloat(style.paddingRight);
    const pagePx = section.width * 96;
    execute('zoom', Math.min(ZOOM_MAX, Math.max(ZOOM_MIN, Math.floor((available / pagePx) * 100))));
  };

  return (
    <div className="app">
      <TitleBar
        title={doc.name}
        dirty={dirty}
        canUndo={Boolean(snapshot.context) && !snapshot.commands.undo?.disabled}
        canRedo={Boolean(snapshot.context) && !snapshot.commands.redo?.disabled}
        onSave={actions.save}
        onUndo={() => execute('undo')}
        onRedo={() => execute('redo')}
        onSearch={search}
      />
      <Ribbon
        snapshot={snapshot}
        execute={execute}
        editor={editor}
        onOpenFile={() => setBackstage(true)}
        onFind={() => openFind('find')}
        onReplace={() => openFind('replace')}
        page={pageActions}
        askFields={askFields}
        showRuler={showRuler}
        onToggleRuler={toggleRuler}
        onFitWidth={fitWidth}
      />
      {showRuler && <Ruler page={pageActions} workspace={workspaceEl} pageIndex={page} revision={snapshot} />}
      <DocumentEditor
        key={doc.key}
        file={doc.file}
        onReady={setHandle}
        onChange={onEdit}
        onPages={onPages}
        onError={showError}
        workspaceRef={setWorkspace}
      />
      <StatusBar
        page={Math.min(page, pages ?? page)}
        pages={pages}
        words={words}
        zoom={zoom}
        onZoom={(z) => execute('zoom', Math.min(ZOOM_MAX, Math.max(ZOOM_MIN, z)))}
      />
      {find && (
        <FindReplace
          tab={find.tab}
          initialQuery={find.query}
          editor={editor}
          onTab={(tab) => setFind({ ...find, tab })}
          onClose={() => {
            setFind(null);
            editor.focus();
          }}
        />
      )}
      {backstage && <Backstage actions={actions} onClose={() => setBackstage(false)} />}
      {prompt && <PromptDialog request={prompt} onDone={() => setPrompt(null)} />}
    </div>
  );
}
