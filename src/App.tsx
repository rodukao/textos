import { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react';
import type { PublicToolbarItemId } from 'superdoc/headless-toolbar';
import { useHeadlessToolbar } from 'superdoc/headless-toolbar/react';
import { Backstage } from './components/Backstage';
import type { FileActions } from './components/file-actions';
import { PromptDialog, type PromptRequest } from './components/PromptDialog';
import { Ribbon } from './components/ribbon/Ribbon';
import { createEditorActions } from './components/ribbon/editor-actions';
import { StatusBar, ZOOM_MAX, ZOOM_MIN } from './components/StatusBar';
import { TitleBar } from './components/TitleBar';
import { DOCX_FILTER, baseName } from './document/format';
import { DocumentEditor, type EditorHandle } from './editor/DocumentEditor';
import { DOCX_MIME } from './engine/superdoc-config';
import { platform, type FileRef } from './platform';

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
  const dirtyRef = useRef(false);
  const workspaceRef = useRef<HTMLElement>(null);
  const searchRef = useRef<HTMLInputElement>(null);

  const superdoc = handle?.superdoc ?? null;
  const { snapshot, execute } = useHeadlessToolbar(superdoc, RIBBON_COMMANDS);
  if (import.meta.env.DEV) (window as unknown as { toolbarSnapshot?: unknown }).toolbarSnapshot = snapshot;

  const ask = useCallback(
    (title: string, label: string, initial = '') =>
      new Promise<string | null>((resolve) => setPrompt({ title, label, initial, resolve })),
    [],
  );
  const editor = useMemo(() => createEditorActions(superdoc, ask), [superdoc, ask]);

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

  // Localizar: Enter vai para a próxima ocorrência, Shift+Enter para a anterior.
  const searchState = useRef({ query: '', index: -1, matches: [] as unknown[] });
  const search = (query: string, backwards: boolean) => {
    const target = superdoc?.activeEditor as unknown as { commands: Record<string, (...a: unknown[]) => unknown> } | null;
    if (!target) return;
    const s = searchState.current;
    if (s.query !== query) {
      const found = target.commands.search?.(query);
      s.matches = Array.isArray(found) ? found : [];
      s.query = query;
      s.index = -1;
    }
    if (!s.matches.length) {
      void platform.alert(`Não encontramos "${query}" no documento.`);
      return;
    }
    s.index = (s.index + (backwards ? -1 : 1) + s.matches.length) % s.matches.length;
    target.commands.goToSearchResult?.(s.matches[s.index]);
  };
  // Uma edição pode mudar as ocorrências: a próxima busca recomeça.
  const onEdit = useCallback(() => {
    searchState.current.query = '';
    onChange();
  }, [onChange]);

  const focusSearch = () => {
    searchRef.current?.focus();
    searchRef.current?.select();
  };

  // Atalhos globais. O ref evita registrar o listener de novo a cada render.
  const actionsRef = useRef({ ...actions, focusSearch });
  useLayoutEffect(() => {
    actionsRef.current = { ...actions, focusSearch };
  });
  useEffect(() => {
    const onKeyDown = (e: KeyboardEvent) => {
      if (!(e.ctrlKey || e.metaKey)) return;
      const a = actionsRef.current;
      const key = e.key.toLowerCase();
      const handler =
        key === 's' ? (e.shiftKey ? a.saveAs : a.save)
        : key === 'o' ? a.open
        : key === 'p' ? a.print
        : key === 'n' && e.altKey ? a.newDocument
        : key === 'f' && !e.shiftKey ? a.focusSearch
        : null;
      if (!handler) return;
      e.preventDefault();
      e.stopPropagation();
      handler();
    };
    // Captura: o atalho do app vence o do editor.
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
        searchRef={searchRef}
      />
      <Ribbon
        snapshot={snapshot}
        execute={execute}
        editor={editor}
        onOpenFile={() => setBackstage(true)}
        onFind={focusSearch}
      />
      <DocumentEditor
        key={doc.key}
        file={doc.file}
        onReady={setHandle}
        onChange={onEdit}
        onPages={onPages}
        onError={showError}
        workspaceRef={workspaceRef}
      />
      <StatusBar
        page={Math.min(page, pages ?? page)}
        pages={pages}
        words={words}
        zoom={zoom}
        onZoom={(z) => execute('zoom', Math.min(ZOOM_MAX, Math.max(ZOOM_MIN, z)))}
      />
      {backstage && <Backstage actions={actions} onClose={() => setBackstage(false)} />}
      {prompt && <PromptDialog request={prompt} onDone={() => setPrompt(null)} />}
    </div>
  );
}
