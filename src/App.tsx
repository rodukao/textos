import { useCallback, useEffect, useLayoutEffect, useRef, useState } from 'react';
import { FileBar, type FileActions } from './components/FileBar';
import { StatusBar } from './components/StatusBar';
import { DOCX_FILTER, baseName } from './document/format';
import { DocumentEditor, type EditorHandle } from './editor/DocumentEditor';
import { DOCX_MIME } from './engine/superdoc-config';
import { platform, type FileRef } from './platform';

const UNTITLED = 'Documento sem título';
const errorMessage = (e: unknown) => (e instanceof Error ? e.message : String(e));

interface DocState {
  /** Muda a cada documento novo/aberto e remonta o DocumentEditor. */
  key: number;
  file: File | null;
  name: string;
  ref: FileRef | null;
}

export default function App() {
  const [doc, setDoc] = useState<DocState>({ key: 0, file: null, name: UNTITLED, ref: null });
  const [dirty, setDirtyState] = useState(false);
  const [pages, setPages] = useState<number | null>(null);
  const dirtyRef = useRef(false);
  const editorRef = useRef<EditorHandle | null>(null);

  const setDirty = useCallback((value: boolean) => {
    dirtyRef.current = value;
    setDirtyState(value);
  }, []);
  const markDirty = useCallback(() => {
    if (!dirtyRef.current) setDirty(true);
  }, [setDirty]);
  const setEditor = useCallback((handle: EditorHandle | null) => {
    editorRef.current = handle;
  }, []);
  const showError = useCallback((msg: string) => {
    void platform.alert(`O documento não pôde ser aberto corretamente.\n\n${msg}`);
  }, []);

  const replaceDocument = (next: Omit<DocState, 'key'>) => {
    setDoc((prev) => ({ ...next, key: prev.key + 1 }));
    setPages(null);
    setDirty(false);
  };

  const confirmDiscard = async () =>
    !dirtyRef.current || platform.confirm('Há alterações não salvas. Deseja descartá-las?');

  const run = (action: () => Promise<unknown>) => () => {
    action().catch((e) => platform.alert(`Não foi possível concluir a operação.\n\n${errorMessage(e)}`));
  };

  const save = async (asNew: boolean) => {
    const editor = editorRef.current;
    if (!editor) return;
    const ref = await platform.saveFile(await editor.exportDocx(), {
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

  // Atalhos globais. O ref evita registrar o listener de novo a cada render.
  const actionsRef = useRef(actions);
  useLayoutEffect(() => {
    actionsRef.current = actions;
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

  useEffect(() => {
    platform.setTitle(`${dirty ? '• ' : ''}${doc.name} — Textos`);
  }, [dirty, doc.name]);

  return (
    <div className="app">
      <FileBar title={doc.name} dirty={dirty} actions={actions} />
      <DocumentEditor
        key={doc.key}
        file={doc.file}
        onReady={setEditor}
        onChange={markDirty}
        onPages={setPages}
        onError={showError}
      />
      <StatusBar pages={pages} platformKind={platform.kind} />
    </div>
  );
}
