import { useEffect, useRef } from 'react';
import { SuperDoc } from 'superdoc';
import 'superdoc/style.css';
import { baseSuperDocConfig } from '../engine/superdoc-config';

/** O que o App pode fazer com o documento aberto. */
export interface EditorHandle {
  exportDocx(): Promise<Blob>;
}

export interface DocumentEditorProps {
  /** Documento a abrir; null = documento em branco. */
  file: File | null;
  onReady: (handle: EditorHandle | null) => void;
  onChange: () => void;
  onPages: (pages: number) => void;
  onError: (message: string) => void;
}

/**
 * Barra de ferramentas + folhas de UM documento, usando o motor SuperDoc
 * (lê e grava .docx preservando a formatação do Word).
 * O App troca a `key` deste componente ao criar/abrir um documento: o motor é
 * recriado do zero, com histórico de desfazer limpo.
 */
/** Só existe um editor por vez, então um id fixo basta. */
const TOOLBAR_ID = 'textos-toolbar';

export function DocumentEditor(props: DocumentEditorProps) {
  const editorRef = useRef<HTMLDivElement>(null);
  // Callbacks sempre atualizados, sem recriar o motor a cada render.
  const callbacks = useRef(props);
  callbacks.current = props;

  useEffect(() => {
    let ready = false;
    const superdoc = new SuperDoc({
      ...baseSuperDocConfig,
      selector: editorRef.current!,
      document: props.file ?? undefined,
      documentMode: 'editing',
      // No SuperDoc 1.x a barra é indicada por seletor (id), não por elemento.
      toolbar: `#${TOOLBAR_ID}`,
      onReady: () => {
        ready = true;
        callbacks.current.onReady({
          exportDocx: () => superdoc.export({ exportType: ['docx'], triggerDownload: false }),
        });
      },
      // Atualizações durante o carregamento não são edições do usuário.
      onEditorUpdate: () => {
        if (ready) callbacks.current.onChange();
      },
      onPaginationUpdate: ({ totalPages }) => callbacks.current.onPages(totalPages),
      onException: (e) => {
        const err = (e as { error?: unknown }).error;
        callbacks.current.onError(err instanceof Error ? err.message : String(err ?? 'erro desconhecido'));
      },
    });
    return () => {
      callbacks.current.onReady(null);
      superdoc.destroy();
    };
    // O documento só muda com uma nova `key`; o motor é criado uma vez por montagem.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  return (
    <>
      <div className="sd-toolbar" id={TOOLBAR_ID} />
      <main className="workspace">
        <div ref={editorRef} />
      </main>
    </>
  );
}
