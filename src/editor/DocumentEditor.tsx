import { useEffect, useRef } from 'react';
import { SuperDoc } from 'superdoc';
import 'superdoc/style.css';
import blankDocumentUrl from '../document/documento-em-branco.docx?url';
import { DOCX_MIME, baseSuperDocConfig } from '../engine/superdoc-config';

/**
 * "Documento em branco" do Word em português, gerado pelo próprio Word: Calibri 11, 8 pt depois
 * do parágrafo, entrelinha 1,08, A4 com margens de 2,5 e 3 cm, idioma pt-BR e os estilos rápidos
 * (Título 1, Sem Espaçamento…). O modelo interno do motor era outro (Arial 12, sem espaço entre
 * parágrafos, em inglês).
 */
async function blankDocument(): Promise<File> {
  const bytes = await (await fetch(blankDocumentUrl)).arrayBuffer();
  return new File([bytes], 'Documento1.docx', { type: DOCX_MIME });
}

/** O que o App pode fazer com o documento aberto. */
export interface EditorHandle {
  superdoc: SuperDoc;
  exportDocx(): Promise<Blob>;
}

export interface DocumentEditorProps {
  /** Documento a abrir; null = documento em branco. */
  file: File | null;
  onReady: (handle: EditorHandle | null) => void;
  onChange: () => void;
  onPages: (pages: number) => void;
  onError: (message: string) => void;
  /** A área rolável das folhas, para a barra de status saber a página visível. */
  workspaceRef?: React.Ref<HTMLElement>;
}

/**
 * As folhas de UM documento, usando o motor SuperDoc (lê e grava .docx preservando
 * a formatação do Word). A faixa de opções é do Textos (components/ribbon) e fala com
 * o motor pela barra "headless" dele.
 * O App troca a `key` deste componente ao criar/abrir um documento: o motor é
 * recriado do zero, com histórico de desfazer limpo.
 */
export function DocumentEditor(props: DocumentEditorProps) {
  const editorRef = useRef<HTMLDivElement>(null);
  // Callbacks sempre atualizados, sem recriar o motor a cada render.
  const callbacks = useRef(props);
  callbacks.current = props;

  useEffect(() => {
    let ready = false;
    let cancelled = false;
    let superdoc: SuperDoc | null = null;

    const createEngine = (file: File | null) =>
      new SuperDoc({
        ...baseSuperDocConfig,
        selector: editorRef.current!,
        document: file ?? undefined,
        documentMode: 'editing',
        onReady: () => {
          ready = true;
          const instance = superdoc!;
          // Só em desenvolvimento: para inspecionar o motor pelo console.
          if (import.meta.env.DEV) (window as unknown as { superdoc?: SuperDoc }).superdoc = instance;
          callbacks.current.onReady({
            superdoc: instance,
            exportDocx: () => instance.export({ exportType: ['docx'], triggerDownload: false }),
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

    void (async () => {
      // Sem arquivo: o "Documento em branco" do Word (se não carregar, o modelo do motor).
      const file = props.file ?? (await blankDocument().catch(() => null));
      if (!cancelled) superdoc = createEngine(file);
    })();

    return () => {
      cancelled = true;
      callbacks.current.onReady(null);
      superdoc?.destroy();
    };
    // O documento só muda com uma nova `key`; o motor é criado uma vez por montagem.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  return (
    <main className="workspace" ref={props.workspaceRef}>
      <div ref={editorRef} />
    </main>
  );
}
