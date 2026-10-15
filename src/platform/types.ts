/**
 * Tudo que depende de onde o app está rodando (navegador ou desktop) passa por
 * esta interface. O resto do app não sabe — nem precisa saber — qual é.
 */

export interface FileFilter {
  name: string;
  /** Extensões sem o ponto, ex.: ['docx'] */
  extensions: string[];
}

/** Referência a um arquivo já aberto/salvo, usada para "Salvar" sem perguntar de novo. */
export interface FileRef {
  name: string;
  /** Caminho absoluto (desktop). */
  path?: string;
  /** Handle da File System Access API (navegadores Chromium). */
  handle?: FileSystemFileHandle;
}

export interface OpenedFile {
  ref: FileRef;
  bytes: Uint8Array;
}

export interface SaveOptions {
  /** Se presente (e gravável), salva direto nele; senão pergunta onde salvar. */
  ref?: FileRef | null;
  suggestedName: string;
  filters: FileFilter[];
}

export interface Platform {
  readonly kind: 'web' | 'desktop';
  /** Retorna null se o usuário cancelar. */
  openFile(filters: FileFilter[]): Promise<OpenedFile | null>;
  /** Retorna null se o usuário cancelar. */
  saveFile(data: Blob, options: SaveOptions): Promise<FileRef | null>;
  confirm(message: string): Promise<boolean>;
  alert(message: string): Promise<void>;
  setTitle(title: string): void;
  /**
   * Pede confirmação ao fechar a janela/aba quando `hasUnsavedChanges()` for true.
   * Retorna uma função que remove o guarda.
   */
  guardClose(hasUnsavedChanges: () => boolean): () => void;
}
