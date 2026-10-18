/** Comandos de arquivo, usados pela guia Arquivo, pela barra de título e pelos atalhos. */
export interface FileActions {
  newDocument: () => void;
  open: () => void;
  save: () => void;
  saveAs: () => void;
  print: () => void;
}
