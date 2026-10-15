import type { FileFilter } from '../platform';

/** O formato nativo do Textos é o do Word. */
export const DOCX_FILTER: FileFilter = { name: 'Documento do Word', extensions: ['docx'] };

export function baseName(name: string): string {
  const i = name.lastIndexOf('.');
  return i > 0 ? name.slice(0, i) : name;
}
