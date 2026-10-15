import type { FileFilter, FileRef, OpenedFile, Platform, SaveOptions } from './types';

// File System Access API: existe no Chrome/Edge/Opera, mas ainda não está nos tipos do TS.
interface PickerType {
  description: string;
  accept: Record<string, string[]>;
}
declare global {
  interface Window {
    showOpenFilePicker?: (options?: { types?: PickerType[]; multiple?: boolean }) => Promise<FileSystemFileHandle[]>;
    showSaveFilePicker?: (options?: { suggestedName?: string; types?: PickerType[] }) => Promise<FileSystemFileHandle>;
  }
}

const MIME: Record<string, string> = {
  docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
};
const mimeOf = (ext: string) => MIME[ext] ?? 'application/octet-stream';

function toPickerTypes(filters: FileFilter[]): PickerType[] {
  return filters.map((f) => {
    const accept: Record<string, string[]> = {};
    for (const ext of f.extensions) (accept[mimeOf(ext)] ??= []).push(`.${ext}`);
    return { description: f.name, accept };
  });
}

const isAbort = (e: unknown) => e instanceof DOMException && e.name === 'AbortError';

/** Fallback para Firefox/Safari: <input type="file">. */
function pickWithInput(filters: FileFilter[]): Promise<File | null> {
  return new Promise((resolve) => {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = filters.flatMap((f) => f.extensions.map((e) => `.${e}`)).join(',');
    input.addEventListener('change', () => resolve(input.files?.[0] ?? null), { once: true });
    input.addEventListener('cancel', () => resolve(null), { once: true });
    input.click();
  });
}

/** Fallback para Firefox/Safari: força o download. */
function download(data: Blob, name: string) {
  const url = URL.createObjectURL(data);
  const a = document.createElement('a');
  a.href = url;
  a.download = name;
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

async function writeHandle(handle: FileSystemFileHandle, data: Blob) {
  const writable = await handle.createWritable();
  await writable.write(data);
  await writable.close();
}

const readFile = async (file: File) => new Uint8Array(await file.arrayBuffer());

export const webPlatform: Platform = {
  kind: 'web',

  async openFile(filters): Promise<OpenedFile | null> {
    if (window.showOpenFilePicker) {
      try {
        const [handle] = await window.showOpenFilePicker({ types: toPickerTypes(filters) });
        const file = await handle.getFile();
        return { ref: { name: file.name, handle }, bytes: await readFile(file) };
      } catch (e) {
        if (isAbort(e)) return null;
        throw e;
      }
    }
    const file = await pickWithInput(filters);
    return file ? { ref: { name: file.name }, bytes: await readFile(file) } : null;
  },

  async saveFile(data, { ref, suggestedName, filters }: SaveOptions): Promise<FileRef | null> {
    if (ref?.handle) {
      await writeHandle(ref.handle, data);
      return ref;
    }
    if (window.showSaveFilePicker) {
      try {
        const handle = await window.showSaveFilePicker({ suggestedName, types: toPickerTypes(filters) });
        await writeHandle(handle, data);
        return { name: handle.name, handle };
      } catch (e) {
        if (isAbort(e)) return null;
        throw e;
      }
    }
    download(data, ref?.name ?? suggestedName);
    return { name: ref?.name ?? suggestedName };
  },

  async confirm(message) {
    return window.confirm(message);
  },

  async alert(message) {
    window.alert(message);
  },

  setTitle(title) {
    document.title = title;
  },

  guardClose(hasUnsavedChanges) {
    const onBeforeUnload = (e: BeforeUnloadEvent) => {
      if (hasUnsavedChanges()) e.preventDefault();
    };
    window.addEventListener('beforeunload', onBeforeUnload);
    return () => window.removeEventListener('beforeunload', onBeforeUnload);
  },
};
