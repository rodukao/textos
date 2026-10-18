import { getCurrentWindow } from '@tauri-apps/api/window';
import { ask, message, open, save } from '@tauri-apps/plugin-dialog';
import { readFile, writeFile } from '@tauri-apps/plugin-fs';
import type { FileRef, OpenedFile, Platform } from './types';

const fileName = (path: string) => path.split(/[\\/]/).pop() ?? path;

export const tauriPlatform: Platform = {
  kind: 'desktop',

  windowControls: {
    minimize: () => void getCurrentWindow().minimize(),
    toggleMaximize: () => void getCurrentWindow().toggleMaximize(),
    // Passa pelo onCloseRequested, que pergunta se há alterações não salvas.
    close: () => void getCurrentWindow().close(),
    onMaximizedChange(listener) {
      const win = getCurrentWindow();
      const check = () => void win.isMaximized().then(listener);
      check();
      const unlisten = win.onResized(check);
      return () => void unlisten.then((fn) => fn());
    },
  },

  async openFile(filters): Promise<OpenedFile | null> {
    const path = await open({ multiple: false, directory: false, filters });
    if (!path) return null;
    return { ref: { name: fileName(path), path }, bytes: await readFile(path) };
  },

  async saveFile(data, { ref, suggestedName, filters }): Promise<FileRef | null> {
    const path = ref?.path ?? (await save({ defaultPath: suggestedName, filters }));
    if (!path) return null;
    await writeFile(path, new Uint8Array(await data.arrayBuffer()));
    return { name: fileName(path), path };
  },

  confirm(msg) {
    return ask(msg, { title: 'Textos', kind: 'warning', okLabel: 'Sim', cancelLabel: 'Cancelar' });
  },

  async alert(msg) {
    await message(msg, { title: 'Textos', kind: 'error' });
  },

  setTitle(title) {
    document.title = title;
    void getCurrentWindow().setTitle(title);
  },

  guardClose(hasUnsavedChanges) {
    const unlisten = getCurrentWindow().onCloseRequested(async (event) => {
      if (!hasUnsavedChanges()) return;
      const leave = await ask('Há alterações não salvas. Deseja sair mesmo assim?', {
        title: 'Textos',
        kind: 'warning',
        okLabel: 'Sair sem salvar',
        cancelLabel: 'Cancelar',
      });
      if (!leave) event.preventDefault();
    });
    return () => void unlisten.then((fn) => fn());
  },
};
