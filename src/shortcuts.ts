/**
 * Atalhos de teclado do Word em português (Ctrl+N negrito, Ctrl+S sublinhado, Ctrl+B salvar…).
 * A mesma tabela serve ao teclado e às dicas dos botões, para as duas nunca divergirem.
 * No macOS vale Cmd no lugar de Ctrl.
 */
export type ShortcutAction =
  | 'new'
  | 'open'
  | 'save'
  | 'saveAs'
  | 'print'
  | 'find'
  | 'replace'
  | 'undo'
  | 'redo'
  | 'selectAll'
  | 'bold'
  | 'italic'
  | 'underline'
  | 'subscript'
  | 'superscript'
  | 'growFont'
  | 'shrinkFont'
  | 'growFontPoint'
  | 'shrinkFontPoint'
  | 'clearFormatting'
  | 'copyFormat'
  | 'alignLeft'
  | 'alignCenter'
  | 'alignRight'
  | 'justify'
  | 'indent'
  | 'outdent'
  | 'lineSpacing1'
  | 'lineSpacing15'
  | 'lineSpacing2'
  | 'showAll'
  | 'heading1'
  | 'heading2'
  | 'heading3'
  | 'normalStyle'
  | 'link'
  | 'pageBreak';

interface Shortcut {
  action: ShortcutAction;
  /** `e.key` em minúsculas ("n", "f12", "+", " "). */
  key: string;
  shift?: boolean;
  alt?: boolean;
  /** Sem Ctrl (ex.: F12). */
  noCtrl?: boolean;
  /** Como aparece na dica, quando não é só a tecla. */
  label?: string;
}

const SHORTCUTS: Shortcut[] = [
  { action: 'new', key: 'o' },
  { action: 'open', key: 'a' },
  { action: 'open', key: 'f12', label: 'Ctrl+F12' },
  { action: 'save', key: 'b' },
  { action: 'saveAs', key: 'f12', noCtrl: true, label: 'F12' },
  { action: 'print', key: 'p' },
  { action: 'find', key: 'l' },
  { action: 'find', key: 'f' },
  { action: 'replace', key: 'u' },
  { action: 'undo', key: 'z' },
  { action: 'redo', key: 'r' },
  { action: 'redo', key: 'y' },
  { action: 'selectAll', key: 't' },
  { action: 'bold', key: 'n' },
  { action: 'italic', key: 'i' },
  { action: 'underline', key: 's' },
  { action: 'subscript', key: '=', label: 'Ctrl+=' },
  { action: 'superscript', key: '+', shift: true, label: 'Ctrl+Shift++' },
  { action: 'growFont', key: '>', shift: true, label: 'Ctrl+Shift+>' },
  { action: 'shrinkFont', key: '<', shift: true, label: 'Ctrl+Shift+<' },
  { action: 'growFontPoint', key: ']', label: 'Ctrl+]' },
  { action: 'shrinkFontPoint', key: '[', label: 'Ctrl+[' },
  { action: 'clearFormatting', key: ' ', label: 'Ctrl+Espaço' },
  { action: 'copyFormat', key: 'c', shift: true },
  { action: 'alignLeft', key: 'q' },
  { action: 'alignCenter', key: 'e' },
  { action: 'alignRight', key: 'g' },
  { action: 'justify', key: 'j' },
  { action: 'indent', key: 'm' },
  { action: 'outdent', key: 'm', shift: true },
  { action: 'lineSpacing1', key: '1' },
  { action: 'lineSpacing15', key: '5' },
  { action: 'lineSpacing2', key: '2' },
  { action: 'showAll', key: '*', shift: true, label: 'Ctrl+*' },
  { action: 'heading1', key: '1', alt: true },
  { action: 'heading2', key: '2', alt: true },
  { action: 'heading3', key: '3', alt: true },
  { action: 'normalStyle', key: 'n', shift: true },
  { action: 'link', key: 'k' },
  { action: 'pageBreak', key: 'enter' },
];

/** Atalhos que valem mesmo com o cursor numa caixa de texto da interface (pesquisa, diálogos). */
const GLOBAL_ACTIONS = new Set<ShortcutAction>(['new', 'open', 'save', 'saveAs', 'print', 'find', 'replace']);

export function matchShortcut(e: KeyboardEvent, inTextField: boolean): ShortcutAction | null {
  const ctrl = e.ctrlKey || e.metaKey;
  const key = e.key.toLowerCase();
  // Shift muda a tecla em símbolos ("+", "*", ">"): nesses casos o Shift já está no próprio `key`.
  const shiftSymbol = key.length === 1 && !/[a-z0-9 ]/.test(key);
  const found = SHORTCUTS.find(
    (s) =>
      s.key === key &&
      (s.noCtrl ? !ctrl : ctrl) &&
      Boolean(s.alt) === e.altKey &&
      (shiftSymbol || Boolean(s.shift) === e.shiftKey),
  );
  if (!found) return null;
  if (inTextField && !GLOBAL_ACTIONS.has(found.action)) return null;
  return found.action;
}

const isMac = typeof navigator !== 'undefined' && /Mac|iPhone|iPad/.test(navigator.platform);

/** "Ctrl+N" (ou "Cmd+N" no macOS) do primeiro atalho da ação. */
export function shortcutLabel(action: ShortcutAction): string {
  const s = SHORTCUTS.find((x) => x.action === action);
  if (!s) return '';
  const label =
    s.label ??
    [s.noCtrl ? null : 'Ctrl', s.alt ? 'Alt' : null, s.shift ? 'Shift' : null, s.key === 'enter' ? 'Enter' : s.key.toUpperCase()]
      .filter(Boolean)
      .join('+');
  return isMac ? label.replace('Ctrl', 'Cmd') : label;
}

/** Dica de botão no formato do Word: "Negrito (Ctrl+N)". */
export const withShortcut = (title: string, action: ShortcutAction) => `${title} (${shortcutLabel(action)})`;
