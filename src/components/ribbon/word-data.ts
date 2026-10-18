/** Listas da faixa de opções, iguais às do Word em português. */

export const FONT_SIZES = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];

/** "Aumentar/Diminuir Tamanho da Fonte" do Word: anda pela lista; acima de 72, de 10 em 10. */
export function stepFontSize(current: number, up: boolean): number {
  if (up) {
    const next = FONT_SIZES.find((s) => s > current);
    return next ?? Math.min(1638, Math.floor(current / 10) * 10 + 10);
  }
  if (current > 72) return Math.max(72, Math.ceil(current / 10) * 10 - 10);
  const prev = [...FONT_SIZES].reverse().find((s) => s < current);
  return prev ?? Math.max(1, current - 1);
}

/** Fontes mais usadas em documentos do Word (as do documento entram também na lista). */
export const COMMON_FONTS = [
  'Aptos',
  'Arial',
  'Arial Black',
  'Book Antiqua',
  'Calibri',
  'Calibri Light',
  'Cambria',
  'Candara',
  'Century Gothic',
  'Comic Sans MS',
  'Consolas',
  'Constantia',
  'Corbel',
  'Courier New',
  'Franklin Gothic Medium',
  'Garamond',
  'Georgia',
  'Lucida Console',
  'Palatino Linotype',
  'Segoe UI',
  'Tahoma',
  'Times New Roman',
  'Trebuchet MS',
  'Verdana',
];

export const LINE_SPACINGS = [1, 1.15, 1.5, 2, 2.5, 3];

/** Cores do tema Office: primeira linha e variações (mais claras 80/60/40%, mais escuras 25/50%). */
const THEME_BASE = ['#FFFFFF', '#000000', '#E7E6E6', '#44546A', '#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47'];

const mix = (hex: string, target: number, amount: number) => {
  const n = parseInt(hex.slice(1), 16);
  const ch = [(n >> 16) & 255, (n >> 8) & 255, n & 255].map((c) => Math.round(c + (target - c) * amount));
  return `#${ch.map((c) => c.toString(16).padStart(2, '0')).join('')}`.toUpperCase();
};

/** Como no Word: branco e preto escurecem/clareiam de outro jeito que as cores. */
function variations(hex: string): string[] {
  if (hex === '#FFFFFF') return [0.05, 0.15, 0.25, 0.35, 0.5].map((a) => mix(hex, 0, a));
  if (hex === '#000000') return [0.5, 0.35, 0.25, 0.15, 0.05].map((a) => mix(hex, 255, a));
  return [...[0.8, 0.6, 0.4].map((a) => mix(hex, 255, a)), ...[0.25, 0.5].map((a) => mix(hex, 0, a))];
}

export const THEME_COLOR_ROWS: string[][] = [
  THEME_BASE,
  ...[0, 1, 2, 3, 4].map((i) => THEME_BASE.map((c) => variations(c)[i])),
];

export const STANDARD_COLORS = ['#C00000', '#FF0000', '#FFC000', '#FFFF00', '#92D050', '#00B050', '#00B0F0', '#0070C0', '#002060', '#7030A0'];

/** As 15 cores de realce do Word, na ordem da paleta dele. */
export const HIGHLIGHT_ROWS = [
  ['#FFFF00', '#00FF00', '#00FFFF', '#FF00FF', '#0000FF'],
  ['#FF0000', '#000080', '#008080', '#008000', '#800080'],
  ['#800000', '#808000', '#808080', '#C0C0C0', '#000000'],
];

/** Nomes dos estilos internos do Word como aparecem no Word em português. */
const STYLE_NAMES_PT: Record<string, string> = {
  normal: 'Normal',
  'no spacing': 'Sem Espaçamento',
  title: 'Título',
  subtitle: 'Subtítulo',
  quote: 'Citação',
  'intense quote': 'Citação Intensa',
  'list paragraph': 'Parágrafo da Lista',
  caption: 'Legenda',
  'toc heading': 'Cabeçalho do Sumário',
  'subtle emphasis': 'Ênfase Sutil',
  emphasis: 'Ênfase',
  'intense emphasis': 'Ênfase Intensa',
  strong: 'Forte',
  'subtle reference': 'Referência Sutil',
  'intense reference': 'Referência Intensa',
  'book title': 'Título do Livro',
  header: 'Cabeçalho',
  footer: 'Rodapé',
  'footnote text': 'Texto de nota de rodapé',
  'table paragraph': 'Parágrafo da Tabela',
  'body text': 'Corpo de Texto',
};

export function styleDisplayName(name: string): string {
  const key = name.trim().toLowerCase();
  const heading = /^heading (\d)$/.exec(key);
  if (heading) return `Título ${heading[1]}`;
  const toc = /^toc (\d)$/.exec(key);
  if (toc) return `Sumário ${toc[1]}`;
  return STYLE_NAMES_PT[key] ?? name;
}

/** Ordem da galeria "Estilos" do Word; os demais estilos rápidos vêm depois, em ordem alfabética. */
export const QUICK_STYLE_ORDER = ['Normal', 'No Spacing', 'heading 1', 'heading 2', 'heading 3', 'Title', 'Subtitle'];
