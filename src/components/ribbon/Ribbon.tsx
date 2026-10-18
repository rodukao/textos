import { useState } from 'react';
import {
  AArrowDown,
  AArrowUp,
  CircleCheck,
  CircleX,
  ClipboardPaste,
  Copy,
  FilePenLine,
  Image,
  Link,
  ListIndentDecrease,
  ListIndentIncrease,
  List,
  ListChevronsUpDown,
  ListOrdered,
  MoveHorizontal,
  Paintbrush,
  Pilcrow,
  RemoveFormatting,
  Replace,
  Ruler,
  Scissors,
  Search,
  SeparatorHorizontal,
  SquareMousePointer,
  Strikethrough,
  Subscript,
  Superscript,
  Table,
  TableOfContents,
  TextAlignCenter,
  TextAlignEnd,
  TextAlignJustify,
  TextAlignStart,
  Highlighter,
  Bold,
  Italic,
  Underline,
  Search as ZoomIcon,
} from 'lucide-react';
import type { PublicToolbarItemId, ToolbarSnapshot } from 'superdoc/headless-toolbar';
import {
  ColorGlyph,
  ColorGrid,
  ComboBox,
  MenuButton,
  MenuItem,
  RibbonButton,
  RibbonGroup,
  Row,
  Rows,
} from './controls';
import { LayoutTab, type AskFields } from './LayoutTab';
import type { PageActions } from './page-actions';
import { StylesGallery } from './StylesGallery';
import {
  COMMON_FONTS,
  FONT_SIZES,
  HIGHLIGHT_ROWS,
  LINE_SPACINGS,
  STANDARD_COLORS,
  THEME_COLOR_ROWS,
  stepFontSize,
} from './word-data';
import type { EditorActions } from './editor-actions';
import { withShortcut } from '../../shortcuts';

export type Execute = (id: PublicToolbarItemId, payload?: unknown) => boolean;

export interface RibbonProps {
  snapshot: ToolbarSnapshot;
  execute: Execute;
  editor: EditorActions;
  onOpenFile: () => void;
  onFind: () => void;
  onReplace: () => void;
  page: PageActions;
  askFields: AskFields;
  showRuler: boolean;
  onFitWidth: () => void;
  onToggleRuler: () => void;
}

type TabId = 'home' | 'insert' | 'layout' | 'review' | 'view';

const TABS: { id: TabId | 'file'; label: string }[] = [
  { id: 'file', label: 'Arquivo' },
  { id: 'home', label: 'Página Inicial' },
  { id: 'insert', label: 'Inserir' },
  { id: 'layout', label: 'Layout' },
  { id: 'review', label: 'Revisão' },
  { id: 'view', label: 'Exibir' },
];

/** Primeira fonte da pilha CSS ("Calibri, sans-serif" → "Calibri"). */
const firstFamily = (value: unknown) =>
  typeof value === 'string' ? value.split(',')[0].trim().replace(/^["']|["']$/g, '') : '';

const parseSize = (value: unknown) => {
  const n = parseFloat(String(value ?? ''));
  return Number.isFinite(n) ? n : null;
};

const formatSize = (n: number) => String(n).replace('.', ',');

export function Ribbon(props: RibbonProps) {
  const [tab, setTab] = useState<TabId>('home');
  return (
    <div className="ribbon-wrap">
      <nav className="ribbon-tabs" role="tablist" aria-label="Faixa de opções">
        {TABS.map((t) => (
          <button
            key={t.id}
            type="button"
            role="tab"
            aria-selected={t.id === tab}
            className={`ribbon-tab${t.id === tab ? ' is-selected' : ''}`}
            onMouseDown={(e) => e.preventDefault()}
            onClick={() => (t.id === 'file' ? props.onOpenFile() : setTab(t.id))}
          >
            {t.label}
          </button>
        ))}
      </nav>
      <div className="ribbon" role="tabpanel">
        {tab === 'home' && <HomeTab {...props} />}
        {tab === 'insert' && <InsertTab {...props} />}
        {tab === 'layout' && <LayoutTab page={props.page} ask={props.askFields} insertPageBreak={props.editor.insertPageBreak} />}
        {tab === 'review' && <ReviewTab {...props} />}
        {tab === 'view' && <ViewTab {...props} />}
      </div>
    </div>
  );
}

function useCommand(props: RibbonProps) {
  const { snapshot, execute } = props;
  const state = (id: PublicToolbarItemId) => snapshot.commands[id];
  return {
    state,
    active: (id: PublicToolbarItemId) => Boolean(state(id)?.active),
    disabled: (id: PublicToolbarItemId) => !snapshot.context || Boolean(state(id)?.disabled),
    run: (id: PublicToolbarItemId, payload?: unknown) => () => {
      execute(id, payload);
    },
  };
}

function HomeTab(props: RibbonProps) {
  const { state, active, disabled, run } = useCommand(props);
  const { editor, execute } = props;
  const [fontColor, setFontColor] = useState('#FF0000');
  const [highlight, setHighlight] = useState('#FFFF00');

  // Sem valor: seleção mista (Word deixa a caixa vazia) ou parágrafo vazio (vale o padrão do documento).
  const defaults = editor.defaultFont();
  const familyState = state('font-family');
  const family = firstFamily(familyState?.value) || (familyState?.active ? '' : (defaults.family ?? ''));
  const families = [...new Set([...COMMON_FONTS, ...editor.documentFonts(), ...(family ? [family] : [])])].sort((a, b) => a.localeCompare(b));
  const sizeState = state('font-size');
  const size = parseSize(sizeState?.value) ?? (sizeState?.active ? null : defaults.size);
  const setSize = (n: number) => execute('font-size', `${n}pt`);
  const align = state('text-align')?.value ?? 'left';
  const lineHeight = Number(state('line-height')?.value ?? 1);

  return (
    <>
      <RibbonGroup label="Área de Transferência">
        <RibbonButton large icon={ClipboardPaste} label="Colar" title="Colar (Ctrl+V)" onClick={editor.paste} />
        <Rows>
          <RibbonButton icon={Scissors} label="Recortar" title="Recortar (Ctrl+X)" onClick={editor.cut} />
          <RibbonButton icon={Copy} label="Copiar" title="Copiar (Ctrl+C)" onClick={editor.copy} />
          <RibbonButton
            icon={Paintbrush}
            label="Pincel de Formatação"
            title={withShortcut('Pincel de Formatação', 'copyFormat')}
            active={active('copy-format')}
            disabled={disabled('copy-format')}
            onClick={run('copy-format')}
          />
        </Rows>
      </RibbonGroup>

      <RibbonGroup label="Fonte">
        <Rows>
          <Row>
            <ComboBox
              title="Fonte"
              width={150}
              value={family}
              options={families.map((f) => ({ label: f, value: f, style: { fontFamily: `"${f}"` } }))}
              onCommit={(f) => {
                execute('font-family', f);
                editor.focus();
              }}
            />
            <ComboBox
              title="Tamanho da Fonte"
              width={52}
              value={size == null ? '' : formatSize(size)}
              options={FONT_SIZES.map((s) => ({ label: formatSize(s), value: String(s) }))}
              onCommit={(v) => {
                const n = parseFloat(v.replace(',', '.'));
                if (Number.isFinite(n) && n >= 1 && n <= 1638) setSize(Math.round(n * 2) / 2);
                editor.focus();
              }}
            />
            <RibbonButton
              icon={AArrowUp}
              title={withShortcut('Aumentar Tamanho da Fonte', 'growFont')}
              disabled={disabled('font-size')}
              onClick={() => setSize(stepFontSize(size ?? 11, true))}
            />
            <RibbonButton
              icon={AArrowDown}
              title={withShortcut('Diminuir Tamanho da Fonte', 'shrinkFont')}
              disabled={disabled('font-size')}
              onClick={() => setSize(stepFontSize(size ?? 11, false))}
            />
            <RibbonButton
              icon={RemoveFormatting}
              title={withShortcut('Limpar Toda a Formatação', 'clearFormatting')}
              disabled={disabled('clear-formatting')}
              onClick={run('clear-formatting')}
            />
          </Row>
          <Row>
            <RibbonButton icon={Bold} title={withShortcut('Negrito', 'bold')} active={active('bold')} disabled={disabled('bold')} onClick={run('bold')} />
            <RibbonButton icon={Italic} title={withShortcut('Itálico', 'italic')} active={active('italic')} disabled={disabled('italic')} onClick={run('italic')} />
            <RibbonButton
              icon={Underline}
              title={withShortcut('Sublinhado', 'underline')}
              active={active('underline')}
              disabled={disabled('underline')}
              onClick={run('underline')}
            />
            <RibbonButton
              icon={Strikethrough}
              title="Tachado"
              active={active('strikethrough')}
              disabled={disabled('strikethrough')}
              onClick={editor.toggleStrike}
            />
            <RibbonButton
              icon={Subscript}
              title={withShortcut('Subscrito', 'subscript')}
              active={editor.vertAlign() === 'subscript'}
              onClick={() => editor.toggleVertAlign('subscript')}
            />
            <RibbonButton
              icon={Superscript}
              title={withShortcut('Sobrescrito', 'superscript')}
              active={editor.vertAlign() === 'superscript'}
              onClick={() => editor.toggleVertAlign('superscript')}
            />
            <span className="ribbon-sep" />
            <MenuButton
              title="Cor do Realce do Texto"
              glyph={
                <ColorGlyph color={highlight}>
                  <Highlighter size={15} strokeWidth={1.8} />
                </ColorGlyph>
              }
              disabled={disabled('highlight-color')}
              onClick={() => execute('highlight-color', highlight)}
              menu={(close) => (
                <ColorGrid
                  rows={HIGHLIGHT_ROWS}
                  noneLabel="Sem Cor"
                  close={close}
                  onPick={(c) => {
                    if (c) setHighlight(c);
                    execute('highlight-color', c ?? 'none');
                  }}
                />
              )}
            />
            <MenuButton
              title="Cor da Fonte"
              glyph={
                <ColorGlyph color={fontColor}>
                  <span className="glyph-letter">A</span>
                </ColorGlyph>
              }
              disabled={disabled('text-color')}
              onClick={() => execute('text-color', fontColor)}
              menu={(close) => (
                <ColorGrid
                  rows={THEME_COLOR_ROWS}
                  standard={STANDARD_COLORS}
                  noneLabel="Automático"
                  close={close}
                  onPick={(c) => {
                    if (c) setFontColor(c);
                    execute('text-color', c ?? 'none');
                  }}
                />
              )}
            />
          </Row>
        </Rows>
      </RibbonGroup>

      <RibbonGroup label="Parágrafo">
        <Rows>
          <Row>
            <RibbonButton
              icon={List}
              title="Marcadores"
              active={active('bullet-list')}
              disabled={disabled('bullet-list')}
              onClick={run('bullet-list')}
            />
            <RibbonButton
              icon={ListOrdered}
              title="Numeração"
              active={active('numbered-list')}
              disabled={disabled('numbered-list')}
              onClick={run('numbered-list')}
            />
            <RibbonButton
              icon={ListIndentDecrease}
              title={withShortcut('Diminuir Recuo', 'outdent')}
              disabled={disabled('indent-decrease')}
              onClick={() => (props.page.paragraph()?.isList ? execute('indent-decrease') : props.page.stepIndent(-1))}
            />
            <RibbonButton
              icon={ListIndentIncrease}
              title={withShortcut('Aumentar Recuo', 'indent')}
              disabled={disabled('indent-increase')}
              onClick={() => (props.page.paragraph()?.isList ? execute('indent-increase') : props.page.stepIndent(1))}
            />
            <span className="ribbon-sep" />
            <RibbonButton
              icon={Pilcrow}
              title={withShortcut('Mostrar Tudo', 'showAll')}
              active={active('formatting-marks')}
              onClick={run('formatting-marks')}
            />
          </Row>
          <Row>
            <RibbonButton icon={TextAlignStart} title={withShortcut('Alinhar à Esquerda', 'alignLeft')} active={align === 'left'} disabled={disabled('text-align')} onClick={run('text-align', 'left')} />
            <RibbonButton icon={TextAlignCenter} title={withShortcut('Centralizar', 'alignCenter')} active={align === 'center'} disabled={disabled('text-align')} onClick={run('text-align', 'center')} />
            <RibbonButton icon={TextAlignEnd} title={withShortcut('Alinhar à Direita', 'alignRight')} active={align === 'right'} disabled={disabled('text-align')} onClick={run('text-align', 'right')} />
            <RibbonButton icon={TextAlignJustify} title={withShortcut('Justificar', 'justify')} active={align === 'justify'} disabled={disabled('text-align')} onClick={run('text-align', 'justify')} />
            <span className="ribbon-sep" />
            <MenuButton
              icon={ListChevronsUpDown}
              title="Espaçamento de Linha e Parágrafo"
              disabled={disabled('line-height')}
              menu={(close) =>
                LINE_SPACINGS.map((v) => (
                  <MenuItem
                    key={v}
                    label={v.toFixed(v === 1.15 ? 2 : 1).replace('.', ',')}
                    checked={Math.abs(v - lineHeight) < 0.01}
                    close={close}
                    onSelect={() => execute('line-height', v)}
                  />
                ))
              }
            />
          </Row>
        </Rows>
      </RibbonGroup>

      <RibbonGroup label="Estilos" className="ribbon-group-styles">
        <StylesGallery
          styles={editor.quickStyles()}
          activeId={String(state('linked-style')?.value ?? 'Normal')}
          disabled={disabled('linked-style')}
          onApply={(style) => execute('linked-style', style)}
        />
      </RibbonGroup>

      <RibbonGroup label="Edição">
        <Rows>
          <RibbonButton icon={Search} label="Localizar" title={withShortcut('Localizar', 'find')} onClick={props.onFind} />
          <RibbonButton icon={Replace} label="Substituir" title={withShortcut('Substituir', 'replace')} onClick={props.onReplace} />
          <RibbonButton icon={SquareMousePointer} label="Selecionar Tudo" title={withShortcut('Selecionar Tudo', 'selectAll')} onClick={editor.selectAll} />
        </Rows>
      </RibbonGroup>
    </>
  );
}

function InsertTab(props: RibbonProps) {
  const { disabled, run } = useCommand(props);
  const { editor, execute } = props;
  return (
    <>
      <RibbonGroup label="Páginas">
        <RibbonButton large icon={SeparatorHorizontal} label="Quebra de Página" title={withShortcut('Quebra de Página', 'pageBreak')} onClick={editor.insertPageBreak} />
      </RibbonGroup>
      <RibbonGroup label="Tabelas">
        <MenuButton
          large
          icon={Table}
          label="Tabela"
          title="Adicionar uma Tabela"
          disabled={disabled('table-insert')}
          menu={(close) => <TableGrid close={close} onPick={(rows, cols) => execute('table-insert', { rows, cols })} />}
        />
      </RibbonGroup>
      <RibbonGroup label="Ilustrações">
        <RibbonButton large icon={Image} label="Imagens" title="Inserir Imagens" disabled={disabled('image')} onClick={run('image')} />
      </RibbonGroup>
      <RibbonGroup label="Links">
        <RibbonButton large icon={Link} label="Link" title={withShortcut('Inserir Link', 'link')} disabled={disabled('link')} onClick={editor.insertLink} />
      </RibbonGroup>
      <RibbonGroup label="Sumário">
        <RibbonButton
          large
          icon={TableOfContents}
          label="Sumário"
          title="Inserir Sumário"
          disabled={disabled('table-of-contents-insert')}
          onClick={run('table-of-contents-insert')}
        />
      </RibbonGroup>
    </>
  );
}

/** Grade de "Inserir Tabela" do Word: passe o mouse para escolher linhas × colunas. */
function TableGrid(props: { close: () => void; onPick: (rows: number, cols: number) => void }) {
  const [hover, setHover] = useState({ r: 0, c: 0 });
  return (
    <div className="table-grid">
      <div className="table-grid-title">{hover.r ? `Tabela ${hover.c}x${hover.r}` : 'Inserir Tabela'}</div>
      {Array.from({ length: 8 }, (_, r) => (
        <div key={r} className="table-grid-row">
          {Array.from({ length: 10 }, (_, c) => (
            <button
              key={c}
              type="button"
              className={`table-grid-cell${r < hover.r && c < hover.c ? ' is-on' : ''}`}
              aria-label={`${c + 1} colunas, ${r + 1} linhas`}
              onMouseEnter={() => setHover({ r: r + 1, c: c + 1 })}
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => {
                props.close();
                props.onPick(r + 1, c + 1);
              }}
            />
          ))}
        </div>
      ))}
    </div>
  );
}

function ReviewTab(props: RibbonProps) {
  const { state, disabled, run } = useCommand(props);
  const { execute } = props;
  const mode = state('document-mode')?.value;
  const tracking = mode === 'suggesting';
  return (
    <>
      <RibbonGroup label="Controle">
        <RibbonButton
          large
          icon={FilePenLine}
          label="Controlar Alterações"
          title="Controlar Alterações"
          active={tracking}
          disabled={disabled('document-mode')}
          onClick={() => execute('document-mode', tracking ? 'editing' : 'suggesting')}
        />
      </RibbonGroup>
      <RibbonGroup label="Alterações">
        <RibbonButton
          large
          icon={CircleCheck}
          label="Aceitar"
          title="Aceitar a alteração selecionada"
          disabled={disabled('track-changes-accept-selection')}
          onClick={run('track-changes-accept-selection')}
        />
        <RibbonButton
          large
          icon={CircleX}
          label="Rejeitar"
          title="Rejeitar a alteração selecionada"
          disabled={disabled('track-changes-reject-selection')}
          onClick={run('track-changes-reject-selection')}
        />
      </RibbonGroup>
    </>
  );
}

function ViewTab(props: RibbonProps) {
  const { state, active, run } = useCommand(props);
  return (
    <>
      <RibbonGroup label="Mostrar">
        <Rows>
          <RibbonButton icon={Ruler} label="Régua" title="Régua" active={props.showRuler} onClick={props.onToggleRuler} />
          <RibbonButton
            icon={Pilcrow}
            label="Marcas de Formatação"
            title="Mostrar marcas de parágrafo e outros símbolos ocultos"
            active={active('formatting-marks')}
            onClick={run('formatting-marks')}
          />
        </Rows>
      </RibbonGroup>
      <RibbonGroup label="Zoom">
        <MenuButton
          large
          icon={ZoomIcon}
          label="Zoom"
          title="Zoom"
          menu={(close) =>
            [50, 75, 100, 125, 150, 200].map((z) => (
              <MenuItem
                key={z}
                label={`${z}%`}
                checked={state('zoom')?.value === z}
                close={close}
                onSelect={() => props.execute('zoom', z)}
              />
            ))
          }
        />
        <RibbonButton large label="100%" title="Zoom de 100%" glyph={<span className="glyph-100">100</span>} onClick={run('zoom', 100)} />
        <RibbonButton
          large
          icon={MoveHorizontal}
          label="Largura da Página"
          title="Ajustar à largura da página"
          onClick={props.onFitWidth}
        />
      </RibbonGroup>
    </>
  );
}
