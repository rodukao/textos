import {
  ArrowDownToLine,
  ArrowUpToLine,
  Columns2,
  FileText,
  ListIndentDecrease,
  ListIndentIncrease,
  RectangleHorizontal,
  SeparatorHorizontal,
  SquareDashed,
} from 'lucide-react';
import { MenuButton, MenuItem, RibbonGroup, Rows, SpinBox } from './controls';
import { CM_PER_INCH, TWIPS_PER_INCH, TWIPS_PER_PT, type PageActions } from './page-actions';
import { MARGIN_PRESETS, PAPER_SIZES, formatDecimal, parseDecimal } from './word-data';

export type AskFields = (title: string, fields: { label: string; initial: string }[]) => Promise<string[] | null>;

const near = (a: number, b: number, tolerance = 0.03) => Math.abs(a - b) <= tolerance;
const cm = (inches: number) => inches * CM_PER_INCH;
const toInches = (centimeters: number) => centimeters / CM_PER_INCH;
const twipsToCm = (twips: number) => (twips / TWIPS_PER_INCH) * CM_PER_INCH;
const cmToTwips = (centimeters: number) => (centimeters / CM_PER_INCH) * TWIPS_PER_INCH;

/**
 * Guia Layout do Word: Configurar Página (Margens, Orientação, Tamanho, Colunas, Quebras) e
 * Parágrafo (Recuar e Espaçamento), na seção e nos parágrafos do cursor.
 */
export function LayoutTab(props: { page: PageActions; ask: AskFields; insertPageBreak: () => void }) {
  const { page, ask } = props;
  const section = page.section();
  const paragraph = page.paragraph();
  const m = section?.margins;

  const customMargins = async () => {
    if (!m) return;
    const values = await ask('Margens Personalizadas', [
      { label: 'Superior (cm):', initial: formatDecimal(cm(m.top)) },
      { label: 'Inferior (cm):', initial: formatDecimal(cm(m.bottom)) },
      { label: 'Esquerda (cm):', initial: formatDecimal(cm(m.left)) },
      { label: 'Direita (cm):', initial: formatDecimal(cm(m.right)) },
    ]);
    if (!values) return;
    const [top, bottom, left, right] = values.map((v) => parseDecimal(v));
    const valid = (n: number | null) => n != null && n >= 0 && n < 50;
    if (![top, bottom, left, right].every(valid)) return;
    page.setMargins({ top: toInches(top!), bottom: toInches(bottom!), left: toInches(left!), right: toInches(right!) });
  };

  const landscape = section?.orientation === 'landscape';
  const portraitSize = section ? [Math.min(section.width, section.height), Math.max(section.width, section.height)] : null;

  return (
    <>
      <RibbonGroup label="Configurar Página">
        <MenuButton
          large
          icon={SquareDashed}
          label="Margens"
          title="Margens"
          disabled={!section}
          menu={(close) => (
            <div className="menu-list">
              {MARGIN_PRESETS.map((p) => (
                <MenuItem
                  key={p.name}
                  label={`${p.name} — Sup. ${formatDecimal(p.top)} · Inf. ${formatDecimal(p.bottom)} · Esq. ${formatDecimal(p.left)} · Dir. ${formatDecimal(p.right)} cm`}
                  checked={Boolean(m && near(cm(m.top), p.top) && near(cm(m.bottom), p.bottom) && near(cm(m.left), p.left) && near(cm(m.right), p.right))}
                  close={close}
                  onSelect={() =>
                    page.setMargins({ top: toInches(p.top), bottom: toInches(p.bottom), left: toInches(p.left), right: toInches(p.right) })
                  }
                />
              ))}
              <div className="menu-sep" />
              <MenuItem label="Margens Personalizadas..." close={close} onSelect={() => void customMargins()} />
            </div>
          )}
        />
        <MenuButton
          large
          icon={RectangleHorizontal}
          label="Orientação"
          title="Orientação"
          disabled={!section}
          menu={(close) => (
            <>
              <MenuItem label="Retrato" checked={!landscape} close={close} onSelect={() => page.setOrientation('portrait')} />
              <MenuItem label="Paisagem" checked={landscape} close={close} onSelect={() => page.setOrientation('landscape')} />
            </>
          )}
        />
        <MenuButton
          large
          icon={FileText}
          label="Tamanho"
          title="Tamanho"
          disabled={!section}
          menu={(close) => (
            <div className="menu-list">
              {PAPER_SIZES.map((p) => (
                <MenuItem
                  key={p.name}
                  label={`${p.name} — ${formatDecimal(p.width)} cm x ${formatDecimal(p.height)} cm`}
                  checked={Boolean(portraitSize && near(cm(portraitSize[0]), p.width, 0.1) && near(cm(portraitSize[1]), p.height, 0.1))}
                  close={close}
                  onSelect={() => page.setPaperSize(toInches(p.width), toInches(p.height))}
                />
              ))}
            </div>
          )}
        />
        <MenuButton
          large
          icon={Columns2}
          label="Colunas"
          title="Colunas"
          disabled={!section}
          menu={(close) =>
            (['Um', 'Dois', 'Três'] as const).map((label, i) => (
              <MenuItem key={label} label={label} checked={section?.columns === i + 1} close={close} onSelect={() => page.setColumns(i + 1)} />
            ))
          }
        />
        <MenuButton
          large
          icon={SeparatorHorizontal}
          label="Quebras"
          title="Adicionar Quebras de Página e de Seção"
          menu={(close) => (
            <div className="menu-list">
              <div className="menu-heading">Quebras de Página</div>
              <MenuItem label="Página" close={close} onSelect={props.insertPageBreak} />
            </div>
          )}
        />
      </RibbonGroup>

      <RibbonGroup label="Parágrafo">
        <Rows>
          <div className="ribbon-caption">Recuar</div>
          <SpinBox
            label="Esquerda:"
            icon={ListIndentIncrease}
            title="Recuo à Esquerda"
            unit="cm"
            step={0.1}
            min={-50}
            max={50}
            value={paragraph ? twipsToCm(paragraph.left) : null}
            format={(n) => formatDecimal(n)}
            parse={parseDecimal}
            onCommit={(v) => page.setIndentation({ left: cmToTwips(v) })}
          />
          <SpinBox
            label="Direita:"
            icon={ListIndentDecrease}
            title="Recuo à Direita"
            unit="cm"
            step={0.1}
            min={-50}
            max={50}
            value={paragraph ? twipsToCm(paragraph.right) : null}
            format={(n) => formatDecimal(n)}
            parse={parseDecimal}
            onCommit={(v) => page.setIndentation({ right: cmToTwips(v) })}
          />
        </Rows>
        <Rows>
          <div className="ribbon-caption">Espaçamento</div>
          <SpinBox
            label="Antes:"
            icon={ArrowUpToLine}
            title="Espaçamento Antes"
            unit="pt"
            step={6}
            min={0}
            max={1584}
            value={paragraph ? paragraph.before / TWIPS_PER_PT : null}
            format={(n) => formatDecimal(n, 1)}
            parse={parseDecimal}
            onCommit={(v) => page.setSpacing({ before: v * TWIPS_PER_PT })}
          />
          <SpinBox
            label="Depois:"
            icon={ArrowDownToLine}
            title="Espaçamento Depois"
            unit="pt"
            step={6}
            min={0}
            max={1584}
            value={paragraph ? paragraph.after / TWIPS_PER_PT : null}
            format={(n) => formatDecimal(n, 1)}
            parse={parseDecimal}
            onCommit={(v) => page.setSpacing({ after: v * TWIPS_PER_PT })}
          />
        </Rows>
      </RibbonGroup>
    </>
  );
}
