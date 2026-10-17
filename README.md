# Textos

Editor de textos de código aberto, compatível com o Microsoft Word. Abre e salva arquivos `.docx`
preservando a formatação, e roda no **navegador** e como app **desktop** (Windows, Linux e macOS)
a partir do mesmo código.

> Projeto da RoduApps, em desenvolvimento.

## Objetivo

Abrir um `.docx` e vê-lo exatamente como no Word, editar e salvar **sem perder nenhuma formatação**.
O progresso é medido de forma objetiva por uma [bancada de fidelidade](#bancada-de-fidelidade-com-o-word)
que usa o próprio Word como referência.

## Stack

| Camada | Tecnologia |
|---|---|
| Motor de documentos (`.docx`) | [textos-engine](https://github.com/rodukao/textos-engine) — nosso fork AGPL do SuperDoc 1.x (OOXML nativo, ProseMirror, paginação) |
| Interface | React 19 + TypeScript + Vite |
| Desktop | [Tauri 2](https://tauri.app) (Rust + WebView do sistema) |

## Como rodar

```bash
npm install
npm run dev            # versão web em http://localhost:1420
npm run desktop        # app desktop (precisa do Rust — veja abaixo)
npm run build          # build web em dist/
npm run desktop:build  # instalador desktop em src-tauri/target/release/bundle/
```

### Pré-requisitos do desktop

- **Rust**: https://rustup.rs
- **Windows**: Visual Studio Build Tools com C++ e Windows SDK (o WebView2 já vem no Windows 11)
- **Linux**: `webkit2gtk-4.1` e dependências — https://tauri.app/start/prerequisites/
- **macOS**: Xcode Command Line Tools (`xcode-select --install`)

## Estrutura

```
src/
  App.tsx                   documento aberto, ações de arquivo e atalhos
  editor/DocumentEditor.tsx monta o SuperDoc (barra de ferramentas + folhas)
  engine/superdoc-config.ts configuração do motor, compartilhada com a bancada
  components/               FileBar, StatusBar
  document/format.ts        filtros de arquivo (.docx)
  platform/                 ÚNICO lugar que sabe se é web ou desktop
    types.ts                interface Platform
    web.ts                  File System Access API (+ fallback download/upload)
    tauri.ts                diálogos e arquivos nativos via plugins do Tauri
src-tauri/                  projeto Rust do Tauri (janela, permissões, ícones)
bench/                      bancada de fidelidade com o Word
```

## Privacidade

Os documentos nunca saem da máquina. O SuperDoc envia telemetria por padrão; o Textos a desliga
(`src/engine/superdoc-config.ts`).

## Bancada de fidelidade com o Word

Mede o quanto o Textos se aproxima do Microsoft Word, usando o **Word instalado como referência**
(Windows, via automação COM). Além da telemetria desligada, a bancada bloqueia qualquer acesso à
rede fora do servidor local.

```bash
npm run bench:samples                 # gera documentos de exemplo no Word (bench/corpus)
npm run bench                         # roda a bancada em bench/corpus
node bench/run.mjs "C:\pasta\docs"    # roda em outra pasta (inclui subpastas)
```

Resultado em `bench/out/report.html` (fora do controle de versão). Para cada `.docx`:

- **Salvar sem perda** — o Textos abre e salva sem editar; o Word abre o arquivo salvo e gera um PDF,
  comparado página a página com o PDF do original. Também compara o XML interno e lista o que sumiu.
- **Exibição** — a página desenhada pelo Textos comparada com a do Word, tolerando suavização de
  fonte e deslocamentos de até 2 px. Na imagem de diferenças: vermelho = só no Word,
  azul = só no Textos, laranja = cor diferente.
- **Desempenho** — tempo para abrir num dispositivo simulado com CPU 4× mais lenta, e memória usada.
  O Textos deve continuar rápido em máquinas modestas: nenhuma melhoria de fidelidade pode piorar isso.

## Motor (textos-engine)

O motor fica em um repositório próprio, [rodukao/textos-engine](https://github.com/rodukao/textos-engine)
(branch `textos`), e o Textos instala o pacote publicado nas *Releases* dele. Para alterar o motor:

1. clone o `textos-engine` ao lado deste repositório e faça a mudança na branch `textos`;
2. gere o pacote (`pnpm run build:superdoc` e `pnpm --dir packages/superdoc pack`);
3. teste no Textos com `npm install <caminho do .tgz>` e rode a bancada;
4. publique uma nova *Release* no `textos-engine` e aponte o `package.json` do Textos para ela.

### Ferramentas de diagnóstico (`bench/tools/`)

Para investigar uma diferença apontada pela bancada, comparando o `.docx` original com o
`roundtrip.docx` salvo pelo Textos (em `bench/out/<documento>/`):

```bash
node bench/tools/xml-ao-redor.mjs <arquivo.docx> "trecho do texto"      # estrutura XML em volta do trecho
node bench/tools/paragrafos-xml.mjs <arquivo.docx> "trecho do texto" 3  # parágrafos do corpo a partir do trecho
node bench/tools/ancora-xml.mjs <arquivo.docx> "trecho"                 # propriedades da imagem/caixa ancorada
node bench/tools/caixa-de-texto-xml.mjs <arquivo.docx> "trecho"         # conteúdo da caixa de texto
node bench/tools/estilo-xml.mjs <arquivo.docx> <styleId>                # definição de um estilo e padrões
node bench/tools/ancoras.mjs <arquivo.docx>                             # objetos flutuantes de cada parágrafo
node bench/tools/inspecionar.mjs <arquivo.docx> <script.js>             # roda JS na página diagramada (DOM)
```

Para documentos reais que não devem ir para o repositório, use uma pasta dentro de `bench/out-reais/`
(fora do controle de versão) e rode `node bench/run.mjs bench/out-reais/<pasta>`.

Para a altura de linha, que no Word depende das métricas de cada fonte:

```bash
node bench/tools/metricas-fonte.mjs C:\Windows\Fonts\calibri.ttf        # altura simples da fonte (× tamanho)
powershell -File bench/tools/altura-linha-word.ps1 -Fontes "Calibri,Cambria" -Linhas "240,259"  # medida no Word
```

O motor usa a altura simples do Word (métricas GDI da fonte, não 1,15 × tamanho) vezes o múltiplo,
sem arredondar: na tela o Word arredonda cada linha para cima em 1/960 de polegada (é o que
`altura-linha-word.ps1` mede), mas ao imprimir ou salvar em PDF ele usa a altura exata, e as quebras
de página seguem esta. A tabela de fontes fica em
`packages/layout-engine/measuring/dom/src/wordLineMetrics.ts` no `textos-engine`.

### Situação atual

**Salvar:** todos os documentos testados (reais e de exemplo) voltam **idênticos no Word** ao
salvar sem edição — 0% de diferença em todas as páginas.

**Altura de linha:** igual à do Word em espaçamento simples, múltiplo, "pelo menos" e em parágrafos
vazios; as quebras de página do documento de exemplo com várias fontes e espaçamentos coincidem.

**Larguras de texto:** o texto é medido e desenhado com as larguras de projeto da fonte, como no
Word (no Windows a 100%, o navegador usaria larguras ajustadas pelo *hinting*, até ~1,8% maiores em
Cambria). Custo atual: abrir fica de 2% a 6% mais lento; otimizar depois.

**Quebras de página:** controle de linhas viúvas e órfãs como no Word (ligado a menos que o
documento o desligue); nos documentos de exemplo as quebras coincidem com as do PDF do Word.

**Objetos flutuantes:** o texto contorna imagens e caixas linha a linha, e cada objeto fica na
página do seu ponto de âncora, como no Word (inclusive quebrando a página antes da linha da âncora
quando o Word faz isso).

Pendências conhecidas, por impacto:

1. **exibição:** posição do texto dentro da linha em espaçamento múltiplo — o Word deixa o espaço
   extra abaixo do texto, o Textos o divide acima e abaixo (≈ 5 px em Times a 1,5)
2. **salvar:** a cópia de compatibilidade com o Word 2007 (`mc:Fallback`) de caixas de texto e as
   marcas de revisão ortográfica (`w:proofErr`) não são mantidas (sem efeito visual no Word atual)

## Atenção: não atualize para o SuperDoc 2.x

A partir de certa versão 2.x, o SuperDoc passou a depender de `@superdoc/docx-engine`, um pacote
**proprietário** cuja licença proíbe, entre outras coisas, usá-lo em editores de documentos
concorrentes e redistribuí-lo. Isso é incompatível com o Textos. Por isso o motor é um fork da
linha 1.x (100% AGPL), e nada do 2.x é usado como referência. Antes de adicionar ou atualizar
qualquer dependência, confira as licenças de **toda** a árvore, incluindo as dependências indiretas.

## Atalhos

| Ação | Atalho |
|---|---|
| Novo | Ctrl+Alt+N |
| Abrir | Ctrl+O |
| Salvar / Salvar como | Ctrl+S / Ctrl+Shift+S |
| Imprimir / PDF | Ctrl+P |

(No macOS, use Cmd no lugar de Ctrl.)

## Licença

[GNU AGPL-3.0](LICENSE) — a mesma do SuperDoc, do qual o Textos depende.
