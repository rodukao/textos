<#
  Gera documentos de exemplo com o próprio Word, cobrindo vários tipos de formatação.
  Uso: powershell -File bench/make-samples.ps1 -OutDir bench/corpus
       powershell -File bench/make-samples.ps1 -OutDir bench/corpus -Only 04   (só os documentos indicados)
#>
param([Parameter(Mandatory = $true)][string]$OutDir, [string[]]$Only)

$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force $OutDir | Out-Null
$OutDir = (Resolve-Path $OutDir).Path
$image = (Resolve-Path (Join-Path $PSScriptRoot '..\src-tauri\icons\128x128.png')).Path

$wdFormatDocumentDefault = 16
$wdDoNotSaveChanges = 0

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0

$logFile = Join-Path $env:TEMP 'textos-make-samples.log'
function Log($m) { Add-Content -Encoding UTF8 $logFile "$(Get-Date -Format 'HH:mm:ss.fff') $m" }

function New-Doc { $word.Documents.Add() }
function Save-Doc($doc, $name) {
  if ($Only -and -not ($Only | Where-Object { $name.StartsWith($_) })) { $doc.Close([ref]$wdDoNotSaveChanges); return }
  # [string]: o Join-Path devolve um PSObject, e o COM do Word trava ao recebê-lo no SaveAs2.
  [string]$path = Join-Path $OutDir $name
  Log "salvando $path"
  $doc.SaveAs2($path, [int]$wdFormatDocumentDefault)
  Log "  fechando"
  $doc.Close([ref]$wdDoNotSaveChanges)
  Log "  ok"
  Write-Output "gerado: $name"
}
# Insere um parágrafo no fim do documento e devolve o Range dele.
function Add-Para($doc, [string]$text, $style = $null) {
  $r = $doc.Content
  $r.InsertParagraphAfter()
  $p = $doc.Paragraphs.Last.Range
  $p.Text = $text
  $p = $doc.Paragraphs.Last.Range
  # Constantes WdBuiltinStyle (ex.: -2 = Título 1) independem do idioma do Word.
  if ($null -ne $style) { $p.Style = $doc.Styles.Item([int]$style) }
  return $p
}

try {
  # ---------- 1. Formatação de caractere e parágrafo, listas, notas ----------
  $d = New-Doc
  $d.Paragraphs.First.Range.Text = 'Relatório de Formatação'
  $d.Paragraphs.First.Range.Style = $d.Styles.Item(-63)   # wdStyleTitle
  $null = Add-Para $d 'Introdução' -2                     # wdStyleHeading1
  $p = Add-Para $d 'Texto normal com negrito, itálico, sublinhado duplo, tachado, cor, realce, fonte Georgia 14, sobrescrito, versalete e espaçamento entre caracteres.' -1
  $t = $p.Text
  function Fmt($word_, $action) { $i = $t.IndexOf($word_); $r = $d.Range($p.Start + $i, $p.Start + $i + $word_.Length); & $action $r }
  Fmt 'negrito' { param($r) $r.Font.Bold = $true }
  Fmt 'itálico' { param($r) $r.Font.Italic = $true }
  Fmt 'sublinhado duplo' { param($r) $r.Font.Underline = 3 }
  Fmt 'tachado' { param($r) $r.Font.StrikeThrough = $true }
  Fmt 'cor' { param($r) $r.Font.Color = 0x2F75B5 }
  Fmt 'realce' { param($r) $r.HighlightColorIndex = 7 }
  Fmt 'fonte Georgia 14' { param($r) $r.Font.Name = 'Georgia'; $r.Font.Size = 14 }
  Fmt 'sobrescrito' { param($r) $r.Font.Superscript = $true }
  Fmt 'versalete' { param($r) $r.Font.SmallCaps = $true }
  Fmt 'espaçamento entre caracteres' { param($r) $r.Font.Spacing = 2 }

  $p = Add-Para $d 'Parágrafo justificado com recuo de primeira linha de 1,25 cm, espaçamento 1,5 entre linhas e 12 pt depois. Este texto é longo o bastante para ocupar várias linhas e mostrar a justificação funcionando como no Word, distribuindo os espaços entre as palavras.' -1
  $p.ParagraphFormat.Alignment = 3
  $p.ParagraphFormat.FirstLineIndent = $word.CentimetersToPoints(1.25)
  $p.ParagraphFormat.LineSpacingRule = 1
  $p.ParagraphFormat.SpaceAfter = 12

  $p = Add-Para $d "Coluna A`tColuna B`tValor" -1
  $p.ParagraphFormat.TabStops.ClearAll()
  $null = $p.ParagraphFormat.TabStops.Add($word.CentimetersToPoints(5), 0)
  $null = $p.ParagraphFormat.TabStops.Add($word.CentimetersToPoints(14), 2, 1)   # direita, pontilhado

  $null = Add-Para $d 'Listas' -2
  $first = $null
  foreach ($item in 'Primeiro item', 'Segundo item', 'Subitem A', 'Subitem B', 'Terceiro item') {
    $p = Add-Para $d $item -1
    if (-not $first) { $first = $p.Start }
  }
  $listRange = $d.Range($first, $d.Paragraphs.Last.Range.End)
  $tmpl = $word.ListGalleries.Item(3).ListTemplates.Item(1)   # wdOutlineNumberGallery
  $listRange.ListFormat.ApplyListTemplate($tmpl, $false, 0)
  foreach ($i in 3, 4) { $d.Paragraphs.Item($d.Paragraphs.Count - 5 + $i).Range.ListFormat.ListIndent() }

  $p = Add-Para $d 'Lista com marcadores' -1
  $p.ListFormat.ApplyBulletDefault()
  $p = Add-Para $d 'Segundo marcador' -1

  $p = Add-Para $d 'Parágrafo com nota de rodapé e um link para o site do Tauri.' -1
  $p.ListFormat.RemoveNumbers()
  $null = $d.Footnotes.Add($d.Range($p.End - 1, $p.End - 1), [Type]::Missing, 'Esta é uma nota de rodapé.')
  $i = $p.Text.IndexOf('site do Tauri')
  $null = $d.Hyperlinks.Add($d.Range($p.Start + $i, $p.Start + $i + 13), 'https://tauri.app')
  Save-Doc $d '01-formatacao.docx'

  # ---------- 2. Tabela, cabeçalho/rodapé, seções, colunas ----------
  $d = New-Doc
  $d.Paragraphs.First.Range.Text = 'Tabela e seções'
  $d.Paragraphs.First.Range.Style = $d.Styles.Item(-2)
  $p = Add-Para $d '' -1
  $tbl = $d.Tables.Add($p, 4, 3)
  $tbl.Borders.Enable = $true
  # O nome do estilo depende do idioma do Word.
  foreach ($styleName in 'Tabela de Grade 4 - Ênfase 1', 'Grid Table 4 - Accent 1') {
    try { $tbl.Style = $styleName; break } catch {}
  }
  $data = @(('Produto', 'Qtd', 'Preço'), ('Caneta', '10', 'R$ 2,50'), ('Caderno', '3', 'R$ 18,90'), ('Total', '', 'R$ 81,70'))
  for ($r = 1; $r -le 4; $r++) { for ($c = 1; $c -le 3; $c++) { $tbl.Cell($r, $c).Range.Text = $data[$r - 1][$c - 1] } }
  $tbl.Cell(4, 1).Merge($tbl.Cell(4, 2))
  $tbl.Cell(2, 3).Shading.BackgroundPatternColor = 0xCCF2FF
  $tbl.Rows.Item(1).HeadingFormat = $true

  $sec = $d.Sections.Item(1)
  $sec.Headers.Item(1).Range.Text = 'Cabeçalho do documento — RoduApps'
  $sec.Headers.Item(1).Range.ParagraphFormat.Alignment = 2
  $f = $sec.Footers.Item(1).Range
  $f.Text = 'Página '
  $f.Collapse(0)
  $null = $d.Fields.Add($f, 33)   # wdFieldPage
  $sec.Footers.Item(1).Range.ParagraphFormat.Alignment = 1

  $end = $d.Content
  $end.Collapse(0)
  $end.InsertBreak(2)   # wdSectionBreakNextPage
  $s2 = $d.Sections.Item(2)
  $s2.PageSetup.Orientation = 1   # paisagem
  $s2.PageSetup.TextColumns.SetCount(2)
  $txt = ('Texto em duas colunas numa seção em paisagem. ' * 40).Trim()
  $null = Add-Para $d $txt -1
  Save-Doc $d '02-tabela-secoes.docx'

  # ---------- 3. Imagem, caixa de texto, comentário, controle de alterações ----------
  $d = New-Doc
  $d.Paragraphs.First.Range.Text = 'Revisão e objetos'
  $d.Paragraphs.First.Range.Style = $d.Styles.Item(-2)
  $p = Add-Para $d 'Imagem embutida na linha:' -1
  $p = Add-Para $d '' -1
  $null = $d.InlineShapes.AddPicture($image, $false, $true, $p)
  $p = Add-Para $d 'Este parágrafo tem um comentário.' -1
  $null = $d.Comments.Add($p, 'Comentário de teste da bancada.')
  $box = $d.Shapes.AddTextbox(1, 300, 200, 180, 60)
  $box.TextFrame.TextRange.Text = 'Caixa de texto flutuante'
  $d.TrackRevisions = $true
  $p = Add-Para $d 'Texto inserido com controle de alterações ativo.' -1
  $d.TrackRevisions = $false
  Save-Doc $d '03-revisao-objetos.docx'

  # ---------- 4. Espaçamento entre linhas múltiplo em várias páginas ----------
  # A altura da linha depende da fonte (Calibri, Cambria, Times) e do múltiplo; um erro pequeno
  # por linha se acumula e muda a quebra de página.
  $texto = 'O espaçamento entre linhas do Word depende das métricas de cada fonte, e não apenas do tamanho: ' +
    'a mesma configuração de 1,08 gera linhas de alturas diferentes em Calibri e em Cambria. Um erro de ' +
    'frações de ponto por linha se acumula ao longo da página e muda onde cada página termina.'
  $d = New-Doc
  $d.Paragraphs.First.Range.Text = 'Espaçamento múltiplo'
  $d.Paragraphs.First.Range.Style = $d.Styles.Item(-2)
  foreach ($bloco in @(
      @{ Fonte = 'Calibri'; Tamanho = 11; Linha = 1.08; Depois = 8 },
      @{ Fonte = 'Cambria'; Tamanho = 11; Linha = 1.08; Depois = 8 },
      @{ Fonte = 'Times New Roman'; Tamanho = 12; Linha = 1.5; Depois = 0 },
      @{ Fonte = 'Georgia'; Tamanho = 10; Linha = 1.15; Depois = 6 })) {
    for ($i = 1; $i -le 9; $i++) {
      $p = Add-Para $d "$($bloco.Fonte) $($bloco.Tamanho) pt, múltiplo $($bloco.Linha) ($i). $texto" -1
      $p.Font.Name = $bloco.Fonte
      $p.Font.Size = $bloco.Tamanho
      $p.ParagraphFormat.SpaceBefore = 0
      $p.ParagraphFormat.SpaceAfter = $bloco.Depois
      $p.ParagraphFormat.LineSpacingRule = 5   # wdLineSpaceMultiple
      $p.ParagraphFormat.LineSpacing = [single](12 * $bloco.Linha)
      $p.ParagraphFormat.WidowControl = $false   # isola a altura de linha (o Textos ainda não faz controle de viúvas)
    }
    $null = Add-Para $d '' -1
  }
  Save-Doc $d '04-espacamento-multiplo.docx'
}
catch {
  Write-Output "ERRO: $($_.Exception.Message)"
  Write-Output $_.InvocationInfo.PositionMessage
  exit 1
}
finally {
  $word.Quit([ref]$wdDoNotSaveChanges)
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
}
