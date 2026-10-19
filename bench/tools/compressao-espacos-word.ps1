<#
  Mede até quanto o Word comprime os espaços para caber uma linha que passaria da largura.
  Usa uma frase em uma linha só e vai reduzindo a largura disponível (recuo direito) até o
  Word quebrar em duas linhas. Imprime, para cada alinhamento, a maior redução que ainda coube.
  Uso: powershell -File bench/tools/compressao-espacos-word.ps1 [-Font Cambria] [-Size 11]
#>
param(
  [string]$Font = 'Cambria',
  [double]$Size = 11,
  [string]$Text = 'O texto justificado aperta os espaços entre as palavras para que mais uma palavra caiba na linha.'
)
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
  $doc = $word.Documents.Add()
  $doc.PageSetup.LeftMargin = [single]72
  $doc.PageSetup.RightMargin = [single]72
  $p = $doc.Paragraphs.Item(1)
  $p.Range.Text = [string]$Text
  $p = $doc.Paragraphs.Item(1)
  $p.Range.Font.Name = [string]$Font
  $p.Range.Font.Size = [single]$Size
  $p.Format.SpaceAfter = [single]0
  $usable = $doc.PageSetup.PageWidth - 144
  # Largura natural: a menor largura de coluna em que a frase cabe sem precisar comprimir nada
  # é medida pelo próprio Word com alinhamento à esquerda e recuo grande o bastante.
  $results = @()
  foreach ($align in @(@{ name = 'esquerda'; v = 0 }, @{ name = 'justificado'; v = 3 })) {
    $p.Format.Alignment = [int]$align.v
    $lastFit = $null
    for ($indent = 0.0; $indent -le 400; $indent += 0.25) {
      $p.Format.RightIndent = [single]$indent
      $lines = $p.Range.ComputeStatistics([int]1)  # wdStatisticLines
      if ($lines -gt 1) { break }
      $lastFit = $indent
    }
    $results += [pscustomobject]@{ alinhamento = $align.name; larguraMinimaQueCabe = [math]::Round($usable - $lastFit, 2) }
  }
  $spaces = ([regex]::Matches($Text, ' ')).Count
  $results | ForEach-Object { "{0,-12} menor largura em que ainda cabe numa linha: {1} pt" -f $_.alinhamento, $_.larguraMinimaQueCabe }
  "espaços na frase: $spaces"
  $doc.Close([ref]0)
} finally {
  $word.Quit([ref]0)
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
}
