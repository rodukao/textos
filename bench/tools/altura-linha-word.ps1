<#
  Mede no Word a altura de linha (distância entre linhas, em pt) para combinações de fonte,
  tamanho e espaçamento múltiplo. Serve para conferir a regra de altura de linha do motor.
  Uso: powershell -File bench/tools/altura-linha-word.ps1 [-Fontes "Calibri,Cambria"] [-Tamanhos "11,12"] [-Linhas "240,259,360"]
  Linhas em 240 avos (240 = simples, 259 = 1,08, 360 = 1,5).
#>
param(
  [string]$Fontes = 'Calibri,Cambria,Times New Roman,Arial',
  [string]$Tamanhos = '10,11,12,14',
  [string]$Linhas = '240,259,276,360'
)
# Listas separadas por vírgula: com -File o PowerShell não converte '240,259' em array.
$listaFontes = $Fontes.Split(',') | ForEach-Object { $_.Trim() }
$listaTamanhos = $Tamanhos.Split(',') | ForEach-Object { [double]::Parse($_.Trim(), [Globalization.CultureInfo]::InvariantCulture) }
$listaLinhas = $Linhas.Split(',') | ForEach-Object { [int]$_.Trim() }
$word = New-Object -ComObject Word.Application
$word.Visible = $false
try {
  foreach ($fonte in $listaFontes) { foreach ($tamanho in $listaTamanhos) { foreach ($linha in $listaLinhas) {
    $d = $word.Documents.Add()
    $d.Content.Text = (('Xg' + "`r") * 20)
    $r = $d.Content
    $r.Font.Name = $fonte
    $r.Font.Size = $tamanho
    $pf = $r.ParagraphFormat
    $pf.SpaceBefore = 0
    $pf.SpaceAfter = 0
    if ($linha -eq 240) { $pf.LineSpacingRule = 0 } else { $pf.LineSpacingRule = 5; $pf.LineSpacing = [single](12.0 * $linha / 240.0) }
    # 6 = wdVerticalPositionRelativeToPage
    $altura = ($d.Paragraphs(11).Range.Information(6) - $d.Paragraphs(1).Range.Information(6)) / 10
    '{0,-18} {1,5} pt  linha {2,3}/240  ->  {3:0.000} pt' -f $fonte, $tamanho, $linha, $altura
    $d.Close(0)
  } } }
}
finally { $word.Quit() }
