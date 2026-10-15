<#
  Usa o Microsoft Word instalado como REFERÊNCIA para a bancada de fidelidade.

  Entrada: arquivo JSON com uma lista de tarefas:
    [{ "op": "pdf",   "in": "C:\...\a.docx", "out": "C:\...\a.pdf" },
     { "op": "check", "in": "C:\...\b.docx" }]
  Saída: o mesmo JSON, com "ok", "pages", "error" e "ms" preenchidos, gravado em -Results.

  Uma única instância do Word (invisível) processa todas as tarefas.
  Macros ficam desativadas e nenhum diálogo é exibido.
#>
param(
  [Parameter(Mandatory = $true)][string]$Jobs,
  [Parameter(Mandatory = $true)][string]$Results
)

$ErrorActionPreference = 'Stop'
# Atenção: no PowerShell $Jobs e $jobs são a MESMA variável, por isso o nome $taskList.
# O ForEach-Object desembrulha o array que o ConvertFrom-Json do PS 5.1 devolve como um item só.
$taskList = @(Get-Content -Raw -Encoding UTF8 $Jobs | ConvertFrom-Json | ForEach-Object { $_ })

$wdDoNotSaveChanges = 0
$wdExportFormatPDF = 17
$wdStatisticPages = 2
$msoAutomationSecurityForceDisable = 3
# Senha falsa: documento protegido falha na hora em vez de abrir um diálogo e travar.
$dummyPassword = '__textos_bench__'

$word = New-Object -ComObject Word.Application
$out = @()
try {
  $word.Visible = $false
  $word.DisplayAlerts = 0
  $word.AutomationSecurity = $msoAutomationSecurityForceDisable
  $word.Options.SaveNormalPrompt = $false

  foreach ($job in $taskList) {
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $result = [ordered]@{ op = $job.op; in = $job.in; out = $job.out; ok = $false; pages = $null; error = $null; ms = 0 }
    $doc = $null
    try {
      # Open(FileName, ConfirmConversions, ReadOnly, AddToRecentFiles, PasswordDocument,
      #      PasswordTemplate, Revert, WritePasswordDocument, WritePasswordTemplate, Format,
      #      Encoding, Visible, OpenAndRepair, DocumentDirection, NoEncodingDialog)
      # OpenAndRepair = false: se o Word precisar "consertar" o arquivo, isso é uma falha.
      # Sempre [string]/[int] puros: o COM do Word trava se receber um PSObject.
      [string]$inPath = $job.in
      $doc = $word.Documents.Open($inPath, $false, $true, $false, $dummyPassword, '', $true, '', '', 0, [Type]::Missing, $false, $false, 0, $true)
      $result.pages = [int]$doc.ComputeStatistics([int]$wdStatisticPages)
      if ($job.op -eq 'pdf') {
        [string]$outPath = $job.out
        $doc.ExportAsFixedFormat($outPath, [int]$wdExportFormatPDF)
      }
      $result.ok = $true
    } catch {
      $result.error = $_.Exception.Message.Trim()
    } finally {
      if ($doc) { try { $doc.Close([ref]$wdDoNotSaveChanges) } catch {} }
      $result.ms = $sw.ElapsedMilliseconds
    }
    $out += [pscustomobject]$result
  }
} finally {
  try { $word.Quit([ref]$wdDoNotSaveChanges) } catch {}
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}

ConvertTo-Json -InputObject @($out) -Depth 4 | Out-File -Encoding utf8 $Results
