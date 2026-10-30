$ErrorActionPreference = 'Stop'
$root = (Resolve-Path '.').Path
$qa = Join-Path $root '_report_qa\word-render'
New-Item -ItemType Directory -Force -Path $qa | Out-Null
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
    $sources = Get-ChildItem $root -Filter '*.docx' | Sort-Object Name
    Write-Output ("Found source documents: {0}" -f $sources.Count)
    for ($n = 1; $n -le $sources.Count; $n++) {
        $source = $sources[$n - 1].FullName
        $target = Join-Path $qa ("report-{0}.pdf" -f $n)
        Write-Output ("Opening document {0}" -f $n)
        try {
            $doc = $word.Documents.Open($source, $false, $true)
            Write-Output ("Opened document {0}" -f $n)
            $doc.ExportAsFixedFormat($target, 17)
            Write-Output ("Exported document {0}" -f $n)
            $doc.Close(0)
        }
        catch {
            Write-Output ("Failed document {0}: {1}" -f $n, $_.Exception.Message)
            throw
        }
        if ($n -eq 1) { break }
    }
}
finally { $word.Quit() }
Get-ChildItem $qa -Filter '*.pdf' | Select-Object Name,Length
