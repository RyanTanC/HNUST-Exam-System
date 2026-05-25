param()

$ErrorActionPreference = 'Stop'
$base = 'https://raw.githubusercontent.com/fwwdn/sensitive-stop-words/master'

$items = @(
    ,@("$base/%E8%89%B2%E6%83%85%E7%B1%BB.txt", 'csv')
    ,@("$base/%E6%94%BF%E6%B2%BB%E7%B1%BB.txt", 'csv')
    ,@("$base/%E5%B9%BF%E5%91%8A.txt", 'line')
    ,@("$base/%E6%B6%89%E6%9E%AA%E6%B6%89%E7%88%86%E8%BF%9D%E6%B3%95%E4%BF%A1%E6%81%AF%E5%85%B3%E9%94%AE%E8%AF%8D.txt", 'line')
)

$all = New-Object System.Collections.Generic.HashSet[string]
$wc = New-Object System.Net.WebClient

foreach ($item in $items) {
    $url = $item[0]
    $fmt = $item[1]
    Write-Host "Downloading $url..."
    $bytes = $wc.DownloadData($url)
    $text = [System.Text.Encoding]::UTF8.GetString($bytes)
    if ($fmt -eq 'csv') {
        $parts = $text -split ','
    } else {
        $parts = $text -split "`n"
    }
    foreach ($word in $parts) {
        $w = $word.Trim()
        if ($w.Length -gt 0) {
            $null = $all.Add($w)
        }
    }
    Write-Host "  -> $($all.Count) unique words so far"
}

$sorted = @($all) | Sort-Object

$dt = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
$js = "// sensitive-words.js -- 敏感词库`n"
$js += "// 来源: https://github.com/fwwdn/sensitive-stop-words`n"
$js += "// 自动构建: $dt`n"
$js += "// 词汇总数: $($sorted.Length)`n"
$js += '(function(){ window.SENSITIVE_WORDS = ['

$chunks = @()
for ($i = 0; $i -lt $sorted.Length; $i++) {
    $w = $sorted[$i] -replace "'", "\'"
    $chunks += "'$w'"
}
$js += ($chunks -join ',')

$js += ']; })();'

$outPath = Join-Path $PSScriptRoot '..'
$outPath = Join-Path $outPath 'js'
$outPath = Join-Path $outPath 'sensitive-words.js'

$utf8NoBom = New-Object System.Text.UTF8Encoding $false
[System.IO.File]::WriteAllText((Resolve-Path $outPath).Path, $js, $utf8NoBom)
Write-Host "Written $($sorted.Length) words to $outPath"
Write-Host "Sample words:"
$sorted | Select-Object -First 5 | ForEach-Object { Write-Host "  $_" }
