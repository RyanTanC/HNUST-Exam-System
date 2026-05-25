param(
    [int]$Count = 10,
    [string]$Name = "AutoTester",
    [string]$Text = "load test"
)

$ErrorActionPreference = 'Stop'
$key = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImVkdG56c3hsb2dkeHlnY2J0d21iIiwicm9sZSI6ImFub24iLCJpYXQiOjE3Nzk2MTE2NzAsImV4cCI6MjA5NTE4NzY3MH0.GYU_aMzvYJM20yV1haMG7UZ0wt8s1URtme7WrvRwuWk"
$base = "https://edtnzsxlogdxygcbtwmb.supabase.co"
$hdr = @{ "apikey" = $key; "Content-Type" = "application/json" }

Write-Host "=== Rate Limit & Filter Test ==="

# 1. Sensitive word filter test
Write-Host "`n[1/4] Filter test"
$js = Get-Content "C:\Users\Hyper_hui\HNUST-Exam-System\js\sensitive-words.js" -Raw -Encoding UTF8
$words = [regex]::Matches($js, "'([^']+)'") | ForEach-Object { $_.Groups[1].Value }
$tests = @(
    @("normal comment", $false), @("fuck you", $true),
    @("sell gun", $true), @("nice weather", $false)
)
$p = 0; $f = 0
foreach ($t in $tests) {
    $txt = $t[0]; $exp = $t[1]; $found = $null
    foreach ($w in $words) { if ($txt.ToLower().Contains($w.ToLower())) { $found = $w; break } }
    $blocked = ($null -ne $found)
    if ($blocked -eq $exp) { $p++; Write-Host "  PASS: $txt" -ForegroundColor Green }
    else { $f++; Write-Host "  FAIL: $txt" -ForegroundColor Red }
}
Write-Host "  $p pass, $f fail"

# 2. Send comments rapidly (test rate limit on Supabase side)
Write-Host "`n[2/4] Sending $Count comments as '$Name'"
try { $ip = (Invoke-RestMethod "https://api.ipify.org?format=json" -UseBasicParsing).ip } catch { $ip = "0.0.0.0" }
$ok = 0; $blocked = 0; $err = 0
$start = Get-Date -Format o

for ($i = 1; $i -le $Count; $i++) {
    $body = "{`"nickname`":`"$Name`",`"ip`":`"$ip`",`"content`":`"$Text #$i`"}"
    try {
        Invoke-WebRequest "$base/rest/v1/comments" -Headers $hdr -Method Post -Body $body -UseBasicParsing -ErrorAction Stop | Out-Null
        $ok++
        $now = Get-Date
        $elapsed = [int]($now - [DateTime]$start).TotalSeconds
        Write-Host "  [$i/$Count] +${elapsed}s -> OK" -ForegroundColor Green
    } catch {
        $code = [int]$_.Exception.Response.StatusCode
        if ($code -eq 429) { $blocked++; Write-Host "  [$i/$Count] -> BLOCKED (429)" -ForegroundColor Red }
        else { $err++; Write-Host "  [$i/$Count] -> HTTP $code" -ForegroundColor DarkRed }
    }
    Start-Sleep -Milliseconds 100
}

$now2 = Get-Date
$dur = [int]($now2 - [DateTime]$start).TotalSeconds
Write-Host "  Sent: $ok | Blocked: $blocked | Errors: $err (${dur}s)" -ForegroundColor $(if ($ok -le 3) { "Green" } else { "Red" })
Write-Host "  => Rate limit: $(if ($blocked -gt 0 -or $ok -le 3) {'WORKING'} else {'CHECK NEEDED'})" -ForegroundColor $(if ($blocked -gt 0 -or $ok -le 3) { "Green" } else { "Yellow" })

# 3. Verify counts via query
Write-Host "`n[3/4] Verify DB counts"
$min = (Get-Date).ToUniversalTime().AddMinutes(-1).ToString("yyyy-MM-ddTHH:mm:ssZ")
$hr = (Get-Date).ToUniversalTime().AddHours(-1).ToString("yyyy-MM-ddTHH:mm:ssZ")
$en = [System.Uri]::EscapeDataString($Name)
try {
    $minC = (Invoke-RestMethod "$base/rest/v1/comments?created_at=gte.$min&select=id&nickname=eq.$en" -Headers @{ "apikey" = $key }).Count
    $hrC = (Invoke-RestMethod "$base/rest/v1/comments?created_at=gte.$hr&select=id&nickname=eq.$en" -Headers @{ "apikey" = $key }).Count
    Write-Host "  1min: $minC | 1hr: $hrC"
    if ($minC -ge 3) { Write-Host "  1min limit ACTIVE" -ForegroundColor Red }
    if ($hrC -ge 12) { Write-Host "  1hr limit ACTIVE" -ForegroundColor Red }
} catch { Write-Host "  Query failed: $_" -ForegroundColor Red }

# 4. Cleanup
Write-Host "`n[4/4] Cleanup"
try {
    $list = Invoke-RestMethod "$base/rest/v1/comments?select=id&nickname=eq.$en" -Headers @{ "apikey" = $key }
    Write-Host "  Deleting $($list.Count) test entries..."
    foreach ($item in $list) {
        try { Invoke-RestMethod "$base/rest/v1/comments?id=eq.$($item.id)" -Headers @{ "apikey" = $key } -Method Delete -UseBasicParsing -ErrorAction Stop | Out-Null } catch {}
    }
} catch { Write-Host "  Cleanup failed: $_" }
Write-Host "  Done"

Write-Host "`n=== Complete ==="