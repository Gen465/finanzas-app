# Finanzas - Auto-import from 1Money CSV
# Drop a new CSV in the OneDrive folder, then double-click import.bat
# Safe to run multiple times - duplicates are skipped automatically.

$SB_URL = 'https://amkactcllziblbxocoik.supabase.co'
$SB_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImFta2FjdGNsbHppYmxieG9jb2lrIiwicm9sZSI6ImFub24iLCJpYXQiOjE3Nzk0MjYyMzQsImV4cCI6MjA5NTAwMjIzNH0.5xInHiBz-fB2PT3uvC6fA7GLcHGxjDTDpvp1_KtMI0k'
$FOLDER = 'C:\Users\gen46\OneDrive\Finanzas\OneMoneyApp'

$MONTHS = @('Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic')

$hdrs = @{
  'apikey'        = $SB_KEY
  'Authorization' = "Bearer $SB_KEY"
}

# Find latest CSV
$csvFile = Get-ChildItem $FOLDER -Filter '*.csv' | Sort-Object LastWriteTime -Descending | Select-Object -First 1
if (-not $csvFile) { Write-Host "No CSV found in $FOLDER"; Read-Host "Press Enter"; exit }
Write-Host ""
Write-Host "  File : $($csvFile.Name)"

# Parse a CSV line handling quoted fields
function Parse-Line($line) {
  $fields = @(); $field = ''; $inQ = $false
  foreach ($ch in $line.ToCharArray()) {
    if     ($ch -eq '"')               { $inQ = -not $inQ }
    elseif ($ch -eq ',' -and -not $inQ){ $fields += $field; $field = '' }
    else                               { $field += $ch }
  }
  return ($fields + $field)
}

$lines   = (Get-Content $csvFile.FullName -Raw -Encoding UTF8) -split '\r?\n'
$section = 'none'
$txList  = [System.Collections.Generic.List[hashtable]]::new()
$balList = [System.Collections.Generic.List[hashtable]]::new()

foreach ($line in $lines) {
  $line = $line.Trim()
  if (-not $line) { continue }
  $f = Parse-Line $line

  if ($f[0] -eq 'DATE'  -and $f[1] -eq 'TYPE')    { $section = 'tx';  continue }
  if ($f[0] -eq 'NAME'  -and $f[1] -eq 'BALANCE')  { $section = 'bal'; continue }

  if ($section -eq 'tx') {
    if ($f[0] -notmatch '^\d{2}/\d{2}/\d{4}$') { continue }
    $parts = $f[0] -split '/'; $d = $parts[0]; $m = $parts[1]; $y = $parts[2]
    $isoDate   = "$y-$($m.PadLeft(2,'0'))-$($d.PadLeft(2,'0'))"
    $monthYear = "$($MONTHS[[int]$m - 1])-$($y.Substring(2))"

    $cat = $f[3].Trim(); $sub = $null
    if ($cat -match '^(.+?)\s*\(\s*(.+?)\s*\)\s*$') {
      $cat = $matches[1].Trim(); $sub = $matches[2].Trim()
    }

    $txList.Add(@{
      date         = $isoDate
      month_year   = $monthYear
      type         = $f[1].Trim()
      from_account = $f[2].Trim()
      category     = $cat
      sub_category = $sub
      amount       = [double]$f[4]
      currency     = $f[5].Trim()
      notes        = if ($f[9].Trim()) { $f[9].Trim() } else { $null }
      tags         = if ($f[8].Trim()) { $f[8].Trim() } else { $null }
    })
  }

  if ($section -eq 'bal') {
    if (-not $f[0].Trim()) { continue }
    $balList.Add(@{
      name     = $f[0].Trim()
      balance  = [double]$f[1]
      currency = if ($f.Count -gt 2 -and $f[2].Trim()) { $f[2].Trim() } else { 'EUR' }
    })
  }
}

Write-Host "  Total: $($txList.Count) transactions  |  $($balList.Count) accounts in file"

# Deduplicate
$dates   = $txList | ForEach-Object { $_.date } | Sort-Object
$minDate = $dates[0]; $maxDate = $dates[-1]
$existing = Invoke-RestMethod "$SB_URL/rest/v1/transactions?date=gte.$minDate&date=lte.$maxDate&select=date,type,from_account,category,amount" `
  -Headers $hdrs -ContentType 'application/json'
$exKeys   = @{}
$existing | ForEach-Object {
  $exKeys["$($_.date)|$($_.type)|$($_.from_account)|$($_.category)|$($_.amount)"] = 1
}

$newTx = $txList | Where-Object {
  -not $exKeys.ContainsKey("$($_.date)|$($_.type)|$($_.from_account)|$($_.category)|$($_.amount)")
}
$newCount = @($newTx).Count
$skipped  = $txList.Count - $newCount
Write-Host "  New  : $newCount to import  |  Skipped: $skipped duplicates"

# Import new transactions in batches of 100
# Helper: send UTF-8 encoded JSON body (fixes PS 5.1 encoding bug with accented chars)
function Invoke-SB($method, $url, $body) {
  $bytes = [System.Text.Encoding]::UTF8.GetBytes($body)
  Invoke-RestMethod $url -Method $method -Headers $hdrs -ContentType 'application/json' -Body $bytes | Out-Null
}

if ($newCount -gt 0) {
  $arr = @($newTx)
  for ($i = 0; $i -lt $arr.Count; $i += 100) {
    $batch    = @($arr[$i..([Math]::Min($i + 99, $arr.Count - 1))])
    $jsonBody = '[' + (($batch | ForEach-Object { ConvertTo-Json $_ -Depth 5 -Compress }) -join ',') + ']'
    Invoke-SB POST "$SB_URL/rest/v1/transactions" $jsonBody
  }
  Write-Host "  OK   : Transactions imported"
}

# Update account balances
$today = Get-Date -Format 'yyyy-MM-dd'
foreach ($b in $balList) {
  $name = [Uri]::EscapeDataString($b.name)
  Invoke-SB PATCH "$SB_URL/rest/v1/accounts?name=eq.$name" (@{ balance = $b.balance; last_updated = $today } | ConvertTo-Json)
}
Write-Host "  OK   : Balances updated ($($balList.Count) accounts)"

# Log import
Invoke-SB POST "$SB_URL/rest/v1/import_log" (@{ filename = $csvFile.Name; row_count = $newCount } | ConvertTo-Json)

Write-Host ""
Write-Host "  Done! Open the app to see your data."
Write-Host ""
Read-Host "  Press Enter to close"
