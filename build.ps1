param([ValidateSet('Debug','Release')][string]$Configuration = 'Release')
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$solution = Join-Path $root 'OfficeTranslate.Word.sln'
$output = Join-Path $root "src\OfficeTranslate.WordAddIn\bin\$Configuration\net48"
$excelOutput = Join-Path $root "src\OfficeTranslate.ExcelAddIn\bin\$Configuration\net48"
$powerPointOutput = Join-Path $root "src\OfficeTranslate.PowerPointAddIn\bin\$Configuration\net48"
$artifacts = Join-Path $root 'artifacts'
$env:DOTNET_CLI_HOME = Join-Path $root 'dotnet-home'
$env:NUGET_PACKAGES = Join-Path $root 'packages'
$env:DOTNET_SKIP_FIRST_TIME_EXPERIENCE = '1'
$env:DOTNET_CLI_TELEMETRY_OPTOUT = '1'

$missing = New-Object System.Collections.Generic.List[string]
$dotnet = Get-Command dotnet.exe -ErrorAction SilentlyContinue
if (-not $dotnet) {
    $missing.Add('.NET SDK 8.x (dotnet.exe was not found)')
} else {
    $sdks = @(& $dotnet.Source --list-sdks 2>$null)
    if ($sdks.Count -eq 0) { $missing.Add('.NET SDK 8.x (only runtimes are installed)') }
}

$netFx = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Microsoft SDKs\NETFXSDK\4.8' -ErrorAction SilentlyContinue
$netFx = if ($netFx) { $netFx } else { Get-ItemProperty 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Microsoft SDKs\NETFXSDK\4.8' -ErrorAction SilentlyContinue }
if (-not $netFx) { $missing.Add('.NET Framework 4.8 Developer Pack') }

$wixCandidates = @(
    $env:WIX,
    (Join-Path ${env:ProgramFiles(x86)} 'WiX Toolset v3.14\bin'),
    (Join-Path ${env:ProgramFiles(x86)} 'WiX Toolset v3.11\bin')
) | Where-Object { $_ }
$candlePath = (Get-Command candle.exe -ErrorAction SilentlyContinue).Source
$lightPath = (Get-Command light.exe -ErrorAction SilentlyContinue).Source
foreach ($candidate in $wixCandidates) {
    if (-not $candlePath -and (Test-Path (Join-Path $candidate 'candle.exe'))) { $candlePath = Join-Path $candidate 'candle.exe' }
    if (-not $lightPath -and (Test-Path (Join-Path $candidate 'light.exe'))) { $lightPath = Join-Path $candidate 'light.exe' }
}
if (-not $candlePath -or -not $lightPath) { $missing.Add('WiX Toolset 3.14 (candle.exe and light.exe were not found)') }

if ($missing.Count -gt 0) {
    $message = "Build prerequisites are missing:`n - " + ($missing -join "`n - ") + "`n`nInstall them, open a new PowerShell window, and run this script again."
    throw $message
}

dotnet restore $solution
if ($LASTEXITCODE -ne 0) { throw "dotnet restore failed with exit code $LASTEXITCODE" }
dotnet build $solution -c $Configuration --no-restore
if ($LASTEXITCODE -ne 0) { throw "dotnet build failed with exit code $LASTEXITCODE" }
New-Item -ItemType Directory -Force -Path $artifacts | Out-Null

& $candlePath (Join-Path $root 'installer\Product.wxs') ("-dBuildOutput={0}" -f $output) ("-dExcelOutput={0}" -f $excelOutput) ("-dPowerPointOutput={0}" -f $powerPointOutput) -out (Join-Path $artifacts 'Product.wixobj') -arch x64
if ($LASTEXITCODE -ne 0) { throw "candle.exe failed with exit code $LASTEXITCODE" }
& $lightPath (Join-Path $artifacts 'Product.wixobj') -out (Join-Path $artifacts 'OfficeTranslate.Office.x64.msi')
if ($LASTEXITCODE -ne 0) { throw "light.exe failed with exit code $LASTEXITCODE" }
Write-Host "Created $artifacts\OfficeTranslate.Office.x64.msi"
