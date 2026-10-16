param(
    [ValidateSet('Debug','Release')][string]$Configuration = 'Release',
    [string]$EvidenceDirectory = (Join-Path $PSScriptRoot '..\TestResults\parallel-word-ranges')
)
# Run in Windows PowerShell -STA on the interactive desktop after building.
# Uses synthetic, unsaved documents only; never registers or installs an add-in.
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
New-Item -ItemType Directory -Force -Path $EvidenceDirectory | Out-Null
$resultFile = Join-Path $EvidenceDirectory 'range-result.json'
$result = [ordered]@{ Interactive = [Environment]::UserInteractive; Session = (Get-Process -Id $PID).SessionId; Cases = @() }
$word = $null
$doc = $null
function Assert-That([bool]$condition, [string]$message) { if (-not $condition) { throw $message } }
try {
    Assert-That ([Threading.Thread]::CurrentThread.ApartmentState -eq 'STA') 'Requires Windows PowerShell -STA'
    Assert-That $result.Interactive 'Requires an interactive desktop'
    $assembly = [Reflection.Assembly]::LoadFrom((Join-Path $root "src\OfficeTranslate.WordAddIn\bin\$Configuration\net48\OfficeTranslate.WordAddIn.dll"))
    $result.Version = $assembly.GetName().Version.ToString()
    $result.Mvid = $assembly.ManifestModule.ModuleVersionId.ToString()
    Assert-That ($result.Version -eq (Get-Content (Join-Path $root 'VERSION.txt') -Raw).Trim()) 'Candidate version mismatch'
    $serviceType = $assembly.GetType('OfficeTranslate.WordAddIn.WordTranslationService', $true)
    $flags = [Reflection.BindingFlags]'NonPublic,Instance'
    $staticFlags = [Reflection.BindingFlags]'NonPublic,Static'
    $read = $serviceType.GetMethod('ReadDocumentParagraphs', $flags)
    $write = $serviceType.GetMethod('WriteTarget', $staticFlags)
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $service = $serviceType.GetConstructors()[0].Invoke([object[]]@($word.PSObject.BaseObject))
    $sourceParts = @('Alpha source paragraph', 'Beta source paragraph', 'Gamma source paragraph')
    $translated = @('Translated first paragraph with a much longer replacement.', 'Second translation.', 'Third.')
    foreach ($kind in @('body', 'table')) {
        foreach ($bilingual in @($false, $true)) {
            $doc = $word.Documents.Add()
            $table = $null
            if ($kind -eq 'body') {
                # Word supplies the final paragraph mark itself.
                $doc.Content.Text = ($sourceParts -join "`r")
            } else {
                $table = $doc.Tables.Add($doc.Range(0, 0), 3, 1)
                for ($i = 0; $i -lt 3; $i++) { $table.Cell($i + 1, 1).Range.Text = $sourceParts[$i] }
            }
            $serviceType.GetField('_taskDocument', $flags).SetValue($service, $doc.PSObject.BaseObject)
            $targets = $read.Invoke($service, @())
            Assert-That ($targets.Count -eq 3) "Unexpected target count kind=$kind"
            for ($i = 0; $i -lt 3; $i++) {
                $write.Invoke($null, [object[]]@($targets[$i].PSObject.BaseObject, $translated[$i], $bilingual)) | Out-Null
                if ($kind -eq 'body') {
                    $expected = ''
                    for ($j = 0; $j -lt 3; $j++) {
                        if ($bilingual) {
                            $expected += $sourceParts[$j] + "`r"
                            if ($j -le $i) { $expected += $translated[$j] + "`r" }
                        } else {
                            $expected += $(if ($j -le $i) { $translated[$j] } else { $sourceParts[$j] }) + "`r"
                        }
                    }
                    $actual = $doc.Content.Text
                    if ($actual -cne $expected) {
                        $firstDifference = 0
                        while ($firstDifference -lt [Math]::Min($actual.Length, $expected.Length) -and $actual[$firstDifference] -ceq $expected[$firstDifference]) { $firstDifference++ }
                        $result.Mismatch = [ordered]@{ ActualLength = $actual.Length; ExpectedLength = $expected.Length; FirstDifference = $firstDifference }
                    }
                    Assert-That ($actual -ceq $expected) "Range mismatch kind=$kind bilingual=$bilingual block=$($i+1)"
                } else {
                    for ($j = 0; $j -lt 3; $j++) {
                        $expected = $sourceParts[$j]
                        if ($j -le $i) { $expected = if ($bilingual) { $expected + "`r" + $translated[$j] } else { $translated[$j] } }
                        $actual = $table.Cell($j + 1, 1).Range.Text.TrimEnd([char[]]"`r`a")
                        Assert-That ($actual -ceq $expected) "Range mismatch kind=$kind bilingual=$bilingual block=$($i+1) cell=$($j+1)"
                    }
                }
            }
            $result.Cases += [ordered]@{ Kind = $kind; Bilingual = $bilingual; Targets = 3; VerifiedCommits = 3; Passed = $true }
            $doc.Close(0)
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($doc)
            $doc = $null
        }
    }
    $result.Passed = $true
} catch {
    $result.Passed = $false
    $result.Error = $_.Exception.Message
} finally {
    if ($null -ne $doc) { try { $doc.Close(0) } catch {} }
    if ($null -ne $word) {
        try { $word.Quit(0) } catch {}
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($word)
    }
    $result | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $resultFile -Encoding UTF8
}
if (-not $result.Passed) { throw "Word range probe failed; see $resultFile" }
