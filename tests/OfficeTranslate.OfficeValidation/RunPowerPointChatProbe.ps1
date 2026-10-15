param([Parameter(Mandatory=$true)][string]$EvidenceDirectory,
    [ValidateSet('HKCU','HKLM')][string]$RegistrationHive='HKCU')
$ErrorActionPreference='Stop'
function Invoke-RegistryCommand([string[]]$Arguments) {
    $previous=$ErrorActionPreference
    try {
        # reg.exe prints successful import status on stderr in Windows PowerShell.
        $ErrorActionPreference='Continue'
        & reg.exe @Arguments 2>&1 | Out-Null
        if($LASTEXITCODE -ne 0){throw "reg.exe failed: $LASTEXITCODE"}
    } finally { $ErrorActionPreference=$previous }
}
if(Get-Process POWERPNT -ErrorAction SilentlyContinue){throw 'Close PowerPoint before isolated validation'}
$evidence=(Resolve-Path -LiteralPath $EvidenceDirectory).Path
$output=Join-Path $evidence 'real-powerpoint-chat.txt'
if(Test-Path -LiteralPath $output){throw 'Preserve existing probe evidence'}
$bin=Join-Path $PSScriptRoot 'bin\Release\net48'
$dll=Join-Path $bin 'OfficeTranslate.OfficeValidation.dll'
$product=Join-Path $PSScriptRoot '..\..\src\OfficeTranslate.PowerPointAddIn\bin\Release\net48\OfficeTranslate.PowerPointAddIn.dll'
if((Get-FileHash $product).Hash -ne (Get-FileHash (Join-Path $bin 'OfficeTranslate.PowerPointAddIn.dll')).Hash){throw 'Probe DLL differs from production candidate'}
$helperGuid='{1CFF20FA-E1F4-4CE9-88AC-092303010001}'
$paneGuid='{EFD0B5E1-8082-4C8C-9429-3011D0230003}'
$helperProgId='OfficeTranslate.PowerPointChatProbe'
$paneProgId='OfficeTranslate.PowerPointChatPane'
$paths=@("Software\Classes\CLSID\$helperGuid", "Software\Classes\$helperProgId",
    "Software\Classes\CLSID\$paneGuid", "Software\Classes\$paneProgId",
    "Software\Microsoft\Office\PowerPoint\Addins\$helperProgId",
    'Software\Microsoft\Office\PowerPoint\Addins\OfficeTranslate.PowerPointAddIn')
$hive=if($RegistrationHive -eq 'HKLM'){[Microsoft.Win32.Registry]::LocalMachine}else{[Microsoft.Win32.Registry]::CurrentUser}
$backups=@{}
$app=$null
Start-Transcript -Path (Join-Path $evidence 'real-powerpoint-run.log') | Out-Null
try {
    foreach($path in $paths) {
        $key=$hive.OpenSubKey($path)
        if($key) {
            $key.Dispose()
            $backup=Join-Path $evidence ("registry-before-"+$backups.Count+'.reg')
            if(Test-Path -LiteralPath $backup){throw 'Preserve existing registry backup'}
            Invoke-RegistryCommand @('export',($RegistrationHive+'\'+$path),$backup,'/y')
            $backups[$path]=$backup
        } else { $backups[$path]=$null }
    }
    foreach($entry in @(
        @{Guid=$helperGuid; ProgId=$helperProgId; Class='OfficeTranslate.OfficeValidation.PowerPointChatProbeAddIn'; Dll=$dll; Assembly='OfficeTranslate.OfficeValidation, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null'},
        @{Guid=$paneGuid; ProgId=$paneProgId; Class='OfficeTranslate.WordAddIn.DocumentChatPane'; Dll=(Join-Path $bin 'OfficeTranslate.PowerPointAddIn.dll'); Assembly='OfficeTranslate.PowerPointAddIn, Version=2.3.1.0, Culture=neutral, PublicKeyToken=null'}
    )) {
        $key=$hive.CreateSubKey("Software\Classes\$($entry.ProgId)\CLSID")
        $key.SetValue('',$entry.Guid);$key.Dispose()
        $key=$hive.CreateSubKey("Software\Classes\CLSID\$($entry.Guid)\ProgId")
        $key.SetValue('',$entry.ProgId);$key.Dispose()
        $key=$hive.CreateSubKey("Software\Classes\CLSID\$($entry.Guid)\InprocServer32")
        try {
            $key.SetValue('','mscoree.dll');$key.SetValue('ThreadingModel','Both')
            $key.SetValue('Class',$entry.Class);$key.SetValue('Assembly',$entry.Assembly)
            $key.SetValue('RuntimeVersion','v4.0.30319');$key.SetValue('CodeBase',$entry.Dll)
        } finally { $key.Dispose() }
    }
    foreach($suffix in 'Control','Implemented Categories\{40FC6ED5-2438-11CF-A3DB-080036F12502}') {
        $key=$hive.CreateSubKey("Software\Classes\CLSID\$paneGuid\$suffix");$key.Dispose()
    }
    $key=$hive.CreateSubKey("Software\Microsoft\Office\PowerPoint\Addins\$helperProgId")
    $key.SetValue('FriendlyName','OfficeTranslate isolated Chat probe');$key.SetValue('LoadBehavior',3,[Microsoft.Win32.RegistryValueKind]::DWord);$key.Dispose()
    # Prevent an already-installed assembly with the same product identity from
    # loading before the candidate. Restore the entire existing key below.
    $key=$hive.CreateSubKey('Software\Microsoft\Office\PowerPoint\Addins\OfficeTranslate.PowerPointAddIn')
    $key.SetValue('LoadBehavior',0,[Microsoft.Win32.RegistryValueKind]::DWord);$key.Dispose()
    $activation=[Activator]::CreateInstance([Type]::GetTypeFromProgID($helperProgId,$true))
    Write-Output ("Direct test-helper COM activation: "+$activation.GetType().FullName)
    $app=New-Object -ComObject PowerPoint.Application
    $app.Visible=-1
    $app.COMAddIns.Update()
    $addin=$app.COMAddIns.Item($helperProgId)
    $addin.Connect=$true
    if(-not $addin.Connect -or $null -eq $addin.Object){throw 'Probe add-in failed to connect'}
    $addin.Object.Run($output)
    $result=Get-Content -LiteralPath $output -Raw
    Write-Output $result
    if($result -notlike 'PASS:*'){throw 'Real PowerPoint probe failed'}
} finally {
    try { if($app) {
        for($index=$app.Presentations.Count;$index -ge 1;$index--) {
            $doc=$app.Presentations.Item($index);$doc.Saved=-1;$doc.Close()
        }
        $app.Quit()
        [Runtime.InteropServices.Marshal]::FinalReleaseComObject($app) | Out-Null
    } } catch { Write-Warning 'Test Office cleanup failed; restoring registry before reporting the error.' }
    foreach($path in $paths) {
        # Only these explicitly enumerated test/installed-addin registration subtrees.
        if(-not $backups.ContainsKey($path)){continue}
        $hive.DeleteSubKeyTree($path,$false)
        if($backups[$path]) {
            Invoke-RegistryCommand @('import',$backups[$path])
            $after=$backups[$path]+'.restored.reg'
            Invoke-RegistryCommand @('export',($RegistrationHive+'\'+$path),$after,'/y')
            if((Get-FileHash $after).Hash -ne (Get-FileHash $backups[$path]).Hash){throw 'Registry restoration mismatch'}
        } else {
            $key=$hive.OpenSubKey($path)
            if($key){$key.Dispose();throw 'Temporary registration remains'}
        }
    }
    Write-Output "$RegistrationHive registrations restored and verified; no MSI/settings changes."
    Stop-Transcript | Out-Null
}
