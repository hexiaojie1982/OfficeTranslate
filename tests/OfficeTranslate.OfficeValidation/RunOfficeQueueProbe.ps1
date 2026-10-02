param([Parameter(Mandatory=$true)][string]$EvidenceDirectory)
$ErrorActionPreference='Stop'
if(Get-Process WINWORD -ErrorAction SilentlyContinue){throw 'Close Word before isolated validation'}
Start-Transcript -Path (Join-Path $EvidenceDirectory 'word-queue-run.log') -Append | Out-Null
$dll=Join-Path $PSScriptRoot 'bin\Release\net48\OfficeTranslate.OfficeValidation.dll'
$core=Join-Path $PSScriptRoot 'bin\Release\net48\OfficeTranslate.Core.dll'
$productCore=Join-Path $PSScriptRoot '..\..\src\OfficeTranslate.Core\bin\Release\net48\OfficeTranslate.Core.dll'
if((Get-FileHash $core).Hash -ne (Get-FileHash $productCore).Hash){throw 'Core does not match production candidate'}
# Existing Office entry, temporary HKCU COM override only; NO HKLM changes.
$path='Software\Classes\CLSID\{7898D80A-8CB7-4E18-9D52-3DC5901786D7}'
$hive=[Microsoft.Win32.Registry]::CurrentUser
$backup=Join-Path $EvidenceDirectory 'queue-word-hkcu-before.reg'
$key=$hive.OpenSubKey($path)
$existed=$null -ne $key
if($key){$key.Dispose()}
if(Test-Path $backup){throw 'Preserve existing registration backup'}
if($existed){
 & reg.exe export ('HKCU\'+$path) $backup /y | Out-Null
 if($LASTEXITCODE -ne 0){throw 'Registration backup failed'}
}
$word=$null
try{
 $key=$hive.CreateSubKey($path+'\InprocServer32')
 try{
  $key.SetValue('','mscoree.dll');$key.SetValue('ThreadingModel','Both')
  $key.SetValue('Class','OfficeTranslate.OfficeValidation.QueueProbeAddIn')
  $key.SetValue('Assembly','OfficeTranslate.OfficeValidation, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null')
  $key.SetValue('RuntimeVersion','v4.0.30319');$key.SetValue('CodeBase',$dll)
 }finally{$key.Dispose()}
 # Office automation startup may suppress COM add-ins. Launch Word normally
 # via the desktop, then attach to that isolated (previously absent) instance.
 Write-Output 'WAITING_FOR_NORMAL_WORD_UI_STARTUP'
 $deadline=[DateTime]::UtcNow.AddSeconds(60)
 while($null -eq $word -and [DateTime]::UtcNow -lt $deadline){
  try{$word=[Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application')}catch{Start-Sleep -Milliseconds 200}
 }
 if($null -eq $word){throw 'Word UI instance did not start'}
 if($word.Documents.Count -eq 0){$word.Documents.Add() | Out-Null}
 $word.COMAddIns.Update()
 $addin=$word.COMAddIns.Item('OfficeTranslate.WordAddIn')
 if(-not $addin.Connect){throw 'Validation add-in did not load at normal UI startup'}
 $automation=$addin.Object
 if($null -eq $automation){throw 'Validation automation object missing'}
 foreach($mode in @('expire','supersede','normal')){
  $output=Join-Path $EvidenceDirectory ("word-real-queue-$mode.txt")
  $automation.RunQueueProbe($mode,$output)
  $deadline=[DateTime]::UtcNow.AddSeconds(15)
  while(-not(Test-Path $output) -and [DateTime]::UtcNow -lt $deadline){Start-Sleep -Milliseconds 100}
  if(-not(Test-Path $output)){throw "Probe did not finish: $mode"}
  $result=Get-Content $output -Raw
  Write-Output $result
  if($result -notmatch 'is_ui=True'){throw 'Not UI thread'}
  if($mode -eq 'normal'){
   if($result -notmatch 'calls=1 reason=none start=0 end=5'){throw 'Normal restore failed'}
  }else{
   if($result -notmatch 'calls=0' -or $result -notmatch 'start=7 end=12'){throw 'Old restore changed newer selection'}
   if($mode -eq 'expire' -and $result -notmatch 'reason=expired'){throw 'Timeout not exercised'}
   if($mode -eq 'supersede' -and $result -notmatch 'reason=superseded'){throw 'Generation not exercised'}
  }
  $dontSave=0;$word.ActiveDocument.Close([ref]$dontSave)
 }
 Write-Output 'PASS_REAL_WORD_QUEUE: actual Office native UI thread; expiry/generation preserve newer real selection; normal restore works.'
}catch{Write-Output $_.ScriptStackTrace;throw
}finally{
 if($word){
  try{$dontSave=0;$word.Quit([ref]$dontSave)}catch{Write-Warning 'Validation Word exit needs recovery'}
  [Runtime.InteropServices.Marshal]::FinalReleaseComObject($word) | Out-Null
 }
 [GC]::Collect();[GC]::WaitForPendingFinalizers()
 $hive.DeleteSubKeyTree($path,$false)
 if($existed){
  & reg.exe import $backup | Out-Null
  if($LASTEXITCODE -ne 0){throw 'Registration restore failed'}
  $verify=Join-Path $EvidenceDirectory 'queue-word-hkcu-restored.reg'
  & reg.exe export ('HKCU\'+$path) $verify /y | Out-Null
  if((Get-FileHash $verify).Hash -ne (Get-FileHash $backup).Hash){throw 'Restoration mismatch'}
 }else{
  $check=$hive.OpenSubKey($path)
  if($check){$check.Dispose();throw 'Temporary override remains'}
 }
 Write-Output 'Word HKCU COM restored and verified; no HKLM or settings changes.'
 Stop-Transcript | Out-Null
}
