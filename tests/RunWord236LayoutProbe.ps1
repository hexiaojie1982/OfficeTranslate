param([Parameter(Mandatory=$true)][string]$EvidenceDirectory)
$ErrorActionPreference='Stop'
'RUNNING' | Set-Content (Join-Path $EvidenceDirectory 'word-layout.exit')
if(Get-Process WINWORD -ErrorAction SilentlyContinue){throw 'Close Word before isolated validation'}
$root=Split-Path $PSScriptRoot -Parent
$dll=Join-Path $root 'src\OfficeTranslate.WordAddIn\bin\Release\net48\OfficeTranslate.WordAddIn.dll'
$assembly=[Reflection.Assembly]::LoadFrom($dll)
[void][Reflection.Assembly]::LoadFrom((Join-Path (Split-Path $dll -Parent) 'OfficeTranslate.Core.dll'))
$serviceType=$assembly.GetType('OfficeTranslate.WordAddIn.WordTranslationService',$true)
$flags=[Reflection.BindingFlags]'Instance,NonPublic'
$word=$null
$doc=$null
try {
    $word=New-Object -ComObject Word.Application
    $word.Visible=$false
    $word.DisplayAlerts=0
    $doc=$word.Documents.Add()
    $service=$serviceType.GetConstructors()[0].Invoke([object[]]@($word.PSObject.BaseObject))
    $serviceType.GetField('_taskDocument',$flags).SetValue($service,$doc.PSObject.BaseObject)
    $method=$serviceType.GetMethod('ShrinkFontToFit',$flags)
    $results=@("assembly=$($assembly.GetName().Version) word_version=$($word.Version) word_build=$($word.Build)")
    foreach($case in @('below_six','at_one','still_overflow')) {
        $box=$doc.Shapes.AddTextbox(1,72,72,30,20)
        $box.TextFrame.MarginLeft=2; $box.TextFrame.MarginRight=2
        $box.TextFrame.MarginTop=1; $box.TextFrame.MarginBottom=1
        $box.TextFrame.WordWrap=-1; $box.TextFrame.AutoSize=0
        $box.TextFrame.TextRange.ParagraphFormat.SpaceBefore=0
        $box.TextFrame.TextRange.ParagraphFormat.SpaceAfter=0
        $box.TextFrame.TextRange.Text=if($case -eq 'still_overflow'){'W'*30}else{'ABC '*7}
        $plan=[OfficeTranslate.Core.ImageOverlayPlanner]::Plan(0,0,1000,1000,1500,1000,30,20,$box.TextFrame.TextRange.Text,0)
        $startSize=if($case -eq 'at_one'){[single]1}else{[single]18}
        $arguments=[object[]]@($box.PSObject.BaseObject,$startSize,'')
        $fit=$method.Invoke($service,$arguments)
        $actualFont=$box.TextFrame.TextRange.Font.Size
        $overflow=$box.TextFrame.Overflowing
        $results+="case=$case plan=$($plan.Verdict) planned_font=$($plan.FontSize) fit=$fit font=$actualFont overflow=$overflow width=$($box.Width) height=$($box.Height)"
        $results | Set-Content (Join-Path $EvidenceDirectory 'word-layout.txt') -Encoding UTF8
        if($case -eq 'still_overflow') {
            if($fit -or -not $overflow -or $actualFont -ne 1){throw 'Expected explicit overflow at 1pt'}
        } else {
            if(-not $fit -or $overflow -or $actualFont -lt 1 -or $actualFont -ge 6){throw 'Expected actual Word fit below 6pt'}
            if($case -eq 'at_one' -and $actualFont -ne 1){throw 'Word did not retain 1pt'}
        }
        if($box.Width -ne 30 -or $box.Height -ne 20){throw 'The fixed overlay rectangle changed'}
        $box.Delete()
    }
    # Investigation only: isolate host paragraph layout from the planner and
    # geometric trust gates. Do not change document defaults or product policy.
    foreach($mode in @('inherited','grid_value_0','grid_value_minus1','exact_spacing','wider')) {
        $width=if($mode -eq 'wider'){60}else{30}
        $box=$doc.Shapes.AddTextbox(1,72,72,$width,20)
        $box.TextFrame.MarginLeft=2; $box.TextFrame.MarginRight=2
        $box.TextFrame.MarginTop=1; $box.TextFrame.MarginBottom=1
        $box.TextFrame.WordWrap=-1; $box.TextFrame.AutoSize=0
        $range=$box.TextFrame.TextRange
        $range.Text='W'*30
        $range.ParagraphFormat.SpaceBefore=0; $range.ParagraphFormat.SpaceAfter=0
        if($mode -eq 'grid_value_0' -or $mode -eq 'exact_spacing'){$range.ParagraphFormat.DisableLineHeightGrid=0}
        if($mode -eq 'grid_value_minus1'){$range.ParagraphFormat.DisableLineHeightGrid=-1}
        if($mode -eq 'exact_spacing'){$range.ParagraphFormat.LineSpacingRule=4; $range.ParagraphFormat.LineSpacing=1.3}
        $range.Font.Size=1
        $doc.Repaginate()
        $results+="investigation=$mode font=$($range.Font.Size) overflow=$($box.TextFrame.Overflowing) grid=$($range.ParagraphFormat.DisableLineHeightGrid) line_rule=$($range.ParagraphFormat.LineSpacingRule) line_pt=$($range.ParagraphFormat.LineSpacing) word_wrap=$($range.ParagraphFormat.WordWrap)"
        $box.Delete()
    }
    $results | Set-Content (Join-Path $EvidenceDirectory 'word-layout.txt') -Encoding UTF8
} catch {
    # Only synthetic content is used. Still avoid exception messages/content.
    $details=@('line='+$_.InvocationInfo.ScriptLineNumber)
    $errorItem=$_.Exception
    while($errorItem){$details+=($errorItem.GetType().FullName+' HResult='+$errorItem.HResult+' stack='+$errorItem.StackTrace); $errorItem=$errorItem.InnerException}
    $details | Set-Content (Join-Path $EvidenceDirectory 'word-layout.error')
    'FAIL' | Set-Content (Join-Path $EvidenceDirectory 'word-layout.exit')
    throw
} finally {
    $dontSave=0
    try {
        if($doc){$doc.Close([ref]$dontSave); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($doc)}
    } finally {
        if($word){$word.Quit([ref]$dontSave); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($word)}
    }
}
'PASS' | Set-Content (Join-Path $EvidenceDirectory 'word-layout.exit')
