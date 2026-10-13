param(
    [Parameter(Mandatory=$true)][string]$EvidenceDirectory,
    [string]$AssemblyDirectory,
    [switch]$ExpectBaseline
)
$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
if(-not $AssemblyDirectory){$AssemblyDirectory=Join-Path $root 'src\OfficeTranslate.WordAddIn\bin\Release\net48'}
$prefix=if($ExpectBaseline){'baseline'}else{'candidate'}
$exitFile=Join-Path $EvidenceDirectory "$prefix-word.exit"
$resultFile=Join-Path $EvidenceDirectory "$prefix-word.txt"
'RUNNING' | Set-Content $exitFile
$word=$null; $doc=$null; $passed=$false
$results=New-Object 'System.Collections.Generic.List[string]'
function Assert-That([bool]$condition,[string]$label){if(-not $condition){throw [InvalidOperationException]::new($label)}}
function Get-BodySnapshot {
    $formats=@($doc.Styles.Item(-1).ParagraphFormat)
    foreach($paragraph in $doc.Content.Paragraphs){$formats+=$paragraph.Format}
    $values=@($doc.Content.Text)
    foreach($p in $formats){
        foreach($property in @('SpaceBefore','SpaceAfter','SpaceBeforeAuto','SpaceAfterAuto','LineSpacingRule','LineSpacing','DisableLineHeightGrid','LeftIndent','RightIndent','FirstLineIndent','Alignment','KeepTogether','KeepWithNext','PageBreakBefore','WidowControl')){
            $values+=$p.$property
        }
    }
    return ($values | ConvertTo-Json -Compress)
}
function New-Box {
    $box=$doc.Shapes.AddTextbox(1,72,72,30,20)
    $box.RelativeHorizontalPosition=1; $box.RelativeVerticalPosition=1
    $box.WrapFormat.Type=3; $box.Left=72; $box.Top=72
    $box.TextFrame.MarginLeft=2; $box.TextFrame.MarginRight=2
    $box.TextFrame.MarginTop=1; $box.TextFrame.MarginBottom=1
    $box.TextFrame.WordWrap=-1; $box.TextFrame.AutoSize=0
    return $box
}
function Assert-Compact($range) {
    $p=$range.ParagraphFormat
    Assert-That ($p.SpaceBefore -eq 0 -and $p.SpaceAfter -eq 0) 'Paragraph gaps remain'
    Assert-That ($p.SpaceBeforeAuto -eq 0 -and $p.SpaceAfterAuto -eq 0) 'Automatic gaps remain'
    Assert-That ($p.LineSpacingRule -eq 0 -and $p.DisableLineHeightGrid -eq -1) 'Line/grid format incorrect'
    # Verify the confusing COM boolean against the persisted OOXML semantics.
    # Never write the XML/text to the diagnostic output.
    [xml]$xml=$range.WordOpenXML
    $ns=New-Object Xml.XmlNamespaceManager($xml.NameTable)
    $ns.AddNamespace('w','http://schemas.openxmlformats.org/wordprocessingml/2006/main')
    $paragraphs=$xml.SelectNodes('//w:body/w:p',$ns)
    Assert-That ($paragraphs.Count -gt 0) 'No paragraphs in exported story'
    foreach($paragraph in $paragraphs){
        $grid=$paragraph.SelectSingleNode('w:pPr/w:snapToGrid',$ns)
        Assert-That ($null -ne $grid -and $grid.GetAttribute('val',$ns.LookupNamespace('w')) -eq '0') 'snapToGrid must be off for every paragraph'
    }
}
try {
    Assert-That ([Environment]::UserInteractive) 'Run in the interactive desktop session'
    Assert-That (-not (Get-Process WINWORD -ErrorAction SilentlyContinue)) 'Close Word before isolated validation'
    $assembly=[Reflection.Assembly]::LoadFrom((Join-Path $AssemblyDirectory 'OfficeTranslate.WordAddIn.dll'))
    [void][Reflection.Assembly]::LoadFrom((Join-Path $AssemblyDirectory 'OfficeTranslate.Core.dll'))
    $expectedVersion=if($ExpectBaseline){'2.1.36.0'}else{'2.1.37.0'}
    Assert-That ($assembly.GetName().Version.ToString() -eq $expectedVersion) 'Unexpected DLL version'
    $serviceType=$assembly.GetType('OfficeTranslate.WordAddIn.WordTranslationService',$true)
    $flags=[Reflection.BindingFlags]'Instance,NonPublic'
    $word=New-Object -ComObject Word.Application
    $word.Visible=$false; $word.DisplayAlerts=0
    $doc=$word.Documents.Add()
    $doc.Content.Text='Synthetic body unchanged by OCR overlay formatting.'
    $bodySnapshot=Get-BodySnapshot
    $service=$serviceType.GetConstructors()[0].Invoke([object[]]@($word.PSObject.BaseObject))
    $serviceType.GetField('_taskDocument',$flags).SetValue($service,$doc.PSObject.BaseObject)
    $shrink=$serviceType.GetMethod('ShrinkFontToFit',$flags)
    $results.Add("assembly=$($assembly.GetName().Version) word=$($word.Version) build=$($word.Build) session=$([Diagnostics.Process]::GetCurrentProcess().SessionId)")
    foreach($case in @('grid','paragraph_spacing','auto_spacing','exact_spacing','multiple_spacing','mixed_paragraphs','cjk_bilingual','true_overflow')) {
        $box=New-Box
        $range=$box.TextFrame.TextRange
        $range.Text=switch($case){
            'paragraph_spacing' {"AA`rBB`rCC"}
            'auto_spacing' {"AA`rBB`rCC"}
            'mixed_paragraphs' {"AA`rBB`rCC`r"}
            'cjk_bilingual' {"密集小字翻译验证`rDense text translation`r第二行中文与英文 ABC"}
            'true_overflow' {("W`r"*300)+'W'}
            default {'W'*30}
        }
        $p=$range.ParagraphFormat
        $p.SpaceBeforeAuto=0; $p.SpaceAfterAuto=0; $p.SpaceBefore=0; $p.SpaceAfter=0
        $p.LineSpacingRule=0; $p.DisableLineHeightGrid=0
        if($case -in @('paragraph_spacing','auto_spacing','exact_spacing','multiple_spacing','mixed_paragraphs')){$p.DisableLineHeightGrid=-1}
        if($case -eq 'paragraph_spacing'){$p.SpaceBefore=12; $p.SpaceAfter=12}
        if($case -eq 'auto_spacing'){$p.SpaceBeforeAuto=-1; $p.SpaceAfterAuto=-1}
        if($case -eq 'exact_spacing'){$p.LineSpacingRule=4; $p.LineSpacing=24}
        if($case -eq 'multiple_spacing'){$p.LineSpacingRule=5; $p.LineSpacing=120}
        if($case -eq 'mixed_paragraphs'){
            $range.Paragraphs.Item(1).SpaceAfter=24
            $range.Paragraphs.Item(2).LineSpacingRule=4; $range.Paragraphs.Item(2).LineSpacing=24
            $range.Paragraphs.Item(3).DisableLineHeightGrid=0
        }
        $range.Font.Size=if($case -eq 'true_overflow'){18}else{1}
        $doc.Repaginate()
        $before=$box.TextFrame.Overflowing
        Assert-That $before "Fixture must overflow before repair: $case"
        $plan=[OfficeTranslate.Core.ImageOverlayPlanner]::Plan(0,0,1000,1000,1500,1000,30,20,$range.Text,0)
        if($case -ne 'true_overflow'){Assert-That ($plan.Verdict.ToString() -eq 'Place') "Fixture must pass planning: $case"}
        $arguments=[object[]]@($box.PSObject.BaseObject,[single]1,'')
        $fit=$shrink.Invoke($service,$arguments)
        $expectedFit=(-not $ExpectBaseline -and $case -ne 'true_overflow')
        Assert-That ($fit -eq $expectedFit) "Unexpected fit: $case"
        Assert-That ($box.TextFrame.Overflowing -eq (-not $expectedFit)) "Overflow disagrees: $case"
        $results.Add("font_probe=$case initial_range_size=$($range.Font.Size) frame_size=$($box.TextFrame.TextRange.Font.Size) initial_length=$($range.End-$range.Start) frame_length=$($box.TextFrame.TextRange.End-$box.TextFrame.TextRange.Start) story_size=$($box.TextFrame.ContainingRange.Font.Size)"); $results | Set-Content $resultFile -Encoding UTF8
        if(-not $ExpectBaseline -and $case -eq 'true_overflow'){
            # Word may return wdUndefined for the aggregate Font.Size of a
            # long story even when every character is 1pt. Check each one.
            foreach($character in $box.TextFrame.ContainingRange.Characters){
                Assert-That ($character.Font.Size -eq 1) 'Overflow tail did not reach the font floor'
            }
        } elseif($case -ne 'true_overflow'){
            Assert-That ($box.TextFrame.ContainingRange.Font.Size -eq 1) 'Font floor changed or overflow tail omitted'
        }
        Assert-That ($box.Width -eq 30 -and $box.Height -eq 20 -and $box.Left -eq 72 -and $box.Top -eq 72) 'Fixed rectangle changed'
        Assert-That ($box.TextFrame.AutoSize -eq 0 -and $box.TextFrame.WordWrap -eq -1) 'Frame policy changed'
        if(-not $ExpectBaseline){Assert-Compact $box.TextFrame.ContainingRange}
        $results.Add("case=$case plan=$($plan.Verdict) before_overflow=$before fit=$fit after_overflow=$($box.TextFrame.Overflowing) font=$($range.Font.Size) width=$($box.Width) height=$($box.Height) grid=$($p.DisableLineHeightGrid) line_rule=$($p.LineSpacingRule) before_pt=$($p.SpaceBefore) after_pt=$($p.SpaceAfter) before_auto=$($p.SpaceBeforeAuto) after_auto=$($p.SpaceAfterAuto)")
        $results | Set-Content $resultFile -Encoding UTF8
        $box.Delete()
    }
    # Exercise the actual production PlaceOverlay path, not just the fitter.
    $targetType=$serviceType.GetNestedType('ImageTarget',[Reflection.BindingFlags]'NonPublic')
    $target=$targetType.GetConstructors()[0].Invoke([object[]]@($doc.Content.PSObject.BaseObject,[single]72,[single]72,[single]30,[single]20,$null))
    $place=$serviceType.GetMethod('PlaceOverlay',$flags)
    $text='W'*30
    $plan=[OfficeTranslate.Core.ImageOverlayPlanner]::Plan(0,0,1000,1000,1500,1000,30,20,$text,0)
    $arguments=[object[]]@($target,$plan,$text,'synthetic237','')
    $count=$doc.Shapes.Count
    $placed=$place.Invoke($service,$arguments)
    Assert-That ($placed -eq (-not $ExpectBaseline)) 'Unexpected production placement'
    if($placed){
        Assert-That ($doc.Shapes.Count -eq $count+1) 'Missing overlay'
        $overlay=$doc.Shapes.Item($doc.Shapes.Count)
        Assert-That ($overlay.Fill.Visible -eq -1 -and $overlay.Fill.Transparency -eq 0 -and $overlay.Fill.ForeColor.RGB -eq 0xFFFFFF) 'Cover must be opaque white'
        Assert-That ($overlay.Line.Visible -eq 0 -and $overlay.WrapFormat.Type -eq 3) 'Cover stacking changed'
        Assert-That ($overlay.Width -eq $plan.Width -and $overlay.Height -eq $plan.Height -and $overlay.Left -eq 72+$plan.Left -and $overlay.Top -eq 72+$plan.Top) 'Production rectangle changed'
        Assert-That ($overlay.TextFrame.AutoSize -eq 0 -and -not $overlay.TextFrame.Overflowing) 'Production layout failed'
        Assert-Compact $overlay.TextFrame.TextRange
        $results.Add("case=production_place placed=True font=$($overlay.TextFrame.TextRange.Font.Size) width=$($overlay.Width) height=$($overlay.Height) opacity=1 fixed_geometry=True")
        $overlay.Delete()
    } else {
        Assert-That ($doc.Shapes.Count -eq $count) 'Failed overlay was not deleted'
        $results.Add('case=production_place placed=False failed_shape_deleted=True')
    }
    if(-not $ExpectBaseline){
        # Force writeback overflow independently of planning to test cleanup.
        $huge=("W`r"*300)+'W'
        $arguments=[object[]]@($target,$plan,$huge,'synthetic237','')
        Assert-That (-not $place.Invoke($service,$arguments)) 'True overflow accepted'
        Assert-That ($doc.Shapes.Count -eq $count) 'True overflow left a clipped cover'
        $results.Add('case=production_overflow rejected=True failed_shape_deleted=True')
        $dead=New-Box
        $dead.Delete()
        $arguments=[object[]]@($dead.PSObject.BaseObject,[single]1,'')
        Assert-That (-not $shrink.Invoke($service,$arguments)) 'Formatting COM failure accepted'
        Assert-That ($arguments[2] -eq '无法设置覆盖框紧凑段落格式，已降级为旁注。') 'Formatting failure not classified'
        $results.Add('case=format_com_failure rejected=True classified=True')
    }
    # Compare the body after removing test shapes (no text/format logging).
    Assert-That ((Get-BodySnapshot) -ceq $bodySnapshot) 'Body or Normal paragraph formatting changed'
    $results.Add('body_text_and_paragraph_format_and_normal_format_unchanged=True')
    $results | Set-Content $resultFile -Encoding UTF8
    $passed=$true
} catch {
    # Numeric metadata only, never exception Message or document content.
    $details=@("line=$($_.InvocationInfo.ScriptLineNumber)", $_.ScriptStackTrace)
    $err=$_.Exception
    while($err){$details+="type=$($err.GetType().FullName) hresult=$($err.HResult)"; $err=$err.InnerException}
    $details | Set-Content (Join-Path $EvidenceDirectory "$prefix-word.error")
} finally {
    $dontSave=0
    try {
        if($doc){$doc.Close([ref]$dontSave); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($doc)}
    } finally {
        if($word){$word.Quit([ref]$dontSave); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($word)}
    }
}
if($passed){'PASS' | Set-Content $exitFile}else{'FAIL' | Set-Content $exitFile; exit 1}
