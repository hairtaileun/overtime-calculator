[CmdletBinding()]
param(
    [string]$WorkbookPath = (Join-Path (Split-Path -Parent $PSScriptRoot) '2026초과v3.2.xlsm'),
    [string]$JsonOut = (Join-Path (Split-Path -Parent $PSScriptRoot) 'audit-evidence.json'),
    [string]$TextOut = (Join-Path (Split-Path -Parent $PSScriptRoot) 'audit-evidence.txt'),
    [switch]$SkipExcelRuntime
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$ExpectedWorkbookSha256 = 'e509d16a4dc832e26dd2f143f80b7ebc13dd3abd7c4b25e20df69fe05b0c3d80'
$ExpectedVbaSha256 = '04e91b0b788b387dddc3ffb8ebc51aa376bac349162bdff309805b93f2289d03'
$ExpectedSheets = @('설정','급량','1월','2월','3월','4월','5월','6월','7월','8월','9월','10월','11월','12월')
$McNs = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
$RelNs = 'http://schemas.openxmlformats.org/package/2006/relationships'
$SpreadsheetNs = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

$script:Failed = $false
$script:Checks = [System.Collections.Generic.List[object]]::new()
$StartedUtc = [DateTime]::UtcNow

function Add-Check {
    param([string]$Name,[bool]$Pass,[object]$Details=$null,[bool]$Required=$true)
    if ($Required -and -not $Pass) { $script:Failed = $true }
    $script:Checks.Add([pscustomobject][ordered]@{name=$Name;pass=$Pass;required=$Required;details=$Details}) | Out-Null
    $tag = if($Pass){'PASS'}elseif($Required){'FAIL'}else{'WARN'}
    Write-Host ("AUDIT_CHECK={0}|{1}" -f $tag,$Name)
    if($null -ne $Details){
        $rendered=if($Details -is [string]){$Details}else{$Details|ConvertTo-Json -Depth 8 -Compress}
        Write-Host ("AUDIT_DETAIL={0}|{1}" -f $Name,$rendered)
    }
}

function Read-EntryBytes([System.IO.Compression.ZipArchiveEntry]$Entry){
    $s=$Entry.Open()
    try{
        $m=[System.IO.MemoryStream]::new()
        try{$s.CopyTo($m);return $m.ToArray()}finally{$m.Dispose()}
    }finally{$s.Dispose()}
}

function Parse-Xml([byte[]]$Bytes,[string]$Part){
    $settings=[System.Xml.XmlReaderSettings]::new()
    $settings.DtdProcessing=[System.Xml.DtdProcessing]::Prohibit
    $settings.XmlResolver=$null
    $mem=[System.IO.MemoryStream]::new($Bytes,$false)
    try{
        $reader=[System.Xml.XmlReader]::Create($mem,$settings)
        try{
            $doc=[System.Xml.XmlDocument]::new()
            $doc.PreserveWhitespace=$true
            $doc.XmlResolver=$null
            $doc.Load($reader)
            return $doc
        }catch{throw ("XML_PARSE_FAILED:{0}:{1}" -f $Part,$_.Exception.Message)}
        finally{$reader.Dispose()}
    }finally{$mem.Dispose()}
}

function Get-RelSource([string]$Part){
    $p=$Part.Replace('\','/')
    if ($p -eq '_rels/.rels') { return '' }
    $marker='/_rels/'
    $i=$p.IndexOf($marker,[StringComparison]::Ordinal)
    if ($i -lt 0 -or -not $p.EndsWith('.rels',[StringComparison]::OrdinalIgnoreCase)) { throw "INVALID_RELS_PART:$Part" }
    $prefix=$p.Substring(0,$i)
    $name=$p.Substring($i+$marker.Length)
    $sourceName=$name.Substring(0,$name.Length-5)
    if([string]::IsNullOrEmpty($prefix)){return $sourceName}
    return "$prefix/$sourceName"
}

function Resolve-Rel([string]$Source,[string]$Target){
    if ($Target.StartsWith('/')) { return $Target.TrimStart('/') }
    $base=if([string]::IsNullOrEmpty($Source)){[Uri]'http://package/'}else{[Uri]("http://package/"+$Source)}
    return [Uri]::UnescapeDataString(([Uri]::new($base,$Target)).AbsolutePath.TrimStart('/'))
}

function Has-Motw([string]$Path){
    try{return $null -ne (Get-Item -LiteralPath $Path -Stream Zone.Identifier -ErrorAction Stop)}catch{return $false}
}

function Release-Com([object]$Object){
    if ($null -ne $Object -and [Runtime.InteropServices.Marshal]::IsComObject($Object)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object)
    }
}

function Near([double]$Actual,[double]$Expected,[double]$Tolerance=0.0000001){
    return [Math]::Abs($Actual-$Expected) -le $Tolerance
}

function Wait-Calculation([object]$Excel){
    $deadline=[DateTime]::UtcNow.AddMinutes(2)
    while([int]$Excel.CalculationState -ne 0){
        if([DateTime]::UtcNow -gt $deadline){throw 'EXCEL_CALCULATION_TIMEOUT'}
        Start-Sleep -Milliseconds 50
    }
}

$WorkbookPath=[IO.Path]::GetFullPath($WorkbookPath)
$JsonOut=[IO.Path]::GetFullPath($JsonOut)
$TextOut=[IO.Path]::GetFullPath($TextOut)
if(-not(Test-Path -LiteralPath $WorkbookPath -PathType Leaf)){throw "WORKBOOK_NOT_FOUND:$WorkbookPath"}

$WorkbookSha256=(Get-FileHash -LiteralPath $WorkbookPath -Algorithm SHA256).Hash.ToLowerInvariant()
Add-Check 'WORKBOOK_SHA256_EXPECTED' ($WorkbookSha256 -ceq $ExpectedWorkbookSha256) @{actual=$WorkbookSha256;expected=$ExpectedWorkbookSha256}
Add-Check 'SOURCE_MOTW_OBSERVATION' $true @{motw_present=(Has-Motw $WorkbookPath)} $false

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
$fs=[IO.File]::Open($WorkbookPath,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::Read)
$zip=[IO.Compression.ZipArchive]::new($fs,[IO.Compression.ZipArchiveMode]::Read,$false)
$entries=@{}
try{
    $duplicates=@($zip.Entries|Group-Object FullName|Where-Object Count -gt 1|ForEach-Object Name)
    Add-Check 'ZIP_NO_DUPLICATE_ENTRIES' ($duplicates.Count -eq 0) $duplicates

    $readFailures=[System.Collections.Generic.List[string]]::new()
    foreach($entry in $zip.Entries){
        if ([string]::IsNullOrEmpty($entry.FullName) -or $entry.FullName.EndsWith('/')) { continue }
        $entries[$entry.FullName]=$entry
        try{[void](Read-EntryBytes $entry)}catch{$readFailures.Add("$($entry.FullName):$($_.Exception.Message)")}
    }
    Add-Check 'ZIP_ALL_ENTRIES_READABLE' ($readFailures.Count -eq 0) @($readFailures)

    $required=@('[Content_Types].xml','xl/workbook.xml','xl/_rels/workbook.xml.rels','xl/vbaProject.bin')
    $missing=@($required|Where-Object{-not$entries.ContainsKey($_)})
    Add-Check 'OOXML_REQUIRED_PARTS_PRESENT' ($missing.Count -eq 0) $missing

    $docs=@{}
    $parseFailures=[System.Collections.Generic.List[string]]::new()
    foreach($part in @($entries.Keys|Where-Object{$_ -match '\.(xml|rels|vml)$'}|Sort-Object)){
        try{$docs[$part]=Parse-Xml (Read-EntryBytes $entries[$part]) $part}catch{$parseFailures.Add($_.Exception.Message)}
    }
    Add-Check 'XML_RELS_VML_STRICT_PARSE' ($parseFailures.Count -eq 0) @($parseFailures)

    $nsFailures=[System.Collections.Generic.List[string]]::new()
    foreach($part in $docs.Keys){
        foreach($element in @($docs[$part].SelectNodes('//*'))){
            if ($null -eq $element -or $element.NodeType -ne [System.Xml.XmlNodeType]::Element) { continue }
            $ign=$element.GetAttribute('Ignorable',$McNs)
            foreach($prefix in ($ign -split '\s+'|Where-Object{$_})){
                if([string]::IsNullOrEmpty($element.GetNamespaceOfPrefix($prefix))){$nsFailures.Add(("{0}:Ignorable:{1}" -f $part,$prefix))}
            }
            if ($element.NamespaceURI -eq $McNs -and $element.LocalName -eq 'Choice') {
                foreach($prefix in ($element.GetAttribute('Requires') -split '\s+'|Where-Object{$_})){
                    if([string]::IsNullOrEmpty($element.GetNamespaceOfPrefix($prefix))){$nsFailures.Add(("{0}:Requires:{1}" -f $part,$prefix))}
                }
            }
        }
    }
    Add-Check 'OOXML_MARKUP_COMPAT_NAMESPACES' ($nsFailures.Count -eq 0) @($nsFailures)

    $relFailures=[System.Collections.Generic.List[string]]::new()
    foreach($relPart in @($entries.Keys|Where-Object{$_ -match '\.rels$'})){
        $doc=$docs[$relPart]
        $mgr=[System.Xml.XmlNamespaceManager]::new($doc.NameTable);$mgr.AddNamespace('r',$RelNs)
        $source=Get-RelSource $relPart
        foreach($rel in @($doc.SelectNodes('/r:Relationships/r:Relationship',$mgr))){
            if ($rel.GetAttribute('TargetMode') -eq 'External') { continue }
            try{
                $target=Resolve-Rel $source $rel.GetAttribute('Target')
                if(-not$entries.ContainsKey($target)){$relFailures.Add(("{0}:{1}->{2}" -f $relPart,$rel.GetAttribute('Id'),$target))}
            }catch{$relFailures.Add(("{0}:{1}:{2}" -f $relPart,$rel.GetAttribute('Id'),$_.Exception.Message))}
        }
    }
    Add-Check 'OOXML_INTERNAL_RELATIONSHIP_TARGETS' ($relFailures.Count -eq 0) @($relFailures)

    $ct=$docs['[Content_Types].xml']
    $mgr=[System.Xml.XmlNamespaceManager]::new($ct.NameTable);$mgr.AddNamespace('ct','http://schemas.openxmlformats.org/package/2006/content-types')
    $defaults=@{};$overrides=@{}
    foreach($n in @($ct.SelectNodes('/ct:Types/ct:Default',$mgr))){$defaults[$n.GetAttribute('Extension').ToLowerInvariant()]=$n.GetAttribute('ContentType')}
    foreach($n in @($ct.SelectNodes('/ct:Types/ct:Override',$mgr))){$overrides[$n.GetAttribute('PartName').TrimStart('/')]=$n.GetAttribute('ContentType')}
    $ctFailures=[System.Collections.Generic.List[string]]::new()
    foreach($part in $entries.Keys){
        if ($part -eq '[Content_Types].xml') { continue }
        $ext=[IO.Path]::GetExtension($part).TrimStart('.').ToLowerInvariant()
        if (-not $overrides.ContainsKey($part) -and ([string]::IsNullOrEmpty($ext) -or -not $defaults.ContainsKey($ext))) { $ctFailures.Add($part) }
    }
    Add-Check 'OOXML_CONTENT_TYPES_COMPLETE' ($ctFailures.Count -eq 0) @($ctFailures)

    $sha=[Security.Cryptography.SHA256]::Create()
    try{$vbaHash=([BitConverter]::ToString($sha.ComputeHash((Read-EntryBytes $entries['xl/vbaProject.bin'])))).Replace('-','').ToLowerInvariant()}finally{$sha.Dispose()}
    Add-Check 'VBA_PROJECT_SHA256' ($vbaHash -ceq $ExpectedVbaSha256) @{actual=$vbaHash;expected=$ExpectedVbaSha256}

    $wbDoc=$docs['xl/workbook.xml']
    $mgr=[System.Xml.XmlNamespaceManager]::new($wbDoc.NameTable);$mgr.AddNamespace('x',$SpreadsheetNs)
    $sheetNames=@($wbDoc.SelectNodes('/x:workbook/x:sheets/x:sheet',$mgr)|ForEach-Object{$_.GetAttribute('name')})
    Add-Check 'WORKBOOK_EXPECTED_SHEETS' ($sheetNames.Count -eq $ExpectedSheets.Count -and -not(Compare-Object $ExpectedSheets $sheetNames -SyncWindow 0)) $sheetNames
    $calc=$wbDoc.SelectSingleNode('/x:workbook/x:calcPr',$mgr)
    $calcPass = (
        $null -ne $calc -and
        $calc.GetAttribute('calcMode') -eq 'auto' -and
        $calc.GetAttribute('fullCalcOnLoad') -eq '1' -and
        $calc.GetAttribute('forceFullCalc') -eq '1'
    )
    Add-Check 'WORKBOOK_FULL_RECALC_FLAGS' $calcPass $(if ($null -eq $calc) { $null } else {@{calcMode=$calc.GetAttribute('calcMode');fullCalcOnLoad=$calc.GetAttribute('fullCalcOnLoad');forceFullCalc=$calc.GetAttribute('forceFullCalc')}})

    $macroBound=$false
    foreach($part in @($entries.Keys|Where-Object{$_ -like 'xl/worksheets/*.xml'})){
        if([Text.Encoding]::UTF8.GetString((Read-EntryBytes $entries[$part])) -match 'macro="\[0\]!RegenerateAllMonthB"'){$macroBound=$true;break}
    }
    Add-Check 'REGENERATE_BUTTON_MACRO_BINDING' $macroBound

    $formulas=[System.Collections.Generic.List[string]]::new()
    foreach($part in @($entries.Keys|Where-Object{$_ -like 'xl/worksheets/*.xml'})){
        $doc=$docs[$part];$mgr=[System.Xml.XmlNamespaceManager]::new($doc.NameTable);$mgr.AddNamespace('x',$SpreadsheetNs)
        foreach($f in @($doc.SelectNodes('//x:f',$mgr))){if ($f.InnerText) { $formulas.Add($f.InnerText) }}
    }
    $f=@($formulas)
    $counts=[ordered]@{
        total=$f.Count
        holiday_meal_3660=@($f|Where-Object{$_ -like '*>=3660*'}).Count
        old_7260=@($f|Where-Object{$_ -like '*7260*'}).Count
        old_holiday_4h_cap=@($f|Where-Object{$_ -like '*MIN(TIME(4,0,0)*'}).Count
        old_9999=@($f|Where-Object{$_ -like '*9999*'}).Count
        meal_cap_b7=@($f|Where-Object{$_ -like '*설정!$B$7*'}).Count
        quarter_cap_b8=@($f|Where-Object{$_ -like '*설정!$B$8*'}).Count
        error_tokens=@($f|Where-Object{$_ -match '#REF!|#DIV/0!|#VALUE!|#N/A|#NAME\?'}).Count
        external_books=@($f|Where-Object{$_ -match '\[[^\]]+\]'}).Count
    }
    Add-Check 'FORMULA_TOTAL_EXPECTED' ($counts.total -eq 1278) $counts
    Add-Check 'HOLIDAY_MEAL_FORMULA_ALL_366_DAYS' ($counts.holiday_meal_3660 -eq 366 -and $counts.old_7260 -eq 0) $counts
    Add-Check 'NO_LEGACY_DAILY_CAP_BYPASS' ($counts.old_holiday_4h_cap -eq 0 -and $counts.old_9999 -eq 0) $counts
    Add-Check 'CONFIGURABLE_MONTHLY_MEAL_CAP_REFERENCES' ($counts.meal_cap_b7 -eq 36) $counts.meal_cap_b7
    Add-Check 'CONFIGURABLE_QUARTER_CAP_REFERENCES' ($counts.quarter_cap_b8 -eq 12) $counts.quarter_cap_b8
    Add-Check 'NO_LITERAL_FORMULA_ERRORS_OR_EXTERNAL_BOOKS' ($counts.error_tokens -eq 0 -and $counts.external_books -eq 0) $counts
}finally{$zip.Dispose();$fs.Dispose()}

if(-not$SkipExcelRuntime){
    $excel=$null;$book=$null;$round=$null;$settings=$null;$jan=$null;$march=$null
    $root=Join-Path ([IO.Path]::GetTempPath()) ('overtime-audit-'+[guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $root -Force|Out-Null
    try{
        $runtime=Join-Path $root 'runtime.xlsm';$roundPath=Join-Path $root 'roundtrip.xlsm'
        Copy-Item $WorkbookPath $runtime -Force;Copy-Item $WorkbookPath $roundPath -Force
        Unblock-File $runtime;Unblock-File $roundPath
        Add-Check 'TEMP_RUNTIME_COPY_UNBLOCKED' (-not(Has-Motw $runtime))
        try{$excel=New-Object -ComObject Excel.Application;Add-Check 'EXCEL_COM_AVAILABLE' $true @{version=[string]$excel.Version;build=[string]$excel.Build;hwnd=[string]$excel.Hwnd}}
        catch{Add-Check 'EXCEL_COM_AVAILABLE' $false $_.Exception.Message;throw}
        $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.AskToUpdateLinks=$false
        try{$excel.AutomationSecurity=1;Add-Check 'EXCEL_TEMP_AUTOMATION_SECURITY' $true 'msoAutomationSecurityLow in isolated Excel process'}
        catch{Add-Check 'EXCEL_TEMP_AUTOMATION_SECURITY' $false $_.Exception.Message}
        $m=[Type]::Missing
        try{$book=$excel.Workbooks.Open($runtime,0,$false,$m,$m,$m,$true,$m,$m,$false,$false,$m,$false,$true,0);Add-Check 'EXCEL_OPEN_NORMAL_NO_REPAIR_REQUEST' ($null -ne $book) @{read_only=[bool]$book.ReadOnly}}
        catch{Add-Check 'EXCEL_OPEN_NORMAL_NO_REPAIR_REQUEST' $false $_.Exception.Message;throw}

        $runtimeNames=@();for($i=1;$i-le$book.Worksheets.Count;$i++){$runtimeNames += [string]$book.Worksheets.Item($i).Name}
        Add-Check 'EXCEL_RUNTIME_EXPECTED_SHEETS' ($runtimeNames.Count -eq 14 -and -not (Compare-Object $ExpectedSheets $runtimeNames -SyncWindow 0)) $runtimeNames
        $excel.CalculateFullRebuild();Wait-Calculation $excel
        Add-Check 'EXCEL_FULL_CALCULATE_REBUILD' $true @{state=[int]$excel.CalculationState}

        $errorCells=[System.Collections.Generic.List[string]]::new()
        for($i=1;$i-le$book.Worksheets.Count;$i++){
            $sheet=$null;$used=$null;$errs=$null
            try{
                $sheet=$book.Worksheets.Item($i);$used=$sheet.UsedRange
                try{$errs=$used.SpecialCells(-4123,16)}catch{$errs=$null}
                if ($null -ne $errs) { foreach ($area in @($errs.Areas)) { $errorCells.Add("$($sheet.Name)!$($area.Address($false,$false))") } }
            }finally{Release-Com $errs;Release-Com $used;Release-Com $sheet}
        }
        Add-Check 'EXCEL_NO_FORMULA_ERROR_CELLS_AFTER_RECALC' ($errorCells.Count -eq 0) @($errorCells)

        $settings=$book.Worksheets.Item('설정');$jan=$book.Worksheets.Item('1월');$march=$book.Worksheets.Item('3월')
        $settings.Range('B2').Value2=8;$settings.Range('B3').Value2=67;$settings.Range('B4').Value2=6/24.0;$settings.Range('B5').Value2=23/24.0;$settings.Range('B6').Value2='예';$settings.Range('B7').Value2=20;$settings.Range('B8').Value2=177
        $jan.Range('B2').Value2='휴일';$jan.Range('C2').Value2=6/24.0;$jan.Range('D2').Value2=23/24.0
        $excel.CalculateFullRebuild();Wait-Calculation $excel
        $v=[double]$jan.Range('E2').Value2
        Add-Check 'SCENARIO_HOLIDAY_DAILY_CAP_8H' (Near $v (8/24.0)) @{actual_days=$v}

        $settings.Range('B2').Value2=6;$excel.CalculateFullRebuild();Wait-Calculation $excel
        $v=[double]$jan.Range('E2').Value2
        Add-Check 'SCENARIO_HOLIDAY_ARBITRARY_DAILY_CAP_6H' (Near $v (6/24.0)) @{actual_days=$v}

        $settings.Range('B2').Value2=8;$jan.Range('C2').Value2=6/24.0;$jan.Range('D2').Value2=7/24.0
        $excel.CalculateFullRebuild();Wait-Calculation $excel
        $meal=[int]$jan.Range('F2').Value2;$ov=[double]$jan.Range('E2').Value2
        Add-Check 'SCENARIO_HOLIDAY_MEAL_60MIN_ZERO' ($meal -eq 0 -and (Near $ov (1/24.0))) @{meal=$meal;overtime_days=$ov}

        $jan.Range('D2').Value2=421/1440.0;$excel.CalculateFullRebuild();Wait-Calculation $excel
        $meal=[int]$jan.Range('F2').Value2;$ov=[double]$jan.Range('E2').Value2
        Add-Check 'SCENARIO_HOLIDAY_MEAL_61MIN_ONE' ($meal -eq 1 -and (Near $ov (61/1440.0))) @{meal=$meal;overtime_days=$ov}

        $jan.Range('E2:E32').Value2=0
        foreach($hours in @(66,67,68)){
            $jan.Range('E2').Value2=$hours/24.0;$excel.CalculateFullRebuild();Wait-Calculation $excel
            $total=[double]$jan.Range('E33').Value2*24;$under=[double]$jan.Range('E34').Value2*24;$over=[double]$jan.Range('E35').Value2*24
            Add-Check ("SCENARIO_MONTH_CAP_{0}H" -f $hours) ((Near $total $hours) -and (Near $under ([Math]::Max(0,67-$hours))) -and (Near $over ([Math]::Max(0,$hours-67)))) @{total=$total;under=$under;over=$over}
        }

        foreach($count in @(19,20,21)){
            $jan.Range('F2:G32').Value2=0;$jan.Range('F2').Value2=$count;$excel.CalculateFullRebuild();Wait-Calculation $excel
            $accepted=[double]$jan.Range('G33').Value2;$under=[double]$jan.Range('G34').Value2;$over=[double]$jan.Range('G35').Value2
            Add-Check ("SCENARIO_MONTH_MEAL_CAP_{0}" -f $count) ((Near $accepted ([Math]::Min($count,20))) -and (Near $under ([Math]::Max(0,20-[Math]::Min($count,20)))) -and (Near $over ([Math]::Max(0,$count-20)))) @{accepted=$accepted;under=$under;over=$over}
        }

        foreach($hours in @(176,177,178)){
            $march.Range('L2').Value2=$hours/24.0;$march.Range('L3:L4').Value2=0;$excel.CalculateFullRebuild();Wait-Calculation $excel
            $total=[double]$march.Range('L5').Value2*24;$under=[double]$march.Range('L6').Value2*24;$over=[double]$march.Range('L7').Value2*24
            Add-Check ("SCENARIO_QUARTER_CAP_{0}H" -f $hours) ((Near $total $hours) -and (Near $under ([Math]::Max(0,177-$hours))) -and (Near $over ([Math]::Max(0,$hours-177)))) @{total=$total;under=$under;over=$over}
        }

        $settings.Range('B1').Value2=2027
        try{
            [void]$excel.Run("'$($book.Name)'!RegenerateAllMonthB");$excel.CalculateFullRebuild();Wait-Calculation $excel
            $d1=[string]$jan.Range('B2').Value2;$d2=[string]$jan.Range('B3').Value2;$d3=[string]$jan.Range('B4').Value2
            Add-Check 'MACRO_REGENERATE_2027_WEEKEND_CLASSIFICATION' ($d1 -eq '평일' -and $d2 -eq '휴일' -and $d3 -eq '휴일') @{jan1=$d1;jan2=$d2;jan3=$d3}
        }catch{Add-Check 'MACRO_REGENERATE_2027_WEEKEND_CLASSIFICATION' $false $_.Exception.Message}

        $book.Close($false);Release-Com $march;$march=$null;Release-Com $jan;$jan=$null;Release-Com $settings;$settings=$null;Release-Com $book;$book=$null

        try{
            $round=$excel.Workbooks.Open($roundPath,0,$false,$m,$m,$m,$true,$m,$m,$false,$false,$m,$false,$true,0)
            $excel.CalculateFullRebuild();Wait-Calculation $excel;$round.Save();$round.Close($false);Release-Com $round;$round=$null
            $round=$excel.Workbooks.Open($roundPath,0,$true,$m,$m,$m,$true,$m,$m,$false,$false,$m,$false,$true,0)
            Add-Check 'EXCEL_SAVE_CLOSE_REOPEN_ROUNDTRIP' ([int]$round.Worksheets.Count -eq 14) @{worksheet_count=[int]$round.Worksheets.Count}
        }catch{Add-Check 'EXCEL_SAVE_CLOSE_REOPEN_ROUNDTRIP' $false $_.Exception.Message}
    }finally{
        if ($null -ne $round) { try { $round.Close($false) } catch {} }
        if ($null -ne $book) { try { $book.Close($false) } catch {} }
        Release-Com $march;Release-Com $jan;Release-Com $settings;Release-Com $round;Release-Com $book
        if ($null -ne $excel) { try { $excel.Quit() } catch {}; Release-Com $excel }
        [GC]::Collect();[GC]::WaitForPendingFinalizers();[GC]::Collect();[GC]::WaitForPendingFinalizers()
        Remove-Item $root -Recurse -Force -ErrorAction SilentlyContinue
    }
}else{Add-Check 'EXCEL_RUNTIME_SKIPPED' $true 'SkipExcelRuntime requested' $false}

$final=if($script:Failed){'FAIL'}else{'PASS'}
$payload=[ordered]@{schema='overtime-calculator-audit/v1';final=$final;workbook=$WorkbookPath;workbook_sha256=$WorkbookSha256;expected_workbook_sha256=$ExpectedWorkbookSha256;started_utc=$StartedUtc.ToString('o');finished_utc=[DateTime]::UtcNow.ToString('o');computer=$env:COMPUTERNAME;user=[Environment]::UserName;runner_name=$env:RUNNER_NAME;checks=@($script:Checks)}
New-Item -ItemType Directory -Path (Split-Path -Parent $JsonOut) -Force|Out-Null
$payload|ConvertTo-Json -Depth 12|Set-Content $JsonOut -Encoding utf8
$lines=[System.Collections.Generic.List[string]]::new();$lines.Add("OVERTIME_CALCULATOR_AUDIT=$final")|Out-Null;$lines.Add("FILE_SHA256=$WorkbookSha256")|Out-Null
foreach($c in $script:Checks){$state=if($c.pass){'PASS'}elseif($c.required){'FAIL'}else{'WARN'};$lines.Add(("{0}={1}"-f$c.name,$state))|Out-Null}
$lines.Add("FINAL_RESULT=$final")|Out-Null;$lines|Set-Content $TextOut -Encoding utf8;$lines|ForEach-Object{Write-Host $_}
Write-Host "AUDIT_JSON=$JsonOut";Write-Host "AUDIT_TEXT=$TextOut"
if ($script:Failed) { exit 90 } else { exit 0 }
