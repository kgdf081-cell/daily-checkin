# 极简 xlsx 读写库：仅依赖 .NET，支持内联字符串、下拉校验、条件格式、冻结窗格。

$script:XStyle = @{
    Default      = 0
    Header       = 1
    Title        = 2
    Date         = 3
    Center       = 4
    Left         = 5
    Percent      = 6
    Int          = 7
    Subtitle     = 8
    Wrap         = 9
    Completed    = 10
    Deleted      = 11
    Section      = 12
    Note         = 13
    Big          = 14
    BoldCenter   = 15
    PercentBold  = 16
    LeftBold     = 17
    Green        = 18
    Red          = 19
    Amber        = 20
    HeaderLeft   = 21
    BigPercent   = 22
    TodayHeader  = 23
    GreenCenter  = 24
    RedCenter    = 25
    AmberCenter  = 26
    GrayCenter   = 27
}

function Get-XColName {
    param([int]$Index)
    $s = ''
    $n = $Index
    while ($n -gt 0) {
        $m = ($n - 1) % 26
        $s = [string][char](65 + $m) + $s
        $n = [int](($n - $m - 1) / 26)
    }
    return $s
}

function Get-XColIndex {
    param([string]$Name)
    $n = 0
    foreach ($ch in $Name.ToUpperInvariant().ToCharArray()) {
        $n = $n * 26 + ([int][char]$ch - 64)
    }
    return $n
}

function Split-XRef {
    param([string]$Ref)
    if ($Ref -match '^([A-Za-z]+)(\d+)$') {
        return @{ Col = (Get-XColIndex $Matches[1]); Row = [int]$Matches[2] }
    }
    throw "无效单元格地址: $Ref"
}

function ConvertTo-XSerial {
    param([datetime]$Date)
    return [int]($Date.Date - [datetime]'1899-12-30').Days
}

function ConvertFrom-XSerial {
    param($Serial)
    return ([datetime]'1899-12-30').AddDays([double]$Serial)
}

function New-XCell {
    param(
        [string]$Ref,
        $Value,
        [int]$Style = 0,
        [string]$Type = 'inline'
    )
    return @{ Ref = $Ref; Value = $Value; Style = $Style; Type = $Type }
}

function New-XSheet {
    param([string]$Name)
    return @{
        Name         = $Name
        Rows         = New-Object 'System.Collections.Generic.SortedDictionary[int,object]'
        Cols         = @()
        Freeze       = $null
        ViewCell     = ''
        Merges       = @()
        Validations  = @()
        CondFormats  = @()
        TabSelected  = $false
        DefaultRowHt = 18
        TabColor     = ''
    }
}

function Add-XRow {
    param(
        $Sheet,
        [int]$Row,
        [double]$Height = 0,
        [bool]$Hidden = $false
    )
    if (-not $Sheet.Rows.ContainsKey($Row)) {
        $Sheet.Rows[$Row] = @{
            Row    = $Row
            Cells  = New-Object 'System.Collections.Generic.SortedDictionary[int,object]'
            Height = $Height
            Hidden = $Hidden
        }
    }
    else {
        if ($Height -gt 0) { $Sheet.Rows[$Row].Height = $Height }
        if ($Hidden) { $Sheet.Rows[$Row].Hidden = $true }
    }
}

function Add-XCell {
    param($Sheet, [string]$Ref, $Value, [int]$Style = 0, [string]$Type = 'inline')
    $p = Split-XRef $Ref
    $rowNum = $p.Row
    if (-not $Sheet.Rows.ContainsKey($rowNum)) {
        Add-XRow -Sheet $Sheet -Row $rowNum
    }
    $isEmpty = ($null -eq $Value) -or ([string]$Value -eq '')
    if ($isEmpty -and $Style -eq 0) { return }
    $Sheet.Rows[$rowNum].Cells[$p.Col] = @{ Ref = $Ref; Value = $Value; Style = $Style; Type = $Type }
}

function Get-XStyleXml {
    return @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="2">
  <numFmt numFmtId="164" formatCode="yyyy\-mm\-dd"/>
  <numFmt numFmtId="165" formatCode="0.0%"/>
</numFmts>
<fonts count="10">
  <font><sz val="11"/><color theme="1"/><name val="Microsoft YaHei"/><family val="2"/></font>
  <font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Microsoft YaHei"/><family val="2"/></font>
  <font><b/><sz val="14"/><color rgb="FF1F4E79"/><name val="Microsoft YaHei"/><family val="2"/></font>
  <font><i/><sz val="11"/><color rgb="FF808080"/><name val="Microsoft YaHei"/><family val="2"/></font>
  <font><i/><strike/><sz val="11"/><color rgb="FF808080"/><name val="Microsoft YaHei"/><family val="2"/></font>
  <font><b/><sz val="11"/><color theme="1"/><name val="Microsoft YaHei"/><family val="2"/></font>
  <font><b/><sz val="12"/><color rgb="FFFFFFFF"/><name val="Microsoft YaHei"/><family val="2"/></font>
  <font><sz val="11"/><color rgb="FF375623"/><name val="Microsoft YaHei"/><family val="2"/></font>
  <font><sz val="11"/><color rgb="FF9C0006"/><name val="Microsoft YaHei"/><family val="2"/></font>
  <font><sz val="11"/><color rgb="FF9C6500"/><name val="Microsoft YaHei"/><family val="2"/></font>
</fonts>
<fills count="7">
  <fill><patternFill patternType="none"/></fill>
  <fill><patternFill patternType="gray125"/></fill>
  <fill><patternFill patternType="solid"><fgColor rgb="FF1F4E79"/><bgColor indexed="64"/></patternFill></fill>
  <fill><patternFill patternType="solid"><fgColor rgb="FFF2F2F2"/><bgColor indexed="64"/></patternFill></fill>
  <fill><patternFill patternType="solid"><fgColor rgb="FF2E75B6"/><bgColor indexed="64"/></patternFill></fill>
  <fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/><bgColor indexed="64"/></patternFill></fill>
  <fill><patternFill patternType="solid"><fgColor rgb="FFDDEBF7"/><bgColor indexed="64"/></patternFill></fill>
</fills>
<borders count="3">
  <border><left/><right/><top/><bottom/><diagonal/></border>
  <border>
    <left style="thin"><color rgb="FFD9D9D9"/></left><right style="thin"><color rgb="FFD9D9D9"/></right>
    <top style="thin"><color rgb="FFD9D9D9"/></top><bottom style="thin"><color rgb="FFD9D9D9"/></bottom><diagonal/>
  </border>
  <border>
    <left style="thin"><color rgb="FFB4C6E7"/></left><right style="thin"><color rgb="FFB4C6E7"/></right>
    <top style="thin"><color rgb="FFB4C6E7"/></top><bottom style="thin"><color rgb="FFB4C6E7"/></bottom><diagonal/>
  </border>
</borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="28">
  <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
  <xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
  <xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
  <xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
  <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
  <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center" wrapText="1"/></xf>
  <xf numFmtId="165" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
  <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
  <xf numFmtId="0" fontId="5" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
  <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="top" wrapText="1"/></xf>
  <xf numFmtId="0" fontId="3" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
  <xf numFmtId="0" fontId="4" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
  <xf numFmtId="0" fontId="6" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
  <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="left" vertical="top" wrapText="1"/></xf>
  <xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
  <xf numFmtId="0" fontId="5" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
  <xf numFmtId="165" fontId="5" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
  <xf numFmtId="0" fontId="5" fillId="6" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
  <xf numFmtId="0" fontId="7" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
  <xf numFmtId="0" fontId="8" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
  <xf numFmtId="0" fontId="9" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
  <xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center" wrapText="1"/></xf>
  <xf numFmtId="165" fontId="2" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
  <xf numFmtId="0" fontId="5" fillId="5" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
  <xf numFmtId="0" fontId="7" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
  <xf numFmtId="0" fontId="8" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
  <xf numFmtId="0" fontId="9" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
  <xf numFmtId="0" fontId="3" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
</cellXfs>
<cellStyles count="1">
  <cellStyle name="Normal" xfId="0" builtinId="0"/>
</cellStyles>
<dxfs count="8">
  <dxf><font><b/><color rgb="FF375623"/></font><fill><patternFill><bgColor rgb="FFE2EFDA"/></patternFill></fill></dxf>
  <dxf><font><color rgb="FF9C0006"/></font><fill><patternFill><bgColor rgb="FFFCE4E4"/></patternFill></fill></dxf>
  <dxf><font><color rgb="FF808080"/></font><fill><patternFill><bgColor rgb="FFF2F2F2"/></patternFill></fill></dxf>
  <dxf><font><color rgb="FF9C6500"/></font><fill><patternFill><bgColor rgb="FFFFF2CC"/></patternFill></fill></dxf>
  <dxf><fill><patternFill><bgColor rgb="FFFFF9E3"/></patternFill></fill><border><left style="dashed"><color rgb="FFFFC000"/></left><right style="dashed"><color rgb="FFFFC000"/></right><top style="dashed"><color rgb="FFFFC000"/></top><bottom style="dashed"><color rgb="FFFFC000"/></bottom></border></dxf>
  <dxf><font><i/><color rgb="FFA6A6A6"/></font><fill><patternFill><bgColor rgb="FFF0F0F0"/></patternFill></fill></dxf>
  <dxf><font><i/><strike/><color rgb="FFA6A6A6"/></font><fill><patternFill><bgColor rgb="FFEDEDED"/></patternFill></fill></dxf>
  <dxf><fill><patternFill><bgColor rgb="FFDDDDDD"/></patternFill></fill></dxf>
</dxfs>
<tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>
</styleSheet>
'@
}

function ConvertTo-XAttr {
    param([string]$Text)
    if ($null -eq $Text) { return '' }
    return [System.Security.SecurityElement]::Escape($Text)
}

function Get-XSheetXml {
    param($Sheet, [bool]$Selected = $false)

    $rowNumbers = @($Sheet.Rows.Keys)
    $maxRow = if ($rowNumbers.Count -gt 0) { ($rowNumbers | Measure-Object -Maximum).Maximum } else { 1 }
    $maxCol = 1
    foreach ($r in $rowNumbers) {
        $cells = $Sheet.Rows[$r].Cells
        if ($cells.Count -gt 0) {
            $m = ($cells.Keys | Measure-Object -Maximum).Maximum
            if ($m -gt $maxCol) { $maxCol = $m }
        }
    }
    foreach ($merge in $Sheet.Merges) {
        foreach ($ref in $merge.Split(':')) {
            $p = Split-XRef $ref
            if ($p.Row -gt $maxRow) { $maxRow = $p.Row }
            if ($p.Col -gt $maxCol) { $maxCol = $p.Col }
        }
    }
    $dim = 'A1:' + (Get-XColName $maxCol) + $maxRow

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    [void]$sb.Append('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">')
    if ($Sheet.TabColor) {
        [void]$sb.Append('<sheetPr><tabColor rgb="' + $Sheet.TabColor + '"/></sheetPr>')
    }
    [void]$sb.Append('<dimension ref="' + $dim + '"/>')
    [void]$sb.Append('<sheetViews><sheetView workbookViewId="0"')
    if ($Selected) { [void]$sb.Append(' tabSelected="1"') }
    [void]$sb.Append('>')
    if ($Sheet.Freeze) {
        $viewCell = if ($Sheet.ViewCell) { $Sheet.ViewCell } else { $Sheet.Freeze }
        $p = Split-XRef $Sheet.Freeze
        $fl = 'I3'
        $xSplit = 8
        $ySplit = 2
        if ($Sheet.Freeze -match '^([A-Z]+)(\d+)$') {
            $xSplit = (Get-XColIndex $Matches[1]) - 1
            $ySplit = [int]$Matches[2] - 1
        }
        [void]$sb.Append('<pane xSplit="' + $xSplit + '" ySplit="' + $ySplit + '" topLeftCell="' + $viewCell + '" activePane="bottomRight" state="frozen"/>')
        [void]$sb.Append('<selection pane="bottomRight" activeCell="' + $viewCell + '" sqref="' + $viewCell + '"/>')
    }
    [void]$sb.Append('</sheetView></sheetViews>')
    [void]$sb.Append('<sheetFormatPr defaultRowHeight="' + $Sheet.DefaultRowHt + '"/>')

    if ($Sheet.Cols.Count -gt 0) {
        [void]$sb.Append('<cols>')
        foreach ($c in $Sheet.Cols) {
            [void]$sb.Append('<col min="' + $c.Min + '" max="' + $c.Max + '" width="' + $c.Width + '" customWidth="1"')
            if ($c.Hidden) { [void]$sb.Append(' hidden="1"') }
            [void]$sb.Append('/>')
        }
        [void]$sb.Append('</cols>')
    }

    [void]$sb.Append('<sheetData>')
    foreach ($rn in $rowNumbers) {
        $row = $Sheet.Rows[$rn]
        [void]$sb.Append('<row r="' + $rn + '"')
        if ($row.Height -gt 0) {
            [void]$sb.Append(' ht="' + $row.Height + '" customHeight="1"')
        }
        if ($row.Hidden) { [void]$sb.Append(' hidden="1"') }
        if ($row.Cells.Count -eq 0) {
            [void]$sb.Append('/>')
            continue
        }
        [void]$sb.Append('>')
        foreach ($ci in $row.Cells.Keys) {
            $cell = $row.Cells[$ci]
            [void]$sb.Append('<c r="' + $cell.Ref + '"')
            if ($cell.Style -ne 0) { [void]$sb.Append(' s="' + $cell.Style + '"') }
            $val = $cell.Value
            $isEmpty = ($null -eq $val) -or ([string]$val -eq '')
            if ($isEmpty) {
                [void]$sb.Append('/>')
                continue
            }
            switch ($cell.Type) {
                'n' {
                    [void]$sb.Append('><v>' + ([string]$val) + '</v></c>')
                }
                'f' {
                    [void]$sb.Append('><f>' + (ConvertTo-XAttr ([string]$val)) + '</f></c>')
                }
                default {
                    [void]$sb.Append(' t="inlineStr"><is><t xml:space="preserve">' + (ConvertTo-XAttr ([string]$val)) + '</t></is></c>')
                }
            }
        }
        [void]$sb.Append('</row>')
    }
    [void]$sb.Append('</sheetData>')

    if ($Sheet.Merges.Count -gt 0) {
        [void]$sb.Append('<mergeCells count="' + $Sheet.Merges.Count + '">')
        foreach ($m in $Sheet.Merges) { [void]$sb.Append('<mergeCell ref="' + $m + '"/>') }
        [void]$sb.Append('</mergeCells>')
    }

    foreach ($cf in $Sheet.CondFormats) {
        [void]$sb.Append('<conditionalFormatting sqref="' + $cf.Sqref + '">')
        foreach ($rule in $cf.Rules) {
            if ($rule.Type -eq 'cellIs') {
                [void]$sb.Append('<cfRule type="cellIs" dxfId="' + $rule.Dxf + '" priority="' + $rule.Priority + '" operator="' + $rule.Operator + '"')
                if ($rule.Stop) { [void]$sb.Append(' stopIfTrue="1"') }
                [void]$sb.Append('><formula>' + (ConvertTo-XAttr $rule.Formula) + '</formula></cfRule>')
            }
            else {
                [void]$sb.Append('<cfRule type="expression" dxfId="' + $rule.Dxf + '" priority="' + $rule.Priority + '"')
                if ($rule.Stop) { [void]$sb.Append(' stopIfTrue="1"') }
                [void]$sb.Append('><formula>' + (ConvertTo-XAttr $rule.Formula) + '</formula></cfRule>')
            }
        }
        [void]$sb.Append('</conditionalFormatting>')
    }

    if ($Sheet.Validations.Count -gt 0) {
        [void]$sb.Append('<dataValidations count="' + $Sheet.Validations.Count + '">')
        foreach ($v in $Sheet.Validations) {
            [void]$sb.Append('<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" errorTitle="输入无效" error="请从下拉列表中选择" sqref="' + $v.Sqref + '">')
            [void]$sb.Append('<formula1>' + (ConvertTo-XAttr $v.Formula) + '</formula1>')
            [void]$sb.Append('</dataValidation>')
        }
        [void]$sb.Append('</dataValidations>')
    }

    [void]$sb.Append('<pageMargins left="0.5" right="0.5" top="0.6" bottom="0.6" header="0.3" footer="0.3"/>')
    [void]$sb.Append('</worksheet>')
    return $sb.ToString()
}

function Write-ZipText {
    param($Zip, [string]$EntryName, [string]$Text)
    $entry = $Zip.CreateEntry($EntryName, [System.IO.Compression.CompressionLevel]::Optimal)
    $stream = $entry.Open()
    $writer = New-Object System.IO.StreamWriter($stream, (New-Object System.Text.UTF8Encoding($false)))
    try {
        $writer.Write($Text)
    }
    finally {
        $writer.Dispose()
        $stream.Dispose()
    }
}

function New-XlsxFile {
    param(
        [string]$Path,
        $Sheets,
        [string]$Title = '每日打卡',
        [string]$Author = 'DailyCheckin'
    )

    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    if (Test-Path -LiteralPath $Path) {
        Remove-Item -LiteralPath $Path -Force
    }
    $zip = [System.IO.Compression.ZipFile]::Open($Path, [System.IO.Compression.ZipArchiveMode]::Create)
    try {
        $contentTypes = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
__SHEET_TYPES__
</Types>
'@
        $sheetTypes = ''
        for ($i = 0; $i -lt $Sheets.Count; $i++) {
            $n = $i + 1
            $sheetTypes += '<Override PartName="/xl/worksheets/sheet' + $n + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        }
        $contentTypes = $contentTypes.Replace('__SHEET_TYPES__', $sheetTypes)
        Write-ZipText $zip '[Content_Types].xml' $contentTypes

        $rootRels = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>
'@
        Write-ZipText $zip '_rels/.rels' $rootRels

        $now = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
        $core = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<dc:title>$(ConvertTo-XAttr $Title)</dc:title>
<dc:creator>$(ConvertTo-XAttr $Author)</dc:creator>
<cp:lastModifiedBy>$(ConvertTo-XAttr $Author)</cp:lastModifiedBy>
<dcterms:created xsi:type="dcterms:W3CDTF">$now</dcterms:created>
<dcterms:modified xsi:type="dcterms:W3CDTF">$now</dcterms:modified>
</cp:coreProperties>
"@
        Write-ZipText $zip 'docProps/core.xml' $core

        $app = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
<Application>Microsoft Excel</Application>
<DocSecurity>0</DocSecurity>
<ScaleCrop>false</ScaleCrop>
<LinksUpToDate>false</LinksUpToDate>
<SharedDoc>false</SharedDoc>
<HyperlinksChanged>false</HyperlinksChanged>
<AppVersion>16.0300</AppVersion>
</Properties>
'@
        Write-ZipText $zip 'docProps/app.xml' $app

        $wbSheets = ''
        $wbRels = ''
        for ($i = 0; $i -lt $Sheets.Count; $i++) {
            $n = $i + 1
            $wbSheets += '<sheet name="' + (ConvertTo-XAttr $Sheets[$i].Name) + '" sheetId="' + $n + '" r:id="rId' + $n + '"/>'
            $wbRels += '<Relationship Id="rId' + $n + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + $n + '.xml"/>'
        }
        $styleRel = $Sheets.Count + 1
        $wbRels += '<Relationship Id="rId' + $styleRel + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
        $workbook = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<bookViews><workbookView xWindow="0" yWindow="0" windowWidth="28800" windowHeight="16000"/></bookViews>
<sheets>$wbSheets</sheets>
<calcPr calcId="191029"/>
</workbook>
"@
        Write-ZipText $zip 'xl/workbook.xml' $workbook

        $rels = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
$wbRels
</Relationships>
"@
        Write-ZipText $zip 'xl/_rels/workbook.xml.rels' $rels
        Write-ZipText $zip 'xl/styles.xml' (Get-XStyleXml)

        for ($i = 0; $i -lt $Sheets.Count; $i++) {
            $n = $i + 1
            $selected = ($i -eq 0)
            Write-ZipText $zip ('xl/worksheets/sheet' + $n + '.xml') (Get-XSheetXml -Sheet $Sheets[$i] -Selected $selected)
        }
    }
    finally {
        $zip.Dispose()
    }
}

function Read-XlsxFile {
    param([string]$Path)
    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
        $relNs = 'http://schemas.openxmlformats.org/package/2006/relationships'

        $shared = @()
        $ssEntry = $zip.GetEntry('xl/sharedStrings.xml')
        if ($ssEntry) {
            $doc = New-Object System.Xml.XmlDocument
            $sr = New-Object System.IO.StreamReader($ssEntry.Open(), [System.Text.Encoding]::UTF8)
            try { $doc.LoadXml($sr.ReadToEnd()) } finally { $sr.Dispose() }
            foreach ($si in $doc.DocumentElement.ChildNodes) {
                if ($si.LocalName -ne 'si') { continue }
                $texts = @()
                foreach ($t in $si.GetElementsByTagName('t', $ns)) { $texts += $t.InnerText }
                $shared += ($texts -join '')
            }
        }

        $wbEntry = $zip.GetEntry('xl/workbook.xml')
        if (-not $wbEntry) { throw '文件不是有效的 xlsx：缺少 workbook.xml' }
        $wbDoc = New-Object System.Xml.XmlDocument
        $wbReader = New-Object System.IO.StreamReader($wbEntry.Open(), [System.Text.Encoding]::UTF8)
        try { $wbDoc.LoadXml($wbReader.ReadToEnd()) } finally { $wbReader.Dispose() }

        $relEntry = $zip.GetEntry('xl/_rels/workbook.xml.rels')
        $relMap = @{}
        if ($relEntry) {
            $relDoc = New-Object System.Xml.XmlDocument
            $relReader = New-Object System.IO.StreamReader($relEntry.Open(), [System.Text.Encoding]::UTF8)
            try { $relDoc.LoadXml($relReader.ReadToEnd()) } finally { $relReader.Dispose() }
            foreach ($rel in $relDoc.DocumentElement.ChildNodes) {
                if ($rel.LocalName -ne 'Relationship') { continue }
                $relMap[$rel.GetAttribute('Id')] = $rel.GetAttribute('Target')
            }
        }

        $result = @{}
        $order = @()
        foreach ($sh in $wbDoc.GetElementsByTagName('sheet', $ns)) {
            $name = $sh.GetAttribute('name')
            $rid = $sh.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
            $target = $relMap[$rid]
            if (-not $target) { continue }
            if ($target.StartsWith('/')) { $entryPath = $target.TrimStart('/') }
            else { $entryPath = 'xl/' + $target.TrimStart('/') }
            $entry = $zip.GetEntry($entryPath)
            if (-not $entry) { continue }

            $doc = New-Object System.Xml.XmlDocument
            $reader = New-Object System.IO.StreamReader($entry.Open(), [System.Text.Encoding]::UTF8)
            try { $doc.LoadXml($reader.ReadToEnd()) } finally { $reader.Dispose() }

            $cells = @{}
            $maxRow = 0
            $maxCol = 0
            foreach ($c in $doc.GetElementsByTagName('c', $ns)) {
                $ref = $c.GetAttribute('r')
                if (-not $ref) { continue }
                $p = Split-XRef $ref
                if ($p.Row -gt $maxRow) { $maxRow = $p.Row }
                if ($p.Col -gt $maxCol) { $maxCol = $p.Col }
                $t = $c.GetAttribute('t')
                $val = $null
                if ($t -eq 'inlineStr') {
                    $is = $c.GetElementsByTagName('is', $ns) | Select-Object -First 1
                    if ($is) {
                        $texts = @()
                        foreach ($tx in $is.GetElementsByTagName('t', $ns)) { $texts += $tx.InnerText }
                        $val = ($texts -join '')
                    }
                }
                elseif ($t -eq 's') {
                    $vv = $c.GetElementsByTagName('v', $ns) | Select-Object -First 1
                    if ($vv) {
                        $idx = [int]$vv.InnerText
                        if ($idx -ge 0 -and $idx -lt $shared.Count) { $val = $shared[$idx] }
                    }
                }
                elseif ($t -eq 'b') {
                    $vv = $c.GetElementsByTagName('v', $ns) | Select-Object -First 1
                    if ($vv) { $val = if ($vv.InnerText -eq '1') { 'TRUE' } else { 'FALSE' } }
                }
                else {
                    $vv = $c.GetElementsByTagName('v', $ns) | Select-Object -First 1
                    if ($vv) { $val = $vv.InnerText }
                }
                $cells[$ref] = @{ Value = $val; Row = $p.Row; Col = $p.Col; Type = $t }
            }
            $result[$name] = @{ Name = $name; Cells = $cells; MaxRow = $maxRow; MaxCol = $maxCol }
            $order += $name
        }
        return @{ Sheets = $result; Order = $order }
    }
    finally {
        $zip.Dispose()
    }
}
