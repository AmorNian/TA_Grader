param(
    [string]$Assignment = "Daily1",
    [int]$Port = 8765,
    [switch]$SelfTest,
    [switch]$NoOpen
)

$ErrorActionPreference = "Stop"

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$HomeworksDir = Join-Path $Root "homeworks"
$WorkbookPath = Join-Path $Root "grade.xlsx"
if (-not (Test-Path $WorkbookPath)) {
    $alternate = Join-Path $Root "homeworks.xlsx"
    if (Test-Path $alternate) {
        $WorkbookPath = $alternate
    }
}

if (-not (Test-Path $HomeworksDir)) {
    throw "Cannot find homeworks folder: $HomeworksDir"
}
if (-not (Test-Path $WorkbookPath)) {
    throw "Cannot find homework.xlsx or homeworks.xlsx in: $Root"
}

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$script:SupportedFileTypes = @(".pdf", ".jpg", ".jpeg", ".png")
$script:ContentTypes = @{
    ".pdf" = "application/pdf"
    ".jpg" = "image/jpeg"
    ".jpeg" = "image/jpeg"
    ".png" = "image/png"
}

function Get-SubmissionFile {
    param([string]$FolderPath)
    $files = Get-ChildItem -Path $FolderPath -File -Recurse |
        Where-Object { $script:SupportedFileTypes -contains $_.Extension.ToLowerInvariant() } |
        Sort-Object @{ Expression = {
            $ext = $_.Extension.ToLowerInvariant()
            $idx = [array]::IndexOf($script:SupportedFileTypes, $ext)
            if ($idx -lt 0) { 999 } else { $idx }
        } }, FullName

    return $files | Select-Object -First 1
}

function Update-Submissions {
    param([switch]$PreserveOnEmpty)

    $nextSubmissions = @()
    $nextIdToIndex = @{}

    $folders = Get-ChildItem -Path $HomeworksDir -Directory | Sort-Object Name
    foreach ($folder in $folders) {
        $id = ($folder.Name -split "_")[0]
        $file = Get-SubmissionFile $folder.FullName
        $item = [ordered]@{
            id = $id
            folder = $folder.Name
            fileName = if ($file) { $file.Name } else { "" }
            hasFile = [bool]$file
            filePath = if ($file) { $file.FullName } else { "" }
            fileType = if ($file) { $file.Extension.TrimStart(".").ToLowerInvariant() } else { "" }
            pdfName = if ($file) { $file.Name } else { "" }
            hasPdf = [bool]$file
            pdfPath = if ($file) { $file.FullName } else { "" }
        }
        $nextIdToIndex[$id] = $nextSubmissions.Count
        $nextSubmissions += [pscustomobject]$item
    }

    if ($PreserveOnEmpty -and $nextSubmissions.Count -eq 0 -and $script:Submissions.Count -gt 0) {
        return
    }

    $script:Submissions = $nextSubmissions
    $script:IdToIndex = $nextIdToIndex
}

Update-Submissions

function New-JsonResponse {
    param(
        [object]$Data,
        [int]$Depth = 8
    )
    return ($Data | ConvertTo-Json -Depth $Depth -Compress)
}

function Write-Response {
    param(
        [System.IO.Stream]$Stream,
        [string]$Body,
        [string]$ContentType = "application/json; charset=utf-8",
        [int]$StatusCode = 200
    )
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($Body)
    $reason = if ($StatusCode -eq 200) { "OK" } elseif ($StatusCode -eq 400) { "Bad Request" } elseif ($StatusCode -eq 404) { "Not Found" } else { "Internal Server Error" }
    $header = "HTTP/1.1 $StatusCode $reason`r`nContent-Type: $ContentType`r`nContent-Length: $($bytes.Length)`r`nConnection: close`r`n`r`n"
    $headerBytes = [System.Text.Encoding]::ASCII.GetBytes($header)
    try {
        $Stream.Write($headerBytes, 0, $headerBytes.Length)
        $Stream.Write($bytes, 0, $bytes.Length)
    } catch [System.IO.IOException] {
        return
    } catch [System.ObjectDisposedException] {
        return
    }
}

function Write-FileResponse {
    param(
        [System.IO.Stream]$Stream,
        [string]$Path,
        [string]$ContentType
    )
    if (-not (Test-Path $Path)) {
        Write-Response $Stream (New-JsonResponse @{ ok = $false; error = "File not found" }) "application/json; charset=utf-8" 404
        return
    }
    $bytes = [System.IO.File]::ReadAllBytes($Path)
    $header = "HTTP/1.1 200 OK`r`nContent-Type: $ContentType`r`nContent-Length: $($bytes.Length)`r`nConnection: close`r`n`r`n"
    $headerBytes = [System.Text.Encoding]::ASCII.GetBytes($header)
    try {
        $Stream.Write($headerBytes, 0, $headerBytes.Length)
        $Stream.Write($bytes, 0, $bytes.Length)
    } catch [System.IO.IOException] {
        return
    } catch [System.ObjectDisposedException] {
        return
    }
}

function ConvertTo-ColName {
    param([int]$Column)
    $name = ""
    while ($Column -gt 0) {
        $Column--
        $name = [char](65 + ($Column % 26)) + $name
        $Column = [Math]::Floor($Column / 26)
    }
    return $name
}

function ConvertFrom-ColName {
    param([string]$Name)
    $n = 0
    foreach ($ch in $Name.ToUpperInvariant().ToCharArray()) {
        if ($ch -lt 'A' -or $ch -gt 'Z') { continue }
        $n = ($n * 26) + ([int][char]$ch - [int][char]'A' + 1)
    }
    return $n
}

function Get-CellRef {
    param([int]$Row, [int]$Column)
    return "$(ConvertTo-ColName $Column)$Row"
}

function Split-CellRef {
    param([string]$Ref)
    if ($Ref -match "^([A-Z]+)(\d+)$") {
        return @{ col = ConvertFrom-ColName $matches[1]; row = [int]$matches[2] }
    }
    return $null
}

function Get-ZipText {
    param(
        [System.IO.Compression.ZipArchive]$Zip,
        [string]$EntryName
    )
    $entry = $Zip.GetEntry($EntryName)
    if (-not $entry) { return $null }
    $reader = New-Object System.IO.StreamReader($entry.Open(), [System.Text.Encoding]::UTF8)
    try { return $reader.ReadToEnd() }
    finally { $reader.Close() }
}

function Set-ZipText {
    param(
        [System.IO.Compression.ZipArchive]$Zip,
        [string]$EntryName,
        [string]$Text
    )
    $old = $Zip.GetEntry($EntryName)
    if ($old) { $old.Delete() }
    $entry = $Zip.CreateEntry($EntryName, [System.IO.Compression.CompressionLevel]::Optimal)
    $writer = New-Object System.IO.StreamWriter($entry.Open(), (New-Object System.Text.UTF8Encoding($false)))
    try { $writer.Write($Text) }
    finally { $writer.Close() }
}

function Open-WorkbookZipRead {
    $stream = [System.IO.File]::Open($WorkbookPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
    try {
        $zip = [System.IO.Compression.ZipArchive]::new($stream, [System.IO.Compression.ZipArchiveMode]::Read)
        return @{ stream = $stream; zip = $zip }
    } catch {
        $stream.Dispose()
        throw
    }
}

function Copy-WorkbookToTemp {
    param([string]$Destination)
    $source = [System.IO.File]::Open($WorkbookPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
    try {
        $target = [System.IO.File]::Open($Destination, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
        try {
            $source.CopyTo($target)
        } finally {
            $target.Dispose()
        }
    } finally {
        $source.Dispose()
    }
}

function ConvertTo-XmlDoc {
    param([string]$Text)
    $doc = New-Object System.Xml.XmlDocument
    $doc.PreserveWhitespace = $true
    $doc.LoadXml($Text)
    return $doc
}

function Get-FirstSheetInfo {
    param([System.IO.Compression.ZipArchive]$Zip)
    $workbook = ConvertTo-XmlDoc (Get-ZipText $Zip "xl/workbook.xml")
    $rels = ConvertTo-XmlDoc (Get-ZipText $Zip "xl/_rels/workbook.xml.rels")
    $ns = New-Object System.Xml.XmlNamespaceManager($workbook.NameTable)
    $ns.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
    $ns.AddNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
    $sheet = $workbook.SelectSingleNode("//x:sheets/x:sheet[1]", $ns)
    if (-not $sheet) { throw "No worksheet found in workbook." }
    $rid = $sheet.GetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")

    $relNs = New-Object System.Xml.XmlNamespaceManager($rels.NameTable)
    $relNs.AddNamespace("rel", "http://schemas.openxmlformats.org/package/2006/relationships")
    $rel = $rels.SelectSingleNode("//rel:Relationship[@Id='$rid']", $relNs)
    if (-not $rel) { throw "Cannot find worksheet relationship: $rid" }

    $target = $rel.GetAttribute("Target").Replace("\", "/")
    if ($target.StartsWith("/")) {
        $entry = $target.TrimStart("/")
    } else {
        $entry = "xl/" + $target.TrimStart("/")
    }
    return @{ name = $sheet.GetAttribute("name"); entry = $entry }
}

function Get-SharedStrings {
    param([System.IO.Compression.ZipArchive]$Zip)
    $text = Get-ZipText $Zip "xl/sharedStrings.xml"
    if (-not $text) { return @() }
    $doc = ConvertTo-XmlDoc $text
    $ns = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
    $ns.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
    $items = @()
    foreach ($si in $doc.SelectNodes("//x:sst/x:si", $ns)) {
        $parts = @()
        foreach ($t in $si.SelectNodes(".//x:t", $ns)) {
            $parts += $t.InnerText
        }
        $items += ($parts -join "")
    }
    return $items
}

function Get-SheetContext {
    param([System.IO.Compression.ZipArchive]$Zip)
    $sheetInfo = Get-FirstSheetInfo $Zip
    $sheetDoc = ConvertTo-XmlDoc (Get-ZipText $Zip $sheetInfo.entry)
    $ns = New-Object System.Xml.XmlNamespaceManager($sheetDoc.NameTable)
    $ns.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
    return @{
        sheetName = $sheetInfo.name
        sheetEntry = $sheetInfo.entry
        sheetDoc = $sheetDoc
        ns = $ns
        sharedStrings = Get-SharedStrings $Zip
    }
}

function Get-SheetUsedRange {
    param([System.Xml.XmlDocument]$Doc, [System.Xml.XmlNamespaceManager]$Ns)
    $maxRow = 1
    $maxCol = 1
    foreach ($cell in $Doc.SelectNodes("//x:sheetData/x:row/x:c", $Ns)) {
        $ref = Split-CellRef $cell.GetAttribute("r")
        if ($ref) {
            if ($ref.row -gt $maxRow) { $maxRow = $ref.row }
            if ($ref.col -gt $maxCol) { $maxCol = $ref.col }
        }
    }
    return @{ rows = $maxRow; cols = $maxCol }
}

function Get-CellNode {
    param(
        [System.Xml.XmlDocument]$Doc,
        [System.Xml.XmlNamespaceManager]$Ns,
        [int]$Row,
        [int]$Column,
        [bool]$Create
    )
    $ref = Get-CellRef $Row $Column
    $cell = $Doc.SelectSingleNode("//x:sheetData/x:row[@r='$Row']/x:c[@r='$ref']", $Ns)
    if ($cell -or -not $Create) { return $cell }

    $sheetData = $Doc.SelectSingleNode("//x:sheetData", $Ns)
    if (-not $sheetData) {
        $sheetData = $Doc.CreateElement("sheetData", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
        $Doc.DocumentElement.AppendChild($sheetData) | Out-Null
    }
    $rowNode = $Doc.SelectSingleNode("//x:sheetData/x:row[@r='$Row']", $Ns)
    if (-not $rowNode) {
        $rowNode = $Doc.CreateElement("row", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
        $rowNode.SetAttribute("r", [string]$Row)
        $sheetData.AppendChild($rowNode) | Out-Null
    }
    $cell = $Doc.CreateElement("c", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
    $cell.SetAttribute("r", $ref)
    $rowNode.AppendChild($cell) | Out-Null
    return $cell
}

function Get-CellText {
    param(
        [System.Xml.XmlDocument]$Doc,
        [System.Xml.XmlNamespaceManager]$Ns,
        [array]$SharedStrings,
        [int]$Row,
        [int]$Column
    )
    $cell = Get-CellNode $Doc $Ns $Row $Column $false
    if (-not $cell) { return "" }
    $type = $cell.GetAttribute("t")
    if ($type -eq "s") {
        $v = $cell.SelectSingleNode("x:v", $Ns)
        if ($v -and $v.InnerText -match "^\d+$") {
            $idx = [int]$v.InnerText
            if ($idx -ge 0 -and $idx -lt $SharedStrings.Count) { return [string]$SharedStrings[$idx] }
        }
        return ""
    }
    if ($type -eq "inlineStr") {
        $parts = @()
        foreach ($t in $cell.SelectNodes(".//x:t", $Ns)) { $parts += $t.InnerText }
        return ($parts -join "")
    }
    $valueNode = $cell.SelectSingleNode("x:v", $Ns)
    if ($valueNode) { return [string]$valueNode.InnerText }
    return ""
}

function Set-CellText {
    param(
        [System.Xml.XmlDocument]$Doc,
        [System.Xml.XmlNamespaceManager]$Ns,
        [int]$Row,
        [int]$Column,
        [string]$Value
    )
    $cell = Get-CellNode $Doc $Ns $Row $Column $true
    while ($cell.FirstChild) { $cell.RemoveChild($cell.FirstChild) | Out-Null }
    $cell.RemoveAttribute("t")

    if ($Value.Trim() -match "^-?\d+(\.\d+)?$") {
        $v = $Doc.CreateElement("v", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
        $v.InnerText = $Value.Trim()
        $cell.AppendChild($v) | Out-Null
    } else {
        $cell.SetAttribute("t", "inlineStr")
        $is = $Doc.CreateElement("is", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
        $t = $Doc.CreateElement("t", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
        $t.InnerText = $Value.Trim()
        $is.AppendChild($t) | Out-Null
        $cell.AppendChild($is) | Out-Null
    }
}

function Find-HeaderCell {
    param(
        [hashtable]$Ctx,
        [string]$Header,
        [bool]$CreateIfMissing
    )
    $used = Get-SheetUsedRange $Ctx.sheetDoc $Ctx.ns
    $rows = $used.rows
    $cols = $used.cols
    $maxHeaderRows = [Math]::Min(20, $rows)

    for ($r = 1; $r -le $maxHeaderRows; $r++) {
        for ($c = 1; $c -le $cols; $c++) {
            $value = Get-CellText $Ctx.sheetDoc $Ctx.ns $Ctx.sharedStrings $r $c
            if ($value.Trim().ToLowerInvariant() -eq $Header.Trim().ToLowerInvariant()) {
                return @{ row = $r; col = $c; created = $false }
            }
        }
    }

    if ($CreateIfMissing) {
        $newCol = $cols + 1
        Set-CellText $Ctx.sheetDoc $Ctx.ns 1 $newCol $Header
        return @{ row = 1; col = $newCol; created = $true }
    }

    throw "Column '$Header' was not found."
}

function Find-StudentRow {
    param(
        [hashtable]$Ctx,
        [string]$StudentId
    )
    $used = Get-SheetUsedRange $Ctx.sheetDoc $Ctx.ns
    $rows = $used.rows
    $cols = $used.cols

    for ($r = 1; $r -le $rows; $r++) {
        for ($c = 1; $c -le $cols; $c++) {
            $value = Get-CellText $Ctx.sheetDoc $Ctx.ns $Ctx.sharedStrings $r $c
            if ($value.Trim().ToUpperInvariant() -eq $StudentId.Trim().ToUpperInvariant()) {
                return $r
            }
        }
    }
    throw "Student id '$StudentId' was not found in workbook."
}

function Find-HeaderColumnByNames {
    param(
        [hashtable]$Ctx,
        [string[]]$Names
    )
    $used = Get-SheetUsedRange $Ctx.sheetDoc $Ctx.ns
    $maxHeaderRows = [Math]::Min(20, $used.rows)
    $nameSet = @{}
    foreach ($name in $Names) {
        $nameSet[$name.Trim().ToLowerInvariant()] = $true
    }

    for ($r = 1; $r -le $maxHeaderRows; $r++) {
        for ($c = 1; $c -le $used.cols; $c++) {
            $value = (Get-CellText $Ctx.sheetDoc $Ctx.ns $Ctx.sharedStrings $r $c).Trim().ToLowerInvariant()
            if ($nameSet.ContainsKey($value)) {
                return @{ row = $r; col = $c }
            }
        }
    }
    return $null
}

function Get-StudentInfoMap {
    $opened = Open-WorkbookZipRead
    $zip = $opened.zip
    try {
        $ctx = Get-SheetContext $zip
        $used = Get-SheetUsedRange $ctx.sheetDoc $ctx.ns
        $idHeader = Find-HeaderColumnByNames $ctx @("学籍番号", "student id", "studentid", "id")
        $nameHeader = Find-HeaderColumnByNames $ctx @("氏名", "name", "名前", "student name", "studentname")
        if (-not $idHeader -or -not $nameHeader) {
            return @{}
        }

        $result = @{}
        $startRow = [Math]::Max($idHeader.row, $nameHeader.row) + 1
        for ($r = $startRow; $r -le $used.rows; $r++) {
            $id = (Get-CellText $ctx.sheetDoc $ctx.ns $ctx.sharedStrings $r $idHeader.col).Trim()
            if (-not $id) { continue }
            $name = (Get-CellText $ctx.sheetDoc $ctx.ns $ctx.sharedStrings $r $nameHeader.col).Trim()
            $result[$id.ToUpperInvariant()] = @{
                id = $id
                name = $name
                row = $r
            }
        }
        return $result
    } finally {
        $zip.Dispose()
        $opened.stream.Dispose()
    }
}

function Get-WorkbookInfo {
    $opened = Open-WorkbookZipRead
    $zip = $opened.zip
    try {
        $ctx = Get-SheetContext $zip
        $used = Get-SheetUsedRange $ctx.sheetDoc $ctx.ns
        $cols = $used.cols
        $headers = @()
        for ($c = 1; $c -le $cols; $c++) {
            $text = (Get-CellText $ctx.sheetDoc $ctx.ns $ctx.sharedStrings 1 $c).Trim()
            if ($text) { $headers += $text }
        }
        return @{
            workbook = [System.IO.Path]::GetFileName($WorkbookPath)
            sheet = $ctx.sheetName
            headers = $headers
            lockedHint = Test-Path (Join-Path (Split-Path $WorkbookPath -Parent) ("~$" + [System.IO.Path]::GetFileName($WorkbookPath)))
        }
    } finally {
        $zip.Dispose()
        $opened.stream.Dispose()
    }
}

function Get-Score {
    param(
        [string]$StudentId,
        [string]$Header
    )
    $opened = Open-WorkbookZipRead
    $zip = $opened.zip
    try {
        $ctx = Get-SheetContext $zip
        $headerInfo = Find-HeaderCell $ctx $Header $false
        $row = Find-StudentRow $ctx $StudentId
        $text = Get-CellText $ctx.sheetDoc $ctx.ns $ctx.sharedStrings $row $headerInfo.col
        return @{ score = $text; row = $row; col = $headerInfo.col }
    } finally {
        $zip.Dispose()
        $opened.stream.Dispose()
    }
}

function Get-GradedMap {
    param(
        [string]$Header
    )
    $opened = Open-WorkbookZipRead
    $zip = $opened.zip
    try {
        $ctx = Get-SheetContext $zip
        $headerInfo = Find-HeaderCell $ctx $Header $false
        $used = Get-SheetUsedRange $ctx.sheetDoc $ctx.ns
        $idSet = @{}
        foreach ($submission in $script:Submissions) {
            $idSet[$submission.id.Trim().ToUpperInvariant()] = $submission.id
        }
        $rowToId = @{}
        for ($r = 1; $r -le $used.rows; $r++) {
            for ($c = 1; $c -le $used.cols; $c++) {
                $value = (Get-CellText $ctx.sheetDoc $ctx.ns $ctx.sharedStrings $r $c).Trim().ToUpperInvariant()
                if ($idSet.ContainsKey($value) -and -not $rowToId.ContainsKey($r)) {
                    $rowToId[$r] = $idSet[$value]
                    break
                }
            }
        }

        $result = @{}
        foreach ($rowKey in $rowToId.Keys) {
            $score = (Get-CellText $ctx.sheetDoc $ctx.ns $ctx.sharedStrings ([int]$rowKey) $headerInfo.col).Trim()
            if ($score) {
                $result[$rowToId[$rowKey]] = $score
            }
        }
        return $result
    } finally {
        $zip.Dispose()
        $opened.stream.Dispose()
    }
}

function Save-Score {
    param(
        [string]$StudentId,
        [string]$Header,
        [string]$Score,
        [bool]$CreateColumn
    )
    $temp = Join-Path ([System.IO.Path]::GetTempPath()) ("grader_" + [guid]::NewGuid().ToString("N") + ".xlsx")
    Copy-WorkbookToTemp $temp
    $zip = [System.IO.Compression.ZipFile]::Open($temp, [System.IO.Compression.ZipArchiveMode]::Update)
    try {
        $ctx = Get-SheetContext $zip
        $headerInfo = Find-HeaderCell $ctx $Header $CreateColumn
        $row = Find-StudentRow $ctx $StudentId
        Set-CellText $ctx.sheetDoc $ctx.ns $row $headerInfo.col $Score
        Set-ZipText $zip $ctx.sheetEntry $ctx.sheetDoc.OuterXml
        $result = @{ row = $row; col = $headerInfo.col; createdColumn = $headerInfo.created }
    } finally {
        $zip.Dispose()
    }

    try {
        Copy-Item -LiteralPath $temp -Destination $WorkbookPath -Force
    } finally {
        Remove-Item -LiteralPath $temp -Force -ErrorAction SilentlyContinue
    }
    return $result
}

function Get-RequestBody {
    param([System.Net.HttpListenerRequest]$Request)
    $reader = New-Object System.IO.StreamReader($Request.InputStream, $Request.ContentEncoding)
    try { return $reader.ReadToEnd() }
    finally { $reader.Close() }
}

function ConvertFrom-QueryString {
    param([string]$Query)
    $result = @{}
    if (-not $Query) { return $result }
    foreach ($part in $Query.TrimStart("?").Split("&")) {
        if (-not $part) { continue }
        $kv = $part.Split("=", 2)
        $key = [System.Uri]::UnescapeDataString($kv[0].Replace("+", " "))
        $value = if ($kv.Count -gt 1) { [System.Uri]::UnescapeDataString($kv[1].Replace("+", " ")) } else { "" }
        $result[$key] = $value
    }
    return $result
}

if ($SelfTest) {
    $info = Get-WorkbookInfo
    Update-Submissions
    "Workbook: $($info.workbook)"
    "Sheet: $($info.sheet)"
    "Submissions: $($script:Submissions.Count)"
    "First submission: $($script:Submissions[0].id) / $($script:Submissions[0].fileName)"
    $studentInfo = Get-StudentInfoMap
    $firstKey = $script:Submissions[0].id.Trim().ToUpperInvariant()
    if ($studentInfo.ContainsKey($firstKey)) {
        "First student name: $($studentInfo[$firstKey].name)"
    }
    try {
        $graded = Get-GradedMap $Assignment
        "Graded in '$Assignment': $($graded.Count)"
    } catch {
        "Graded in '$Assignment': 0 ($($_.Exception.Message))"
    }
    "Headers: $($info.headers -join ', ')"
    if ($info.lockedHint) {
        "Warning: Excel lock file exists. Close the workbook in Excel before saving scores."
    }
    exit 0
}

$Html = @'
<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Homework Grader</title>
  <style>
    * { box-sizing: border-box; }
    body { margin: 0; font-family: "Microsoft YaHei", "Segoe UI", Arial, sans-serif; color: #1f2937; background: #f3f4f6; }
    .app { height: 100vh; display: grid; grid-template-columns: 280px minmax(420px, 1fr) 320px; grid-template-rows: 56px 1fr; }
    header { grid-column: 1 / 4; display: flex; align-items: center; gap: 12px; padding: 10px 14px; background: #ffffff; border-bottom: 1px solid #d1d5db; }
    header strong { font-size: 16px; white-space: nowrap; }
    label { font-size: 12px; color: #4b5563; display: block; margin-bottom: 5px; }
    input, button, select { font: inherit; }
    input[type="text"] { width: 100%; height: 34px; border: 1px solid #cbd5e1; border-radius: 6px; padding: 6px 9px; background: #fff; }
    button { height: 34px; border: 1px solid #b6c3d1; border-radius: 6px; background: #ffffff; padding: 0 12px; cursor: pointer; }
    button.primary { background: #2563eb; color: #fff; border-color: #2563eb; }
    button:disabled { opacity: .45; cursor: default; }
    .field { min-width: 160px; }
    .spacer { flex: 1; }
    .status { color: #4b5563; font-size: 13px; min-width: 240px; text-align: right; }
    aside { overflow: hidden; background: #ffffff; border-right: 1px solid #d1d5db; display: flex; flex-direction: column; }
    .search { padding: 12px; border-bottom: 1px solid #e5e7eb; }
    .list { overflow: auto; }
    .row { width: 100%; min-height: 52px; text-align: left; border: 0; border-bottom: 1px solid #edf0f3; border-radius: 0; background: #fff; padding: 8px 12px; display: block; }
    .row.active { background: #e8f0ff; box-shadow: inset 3px 0 #2563eb; }
    .row.graded { background: #f0fdf4; }
    .row.graded.active { background: #dff7e8; box-shadow: inset 3px 0 #16a34a; }
    .sidline { display: flex; align-items: center; gap: 8px; min-width: 0; }
    .sid { font-weight: 700; display: block; flex: 1; min-width: 0; }
    .studentname { color: #374151; font-size: 13px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
    .badge { flex: 0 0 auto; border-radius: 999px; padding: 2px 7px; color: #166534; background: #dcfce7; border: 1px solid #86efac; font-size: 12px; line-height: 1.3; }
    .row.active .badge { background: #bbf7d0; }
    .filename { color: #6b7280; font-size: 12px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
    main { background: #9ca3af; min-width: 0; }
    iframe { width: 100%; height: 100%; border: 0; background: #52525b; display: block; }
    .panel { background: #ffffff; border-left: 1px solid #d1d5db; padding: 16px; overflow: auto; }
    .current { border-bottom: 1px solid #e5e7eb; padding-bottom: 14px; margin-bottom: 16px; }
    .current h2 { margin: 0 0 6px; font-size: 22px; letter-spacing: 0; }
    .muted { color: #6b7280; font-size: 13px; word-break: break-all; }
    .scorebox { margin: 12px 0; }
    .scorebox input { height: 44px; font-size: 22px; font-weight: 700; }
    .actions { display: grid; grid-template-columns: 1fr 1fr; gap: 8px; margin-top: 10px; }
    .actions .wide { grid-column: 1 / 3; }
    .hint { color: #6b7280; font-size: 12px; line-height: 1.55; margin-top: 14px; }
    .warning { color: #b45309; font-size: 13px; margin-top: 10px; display: none; }
    @media (max-width: 980px) {
      .app { grid-template-columns: 220px 1fr; grid-template-rows: 56px minmax(360px, 1fr) 250px; }
      header { grid-column: 1 / 3; }
      .panel { grid-column: 1 / 3; border-left: 0; border-top: 1px solid #d1d5db; }
    }
  </style>
</head>
<body>
  <div class="app">
    <header>
      <strong>作业评分</strong>
      <div class="field">
        <label for="assignment">成绩列名</label>
        <input id="assignment" type="text" value="__ASSIGNMENT__">
      </div>
      <button id="reload">刷新</button>
      <div class="spacer"></div>
      <div id="status" class="status">正在载入...</div>
    </header>
    <aside>
      <div class="search">
        <label for="filter">搜索学号或姓名</label>
        <input id="filter" type="text" placeholder="例如 23B10064">
      </div>
      <div id="list" class="list"></div>
    </aside>
    <main>
      <iframe id="viewer" title="作业文件"></iframe>
    </main>
    <section class="panel">
      <div class="current">
        <h2 id="sid">-</h2>
        <div id="studentname" class="studentname"></div>
        <div id="filename" class="muted"></div>
        <div id="locked" class="warning">检测到 Excel 临时锁文件。保存失败时，请先关闭 Excel 里的表格。</div>
      </div>
      <label for="score">分数</label>
      <div class="scorebox"><input id="score" type="text" autocomplete="off"></div>
      <label><input id="createColumn" type="checkbox" checked> 如果列名不存在，自动新建这一列</label>
      <div class="actions">
        <button id="prev">上一个</button>
        <button id="next">下一个</button>
        <button id="save" class="primary wide">保存并下一个</button>
      </div>
      <div class="hint">
        Enter：保存并下一个。<br>
        Ctrl+Enter：只保存当前。<br>
        成绩会写入表格中包含该学号的行，以及上方填写的成绩列。
      </div>
    </section>
  </div>
  <script>
    let submissions = [];
    let filtered = [];
    let index = 0;
    const el = (id) => document.getElementById(id);

    function setStatus(text, isError = false) {
      el('status').textContent = text;
      el('status').style.color = isError ? '#b91c1c' : '#4b5563';
    }

    async function api(path, options) {
      const res = await fetch(path, options);
      const data = await res.json();
      if (!res.ok || data.ok === false) throw new Error(data.error || '请求失败');
      return data;
    }

    function escapeHtml(value) {
      return String(value || '').replace(/[&<>"']/g, ch => ({
        '&': '&amp;',
        '<': '&lt;',
        '>': '&gt;',
        '"': '&quot;',
        "'": '&#39;'
      }[ch]));
    }

    function chooseAssignment(headers) {
      const input = el('assignment');
      const current = input.value.trim();
      if (!headers || !headers.length) return;
      const exists = headers.some(h => String(h).trim().toLowerCase() === current.toLowerCase());
      if (!current || !exists) {
        input.value = headers[headers.length - 1];
      }
    }

    async function loadData() {
      setStatus('正在载入...');
      const data = await api('/data');
      chooseAssignment(data.workbook.headers);
      submissions = data.submissions.map(s => ({...s, graded: false, score: ''}));
      el('locked').style.display = data.workbook.lockedHint ? 'block' : 'none';
      renderList();
      await loadGrades();
      selectByFilteredIndex(0);
      setStatus(`共 ${submissions.length} 份作业，表格：${data.workbook.workbook}`);
    }

    async function loadGrades() {
      const assignment = el('assignment').value.trim();
      submissions.forEach(s => { s.graded = false; s.score = ''; });
      if (!assignment) {
        renderList();
        return;
      }
      try {
        const data = await api(`/grades?assignment=${encodeURIComponent(assignment)}`);
        submissions.forEach(s => {
          if (Object.prototype.hasOwnProperty.call(data.grades, s.id)) {
            s.graded = true;
            s.score = String(data.grades[s.id]);
          }
        });
      } catch (err) {
        setStatus(`当前列还没有评分记录：${assignment}`);
      }
      renderList();
    }

    function renderList() {
      const q = el('filter').value.trim().toUpperCase();
      filtered = submissions.map((s, i) => ({...s, realIndex: i}))
        .filter(s => !q || s.id.toUpperCase().includes(q) || String(s.studentName || s.name || '').toUpperCase().includes(q));
      el('list').innerHTML = '';
      filtered.forEach((s, i) => {
        const btn = document.createElement('button');
        btn.className = 'row' + (s.graded ? ' graded' : '') + (s.realIndex === index ? ' active' : '');
        const badge = s.graded ? `<span class="badge">已评分 ${s.score}</span>` : '';
        const name = s.studentName || s.name || '';
        const nameLine = name ? `<span class="studentname">${escapeHtml(name)}</span>` : '';
        btn.innerHTML = `<span class="sidline"><span class="sid">${escapeHtml(s.id)}</span>${badge}</span>${nameLine}<span class="filename">${escapeHtml(s.fileName || '未找到作业文件')}</span>`;
        btn.onclick = () => selectByRealIndex(s.realIndex);
        el('list').appendChild(btn);
      });
    }

    async function selectByFilteredIndex(i) {
      if (!filtered.length) return;
      i = Math.max(0, Math.min(filtered.length - 1, i));
      await selectByRealIndex(filtered[i].realIndex);
    }

    async function selectByRealIndex(i) {
      if (!submissions.length) return;
      index = Math.max(0, Math.min(submissions.length - 1, i));
      const s = submissions[index];
      el('sid').textContent = s.id;
      el('studentname').textContent = s.studentName || s.name || '';
      el('filename').textContent = s.fileName || s.folder;
      el('viewer').src = s.hasFile ? `/file?i=${index}` : 'about:blank';
      el('score').value = '';
      renderList();
      await loadScore();
      el('score').focus();
      el('score').select();
    }

    async function loadScore() {
      const s = submissions[index];
      const assignment = el('assignment').value.trim();
      if (!s || !assignment) return;
      try {
        const data = await api(`/score?id=${encodeURIComponent(s.id)}&assignment=${encodeURIComponent(assignment)}`);
        el('score').value = data.score || '';
        s.graded = !!data.score;
        s.score = data.score || '';
        renderList();
        setStatus(data.score ? `已读取 ${s.id} 的已有分数` : `当前：${s.id}`);
      } catch (err) {
        setStatus(`当前：${s.id}（${err.message}）`);
      }
    }

    async function save(stay) {
      const s = submissions[index];
      const assignment = el('assignment').value.trim();
      if (!s || !assignment) {
        setStatus('请填写成绩列名', true);
        return;
      }
      const score = el('score').value.trim();
      if (!score) {
        setStatus('请先输入分数', true);
        el('score').focus();
        return;
      }
      el('save').disabled = true;
      try {
        const data = await api('/save', {
          method: 'POST',
          headers: {'Content-Type': 'application/json'},
          body: JSON.stringify({
            id: s.id,
            assignment,
            score,
            createColumn: el('createColumn').checked
          })
        });
        setStatus(`已保存 ${s.id}：第 ${data.row} 行，第 ${data.col} 列`);
        s.graded = true;
        s.score = score;
        renderList();
        if (!stay) await go(1);
      } catch (err) {
        setStatus(err.message, true);
      } finally {
        el('save').disabled = false;
      }
    }

    async function go(delta) {
      const pos = filtered.findIndex(s => s.realIndex === index);
      await selectByFilteredIndex(pos + delta);
    }

    el('save').onclick = () => save(false);
    el('prev').onclick = () => go(-1);
    el('next').onclick = () => go(1);
    el('reload').onclick = loadData;
    el('filter').oninput = () => { renderList(); };
    el('assignment').addEventListener('change', async () => {
      await loadGrades();
      await loadScore();
    });
    el('score').addEventListener('keydown', (e) => {
      if (e.key === 'Enter') {
        e.preventDefault();
        save(e.ctrlKey);
      }
    });

    loadData().catch(err => setStatus(err.message, true));
  </script>
</body>
</html>
'@
$Html = $Html.Replace("__ASSIGNMENT__", [System.Net.WebUtility]::HtmlEncode($Assignment))

$prefix = "http://127.0.0.1:$Port/"
$listener = [System.Net.Sockets.TcpListener]::new([System.Net.IPAddress]::Parse("127.0.0.1"), $Port)
try {
    $listener.Start()
} catch {
    throw "Cannot start local server at $prefix. Try another port: powershell -ExecutionPolicy Bypass -File .\grader.ps1 -Port 8766"
}

Write-Host "Homework grader is running at $prefix"
Write-Host "Workbook: $WorkbookPath"
Write-Host "Submissions: $($script:Submissions.Count)"
Write-Host "Press Ctrl+C in this window to stop."
if (-not $NoOpen) {
    Start-Process $prefix
}

$running = $true
while ($running) {
    $client = $listener.AcceptTcpClient()
    $stream = $client.GetStream()
    try {
        $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8, $false, 8192, $true)
        $requestLine = $reader.ReadLine()
        if (-not $requestLine) {
            $client.Close()
            continue
        }

        $parts = $requestLine.Split(" ")
        $method = $parts[0]
        $target = $parts[1]
        $headers = @{}
        while ($true) {
            $line = $reader.ReadLine()
            if ($null -eq $line -or $line -eq "") { break }
            $colon = $line.IndexOf(":")
            if ($colon -gt 0) {
                $headers[$line.Substring(0, $colon).Trim().ToLowerInvariant()] = $line.Substring($colon + 1).Trim()
            }
        }

        $body = ""
        if ($headers.ContainsKey("content-length")) {
            $length = [int]$headers["content-length"]
            if ($length -gt 0) {
                $buffer = New-Object char[] $length
                $read = $reader.Read($buffer, 0, $length)
                $body = -join $buffer[0..($read - 1)]
            }
        }

        $uri = [System.Uri]::new("http://127.0.0.1$target")
        $requestPath = $uri.AbsolutePath
        $query = ConvertFrom-QueryString $uri.Query

        if ($requestPath -eq "/") {
            Write-Response $stream $Html "text/html; charset=utf-8"
        } elseif ($requestPath -eq "/data") {
            Update-Submissions -PreserveOnEmpty
            $info = Get-WorkbookInfo
            $studentInfo = Get-StudentInfoMap
            $responseSubmissions = @()
            foreach ($submission in $script:Submissions) {
                $key = $submission.id.Trim().ToUpperInvariant()
                $meta = if ($studentInfo.ContainsKey($key)) { $studentInfo[$key] } else { $null }
                $responseSubmissions += [pscustomobject][ordered]@{
                    id = $submission.id
                    name = if ($meta) { $meta.name } else { "" }
                    studentName = if ($meta) { $meta.name } else { "" }
                    row = if ($meta) { $meta.row } else { $null }
                    folder = $submission.folder
                    fileName = $submission.fileName
                    hasFile = $submission.hasFile
                    fileType = $submission.fileType
                    pdfName = $submission.pdfName
                    hasPdf = $submission.hasPdf
                }
            }
            Write-Response $stream (New-JsonResponse @{ ok = $true; submissions = $responseSubmissions; workbook = $info })
        } elseif ($requestPath -eq "/file" -or $requestPath -eq "/pdf") {
            $i = [int]$query["i"]
            if ($i -lt 0 -or $i -ge $script:Submissions.Count) {
                Write-Response $stream (New-JsonResponse @{ ok = $false; error = "Bad file index" }) "application/json; charset=utf-8" 400
            } else {
                $submission = $script:Submissions[$i]
                if (-not $submission.hasFile) {
                    Write-Response $stream (New-JsonResponse @{ ok = $false; error = "No submission file" }) "application/json; charset=utf-8" 404
                } else {
                    $ext = [System.IO.Path]::GetExtension($submission.filePath).ToLowerInvariant()
                    $contentType = if ($script:ContentTypes.ContainsKey($ext)) { $script:ContentTypes[$ext] } else { "application/octet-stream" }
                    Write-FileResponse $stream $submission.filePath $contentType
                }
            }
        } elseif ($requestPath -eq "/score") {
            $id = [string]$query["id"]
            $header = [string]$query["assignment"]
            $score = Get-Score $id $header
            Write-Response $stream (New-JsonResponse (@{ ok = $true } + $score))
        } elseif ($requestPath -eq "/grades") {
            $header = [string]$query["assignment"]
            try {
                $grades = Get-GradedMap $header
            } catch {
                $grades = @{}
            }
            Write-Response $stream (New-JsonResponse @{ ok = $true; grades = $grades })
        } elseif ($requestPath -eq "/save" -and $method -eq "POST") {
            $payload = $body | ConvertFrom-Json
            $result = Save-Score ([string]$payload.id) ([string]$payload.assignment) ([string]$payload.score) ([bool]$payload.createColumn)
            Write-Response $stream (New-JsonResponse (@{ ok = $true } + $result))
        } else {
            Write-Response $stream (New-JsonResponse @{ ok = $false; error = "Not found" }) "application/json; charset=utf-8" 404
        }
    } catch {
        Write-Response $stream (New-JsonResponse @{ ok = $false; error = $_.Exception.Message }) "application/json; charset=utf-8" 500
    } finally {
        $stream.Close()
        $client.Close()
    }
}
