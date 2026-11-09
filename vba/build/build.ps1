#requires -Version 5.1
<#
  Build add-in VBA "Nhật ký TC" bằng Excel trên Windows:
    1. Nạp mã nguồn ASCII (vba/build/src-ascii, sinh từ vba/src bằng vba/tools/vba_ascii.py) vào workbook mới.
    2. Biên dịch dự án VBA (Debug > Compile). Lỗi biên dịch -> Excel hiện hộp thoại, build dừng.
    3. Lưu thành vba/dist/NhatKyThiCong_Core.xlam, gắn Ribbon (customUI14.xml + ảnh nút).
    4. Mở lại file, chạy kiểm thử phần lõi (modNKTC_Tests.NKTC_RunTests).
  Kết quả ghi ra màn hình và vba/dist/ket-qua-build.txt.
  Yêu cầu một lần: Excel > File > Options > Trust Center > Trust Center Settings > Macro Settings >
  tích "Trust access to the VBA project object model".
#>
param([switch]$SkipTests)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.Encoding]::UTF8

$vbaDir = Split-Path -Parent $PSScriptRoot
$repoDir = Split-Path -Parent $vbaDir
$srcDir = Join-Path $vbaDir 'build\src-ascii'
$ribbonXml = Join-Path $vbaDir 'ribbon\customUI14.xml'
$iconDir = Join-Path $repoDir 'assets\ribbon'
$distDir = Join-Path $vbaDir 'dist'
$coreName = 'NhatKyThiCong_Core.xlam'
$corePath = Join-Path $distDir $coreName
$resultPath = Join-Path $distDir 'ket-qua-build.txt'
$xlOpenXMLAddIn = 55

New-Item -ItemType Directory -Force -Path $distDir | Out-Null
Set-Content -Path $resultPath -Value ("Build Nhật ký TC – " + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')) -Encoding UTF8
function Log([string]$message) {
    Write-Host $message
    Add-Content -Path $resultPath -Value $message -Encoding UTF8
}

# Ảnh nút: id trong customUI14.xml -> tên tệp trong assets/ribbon (bản 32 px).
$images = [ordered]@{
    imgTable = 'table'; imgExport = 'export'; imgTemplate = 'template'; imgAssign = 'assign'
    imgFirst = 'first'; imgPrev = 'prev'; imgNext = 'next'; imgLast = 'last'
    imgHelp = 'help'; imgAbout = 'about'
}

function Add-RibbonToPackage([string]$path) {
    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [IO.Compression.ZipFile]::Open($path, [IO.Compression.ZipArchiveMode]::Update)
    try {
        function Set-Entry([string]$name, [byte[]]$bytes) {
            $old = $zip.GetEntry($name)
            if ($old) { $old.Delete() }
            $stream = $zip.CreateEntry($name).Open()
            try { $stream.Write($bytes, 0, $bytes.Length) } finally { $stream.Dispose() }
        }
        function Get-EntryText([string]$name) {
            $reader = New-Object IO.StreamReader($zip.GetEntry($name).Open(), [Text.Encoding]::UTF8)
            try { return $reader.ReadToEnd() } finally { $reader.Dispose() }
        }
        $utf8 = New-Object Text.UTF8Encoding($false)

        Set-Entry 'customUI/customUI14.xml' ([IO.File]::ReadAllBytes($ribbonXml))
        $rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        foreach ($id in $images.Keys) {
            $file = $images[$id] + '.png'
            Set-Entry ('customUI/images/' + $file) ([IO.File]::ReadAllBytes((Join-Path $iconDir ($images[$id] + '-32.png'))))
            $rels += '<Relationship Id="' + $id + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="images/' + $file + '"/>'
        }
        $rels += '</Relationships>'
        Set-Entry 'customUI/_rels/customUI14.xml.rels' ($utf8.GetBytes($rels))

        $root = Get-EntryText '_rels/.rels'
        if ($root -notmatch 'customUI/customUI14\.xml') {
            $root = $root.Replace('</Relationships>', '<Relationship Id="rIdNhatKyTCUI" Type="http://schemas.microsoft.com/office/2007/relationships/ui/extensibility" Target="customUI/customUI14.xml"/></Relationships>')
            Set-Entry '_rels/.rels' ($utf8.GetBytes($root))
        }
        $types = Get-EntryText '[Content_Types].xml'
        if ($types -notmatch 'Extension="png"') {
            $types = $types.Replace('<Default Extension="xml"', '<Default Extension="png" ContentType="image/png"/><Default Extension="xml"')
            Set-Entry '[Content_Types].xml' ($utf8.GetBytes($types))
        }
    } finally {
        $zip.Dispose()
    }
}

$excel = $null
$exitCode = 1
try {
    if (Test-Path $corePath) { Remove-Item $corePath -Force }
    $excel = New-Object -ComObject Excel.Application
    # Hiện Excel để nếu có lỗi biên dịch thì thấy được hộp thoại báo lỗi.
    $excel.Visible = $true
    $excel.DisplayAlerts = $false

    $wb = $excel.Workbooks.Add()
    try {
        $project = $wb.VBProject
        $null = $project.VBComponents.Count
    } catch {
        throw 'Excel chưa cho phép truy cập dự án VBA. Vào File > Options > Trust Center > Trust Center Settings > Macro Settings, tích "Trust access to the VBA project object model", rồi chạy lại.'
    }

    Log 'Nạp mã nguồn:'
    Get-ChildItem -Path $srcDir -Include '*.bas', '*.cls' -File -Recurse | Sort-Object Name | ForEach-Object {
        [void]$project.VBComponents.Import($_.FullName)
        Log ('  + ' + $_.Name)
    }

    # Debug > Compile VBAProject (ID 578). Biên dịch xong thì lệnh bị làm mờ.
    $compile = $excel.VBE.CommandBars.FindControl(1, 578)
    if ($null -ne $compile -and $compile.Enabled) { $compile.Execute() }
    if ($null -ne $compile -and $compile.Enabled) {
        throw 'Biên dịch VBA có lỗi (xem hộp thoại và cửa sổ VBA trong Excel). Chụp màn hình gửi lại.'
    }
    Log 'Biên dịch: OK'

    $wb.SaveAs($corePath, $xlOpenXMLAddIn)
    $wb.Close($false)
    Add-RibbonToPackage $corePath
    Log ('Đã tạo: ' + $corePath)

    if ($SkipTests) {
        $exitCode = 0
    } else {
        $addin = $excel.Workbooks.Open($corePath)
        $result = [string]$excel.Run("'" + $coreName + "'!modNKTC_Tests.NKTC_RunTests")
        $addin.Close($false)
        Log ('Kiểm thử phần lõi: ' + $result)
        if ($result.StartsWith('PASS')) { $exitCode = 0 }
    }
} catch {
    Log ('LỖI: ' + $_.Exception.Message)
} finally {
    if ($null -ne $excel) {
        $excel.Quit()
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

if ($exitCode -eq 0) { Log 'KẾT QUẢ: ĐẠT' } else { Log 'KẾT QUẢ: CHƯA ĐẠT – gửi nội dung tệp ket-qua-build.txt cho Claude.' }
exit $exitCode
