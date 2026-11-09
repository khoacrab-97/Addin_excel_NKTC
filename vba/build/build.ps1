#requires -Version 5.1
<#
  Build add-in VBA "Nhật ký TC" bằng Excel trên Windows:
    1. Nạp mã nguồn ASCII (vba/build/src-ascii, sinh từ vba/src bằng vba/tools/vba_ascii.py) vào workbook mới;
       tạo UserForm theo bố cục vba/forms/<tên>.json và chép mã src-ascii/<tên>.frmcode vào form.
    2. Biên dịch dự án VBA (Debug > Compile). Lỗi biên dịch -> Excel hiện hộp thoại, build dừng.
    3. Lưu thành vba/dist/NhatKyThiCong_Core.xlam, gắn Ribbon (customUI14.xml + ảnh nút).
    4. Dựng bộ nạp vba/dist/NhatKyThiCong.xlam (vba/build/loader-ascii, sinh từ vba/loader).
    5. Mở bộ nạp (bộ nạp tự mở lõi), chạy kiểm thử phần lõi (modNKTC_Tests.NKTC_RunTests).
    6. Ghi vba/dist/version.txt (phiên bản, tên tệp, SHA-256, kích thước của lõi) cho kênh cập nhật.
  Kết quả ghi ra màn hình và vba/dist/ket-qua-build.txt. Phát hành: tải 3 tệp NhatKyThiCong.xlam,
  NhatKyThiCong_Core.xlam, version.txt lên thư mục vba/release của repo (xem vba/README.md).
  Yêu cầu một lần: Excel > File > Options > Trust Center > Trust Center Settings > Macro Settings >
  tích "Trust access to the VBA project object model".
#>
param([switch]$SkipTests)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.Encoding]::UTF8

$vbaDir = Split-Path -Parent $PSScriptRoot
$repoDir = Split-Path -Parent $vbaDir
$srcDir = Join-Path $vbaDir 'build\src-ascii'
$loaderSrcDir = Join-Path $vbaDir 'build\loader-ascii'
$formsDir = Join-Path $vbaDir 'forms'
$templatePath = Join-Path $repoDir 'templates\Mau_NKTC.xlsx'
$ribbonXml = Join-Path $vbaDir 'ribbon\customUI14.xml'
$iconDir = Join-Path $repoDir 'assets\ribbon'
$distDir = Join-Path $vbaDir 'dist'
$coreName = 'NhatKyThiCong_Core.xlam'
$corePath = Join-Path $distDir $coreName
$loaderName = 'NhatKyThiCong.xlam'
$loaderPath = Join-Path $distDir $loaderName
$versionPath = Join-Path $distDir 'version.txt'
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

# Loại ô trong vba/forms/*.json -> ProgID của MSForms. Không dùng RefEdit/ActiveX: Microsoft 365 mới chặn
# chèn ActiveX ("The subject is not trusted"); ô chọn vùng = TextBox + nút Chọn (Application.InputBox).
$progIds = @{
    Label = 'Forms.Label.1'; TextBox = 'Forms.TextBox.1'; ComboBox = 'Forms.ComboBox.1'; ListBox = 'Forms.ListBox.1'
    CommandButton = 'Forms.CommandButton.1'; OptionButton = 'Forms.OptionButton.1'; CheckBox = 'Forms.CheckBox.1'
    Frame = 'Forms.Frame.1'
}
$geometryKeys = @('left', 'top', 'width', 'height')

function Set-UiFont($target) {
    try { $target.Font.Name = 'Segoe UI'; $target.Font.Size = 9 } catch { }
}

# Chạy một bước; lỗi -> báo rõ bước nào (thông báo gốc của COM/.NET thường không nói vị trí).
function Invoke-Step([string]$step, [scriptblock]$action) {
    try { & $action } catch { throw ($step + ': ' + $_.Exception.Message) }
}

function Add-UserForm($project, [string]$specPath) {
    $spec = Get-Content -Path $specPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $formName = [string]$spec.name
    $component = $null
    Invoke-Step ($formName + ' – tạo form') { $script:component = $project.VBComponents.Add(3) }   # vbext_ct_MSForm
    $component = $script:component
    Invoke-Step ($formName + ' – đặt tên') { $component.Name = $formName }
    Invoke-Step ($formName + ' – tiêu đề') { $component.Properties.Item('Caption').Value = [string]$spec.caption }
    Invoke-Step ($formName + ' – rộng') { $component.Properties.Item('Width').Value = [single]$spec.width }
    Invoke-Step ($formName + ' – cao') { $component.Properties.Item('Height').Value = [single]$spec.height }
    $designer = $component.Designer
    Set-UiFont $designer
    foreach ($c in $spec.controls) {
        $controlName = [string]$c.name
        $progId = $progIds[[string]$c.type]
        if (-not $progId) { throw ('Loại ô không hỗ trợ "' + $c.type + '" trong ' + $specPath) }
        $script:control = $null
        Invoke-Step ($formName + ' – thêm ô ' + $c.type + ' "' + $controlName + '"') {
            $script:control = $designer.Controls.Add([string]$progId, [string]$controlName, $true)
        }
        $control = $script:control
        Set-UiFont $control
        foreach ($p in $c.PSObject.Properties) {
            if ($p.Name -eq 'type' -or $p.Name -eq 'name') { continue }
            $value = $p.Value
            if ($geometryKeys -contains $p.Name) { $value = [single]$value }
            $propName = $p.Name
            Invoke-Step ($formName + ' – ' + $controlName + '.' + $propName) { $control.$propName = $value }
        }
    }
    $codePath = Join-Path $srcDir ($formName + '.frmcode')
    if (-not (Test-Path $codePath)) { throw ('Thiếu mã của form: ' + $codePath) }
    Invoke-Step ($formName + ' – chép mã') {
        $module = $component.CodeModule
        $count = [int]$module.CountOfLines
        if ($count -gt 0) { $module.DeleteLines(1, $count) }
        $module.AddFromString([IO.File]::ReadAllText($codePath))
    }
}

# Module ThisWorkbook của workbook (tên mã thường là "ThisWorkbook"; dò theo loại nếu Excel chưa trả về).
function Get-WorkbookModule($wb) {
    $name = [string]$wb.CodeName
    if ($name) { return $wb.VBProject.VBComponents.Item($name).CodeModule }
    $sheetNames = @()
    foreach ($s in $wb.Worksheets) { $sheetNames += [string]$s.CodeName }
    foreach ($c in $wb.VBProject.VBComponents) {
        if ([int]$c.Type -eq 100 -and ($sheetNames -notcontains [string]$c.Name)) { return $c.CodeModule }
    }
    throw 'Không tìm thấy module ThisWorkbook.'
}

function Set-ModuleCode($module, [string]$path) {
    $count = [int]$module.CountOfLines
    if ($count -gt 0) { $module.DeleteLines(1, $count) }
    $module.AddFromString([IO.File]::ReadAllText($path))
}

# Debug > Compile VBAProject (ID 578) cho dự án đang chọn trong VBE. Biên dịch xong thì lệnh bị làm mờ.
function Invoke-Compile($excel, $project, [string]$label) {
    $excel.VBE.ActiveVBProject = $project
    $compile = $excel.VBE.CommandBars.FindControl(1, 578)
    if ($null -ne $compile -and $compile.Enabled) { $compile.Execute() }
    if ($null -ne $compile -and $compile.Enabled) {
        throw ('Biên dịch VBA (' + $label + ') có lỗi (xem hộp thoại và cửa sổ VBA trong Excel). Chụp màn hình gửi lại.')
    }
    Log ('Biên dịch ' + $label + ': OK')
}

function Import-Modules($project, [string]$dir) {
    Get-ChildItem -Path $dir -Include '*.bas', '*.cls' -File -Recurse | Sort-Object Name | ForEach-Object {
        [void]$project.VBComponents.Import($_.FullName)
        Log ('  + ' + $_.Name)
    }
}

$excel = $null
$exitCode = 1
try {
    foreach ($old in @($corePath, $loaderPath, $versionPath)) { if (Test-Path $old) { Remove-Item $old -Force } }
    $excel = New-Object -ComObject Excel.Application
    # Hiện Excel để nếu có lỗi biên dịch thì thấy được hộp thoại báo lỗi.
    $excel.Visible = $true
    $excel.DisplayAlerts = $false

    # ---------- Lõi: NhatKyThiCong_Core.xlam ----------
    $wb = $excel.Workbooks.Add()
    try {
        $project = $wb.VBProject
        $null = $project.VBComponents.Count
    } catch {
        throw 'Excel chưa cho phép truy cập dự án VBA. Vào File > Options > Trust Center > Trust Center Settings > Macro Settings, tích "Trust access to the VBA project object model", rồi chạy lại.'
    }

    Log 'Nạp mã nguồn lõi:'
    Import-Modules $project $srcDir
    Get-ChildItem -Path $formsDir -Filter '*.json' -File | Sort-Object Name | ForEach-Object {
        Add-UserForm $project $_.FullName
        Log ('  + ' + $_.BaseName + ' (UserForm)')
    }
    Set-ModuleCode (Get-WorkbookModule $wb) (Join-Path $srcDir 'ThisWorkbook.doccode')
    Log '  + ThisWorkbook'

    # Biểu mẫu "Mẫu NKTC": chép vào add-in để nút Chèn form mẫu dùng được khi không có Internet.
    $template = $excel.Workbooks.Open($templatePath, 0, $true)
    try {
        $formSheetName = [string]$template.Worksheets.Item(1).Name
        $template.Worksheets.Item(1).Copy([Reflection.Missing]::Value, $wb.Worksheets.Item($wb.Worksheets.Count))
    } finally {
        $template.Close($false)
    }
    for ($i = [int]$wb.Worksheets.Count; $i -ge 1; $i--) {
        $sheet = $wb.Worksheets.Item($i)
        if ([string]$sheet.Name -ne $formSheetName) { $sheet.Delete() }
    }
    Log ('  + biểu mẫu "' + $formSheetName + '"')

    Invoke-Compile $excel $project 'lõi'
    $wb.SaveAs($corePath, $xlOpenXMLAddIn)
    $wb.Close($false)
    Add-RibbonToPackage $corePath
    Log ('Đã tạo: ' + $corePath)

    # ---------- Bộ nạp: NhatKyThiCong.xlam ----------
    $loaderWb = $excel.Workbooks.Add()
    Log 'Nạp mã nguồn bộ nạp:'
    Import-Modules $loaderWb.VBProject $loaderSrcDir
    Set-ModuleCode (Get-WorkbookModule $loaderWb) (Join-Path $loaderSrcDir 'ThisWorkbook.doccode')
    Log '  + ThisWorkbook'
    Invoke-Compile $excel $loaderWb.VBProject 'bộ nạp'
    $loaderWb.SaveAs($loaderPath, $xlOpenXMLAddIn)
    $loaderWb.Close($false)
    Log ('Đã tạo: ' + $loaderPath)

    if ($SkipTests) {
        $exitCode = 0
    } else {
        # Mở bộ nạp như khi Excel khởi động: bộ nạp phải tự mở lõi.
        $loader = $excel.Workbooks.Open($loaderPath)
        $core = $null
        foreach ($w in $excel.Workbooks) { if ([string]$w.Name -eq $coreName) { $core = $w } }
        if ($null -eq $core) { throw 'Bộ nạp không mở được lõi NhatKyThiCong_Core.xlam.' }
        Log 'Bộ nạp mở lõi: OK'
        $result = [string]$excel.Run("'" + $coreName + "'!modNKTC_Tests.NKTC_RunTests")
        $version = [string]$excel.Run("'" + $coreName + "'!modNKTC_Ribbon.NKTC_Version")
        $core.Close($false)
        $loader.Close($false)
        Log ('Kiểm thử phần lõi: ' + $result)
        if ($result.StartsWith('PASS')) { $exitCode = 0 }

        # version.txt cho kênh cập nhật (tệp văn bản UTF-8 không BOM, xuống dòng LF).
        $hash = (Get-FileHash -Path $corePath -Algorithm SHA256).Hash.ToLowerInvariant()
        $size = (Get-Item $corePath).Length
        $manifest = "version=$version`nfile=$coreName`nsha256=$hash`nsize=$size`n"
        [IO.File]::WriteAllText($versionPath, $manifest, (New-Object Text.UTF8Encoding($false)))
        Log ('Đã tạo: ' + $versionPath + ' (phiên bản ' + $version + ')')
    }
} catch {
    Log ('LỖI: ' + $_.Exception.Message)
    Log ('  (build.ps1 dòng ' + $_.InvocationInfo.ScriptLineNumber + ': ' + $_.InvocationInfo.Line.Trim() + ')')
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
