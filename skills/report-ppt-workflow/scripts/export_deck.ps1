# 导出/渲染 pptx -> PDF + 逐页PNG(Windows,PowerPoint COM)
# 用法: powershell -File export_deck.ps1 deck.pptx [输出目录]
param(
    [Parameter(Mandatory = $true)][string]$Pptx,
    [string]$OutDir = ""
)
$ErrorActionPreference = "Stop"
$pptxAbs = (Resolve-Path $Pptx).Path   # 必须 .Path 取字符串,PathInfo 对象会让 COM 报 E_FAIL
if (-not $OutDir) { $OutDir = Join-Path (Split-Path $pptxAbs) "export" }
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null

$pp = New-Object -ComObject PowerPoint.Application
try {
    # msoTrue=-1, msoFalse=0。注意:WithWindow 必须为 msoTrue,否则部分环境 E_FAIL
    $pres = $pp.Presentations.Open($pptxAbs, -1, 0, -1)
    # PDF
    $pres.SaveAs((Join-Path $OutDir "deck.pdf"), 32)          # 32 = ppSaveAsPDF
    # 逐页 PNG(验收/传阅用)
    $pngDir = Join-Path $OutDir "png"
    New-Item -ItemType Directory -Force -Path $pngDir | Out-Null
    $pres.Export($pngDir, "PNG", 1920, 1080)
    $pres.Close()
} finally { $pp.Quit() }

Write-Output "完成: $OutDir\deck.pdf ; $pngDir\*.png"
