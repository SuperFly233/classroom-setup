# ============================================================
#  机房一键装软件 · setup.ps1  by Ikokei
# ============================================================
$Host.UI.RawUI.WindowTitle = "机房一键装软件 · Ikokei"
function Write-Ok($m)   { Write-Host "  OK $m" -ForegroundColor Green }
function Write-Step($m) { Write-Host "  >> $m" -ForegroundColor Yellow }
function Write-Skip($m) { Write-Host "  -- $m" -ForegroundColor DarkGray }
function Write-Err($m)  { Write-Host "  !! $m" -ForegroundColor Red }
function Install-App {
    param([string]$Name,[string]$Url,[string]$Args="/S",[string]$DetectPath="")
    if ($DetectPath -and (Test-Path ([System.Environment]::ExpandEnvironmentVariables($DetectPath)))) {
        Write-Skip "$Name 已安装，跳过"; return
    }
    Write-Step "下载 $Name ..."
    $tmp = "$env:TEMP\ikokei_$Name.exe"
    try { (New-Object System.Net.WebClient).DownloadFile($Url, $tmp) }
    catch { Write-Err "$Name 下载失败: $($_.Exception.Message)"; return }
    Write-Step "安装 $Name ..."
    try { Start-Process -FilePath $tmp -ArgumentList $Args -Wait; Write-Ok "$Name 完成" }
    catch { Write-Err "$Name 安装失败: $($_.Exception.Message)" }
    Remove-Item $tmp -Force -ErrorAction SilentlyContinue
}
Clear-Host
Write-Host "  ==============================================" -ForegroundColor Cyan
Write-Host "     机房一键装软件  by Ikokei" -ForegroundColor Cyan
Write-Host "  ==============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  [ 软件安装 ]" -ForegroundColor Cyan

Install-App `
    -Name "WPS Office" `
    -Url "https://official-package.wpscdn.cn/wps/download/WPS_Setup_25225.exe" `
    -Args "/S /v/qn" `
    -DetectPath "C:\Program Files (x86)\Kingsoft\WPS Office"

Install-App `
    -Name "PixPin" `
    -Url "https://down.pixpin.cn/PixPin_cn_zh-cn_3.0.8.0.exe" `
    -Args "/S" `
    -DetectPath "%LOCALAPPDATA%\PixPin\PixPin.exe"

Install-App `
    -Name "OCS Helper" `
    -Url "https://mooc1.chaoxing.com/softdownload/ocsclient/OCS.exe" `
    -Args "/S" `
    -DetectPath "C:\Program Files (x86)\OCS\OCS.exe"
Write-Host "  [ 打开网站 ]" -ForegroundColor Cyan
Write-Step "打开网站..."
@(
    "https://mooc1.chaoxing.com" # chaoxing,
    "http://10.174.234.251:85/" # BJYSoft
) | ForEach-Object { Start-Process $_; Start-Sleep -Milliseconds 400 }
Write-Ok "网站已打开"
Write-Host ""
Write-Host "  ==============================================" -ForegroundColor Cyan
Write-Host "     全部完成！祝上课顺利 (o w o)" -ForegroundColor Cyan
Write-Host "  ==============================================" -ForegroundColor Cyan
Write-Host ""
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
