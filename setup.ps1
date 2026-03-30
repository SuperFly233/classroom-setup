# ============================================================
#  机房一键装软件 · setup.ps1
#  by Ikokei · https://github.com/SuperFly233
# ============================================================

$Host.UI.RawUI.WindowTitle = "机房一键装软件 · Ikokei"

function Write-Header {
    Clear-Host
    Write-Host ""
    Write-Host "  ╔══════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "  ║     机房一键装软件  ·  by Ikokei          ║" -ForegroundColor Cyan
    Write-Host "  ╚══════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host ""
}

function Write-Step($msg) {
    Write-Host "  ▶ $msg" -ForegroundColor Yellow
}

function Write-Ok($msg) {
    Write-Host "  ✓ $msg" -ForegroundColor Green
}

function Write-Skip($msg) {
    Write-Host "  · $msg" -ForegroundColor DarkGray
}

function Write-Err($msg) {
    Write-Host "  ✗ $msg" -ForegroundColor Red
}

# ── 静默安装软件函数 ──────────────────────────────────────────
function Install-App {
    param(
        [string]$Name,
        [string]$Url,
        [string]$Args = "/S",
        [string]$DetectPath = ""
    )

    # 检测是否已安装
    if ($DetectPath -and (Test-Path $DetectPath)) {
        Write-Skip "$Name 已安装，跳过"
        return
    }

    Write-Step "下载 $Name ..."
    $tmp = "$env:TEMP\ikokei_$Name.exe"

    try {
        $wc = New-Object System.Net.WebClient
        $wc.DownloadFile($Url, $tmp)
    } catch {
        Write-Err "$Name 下载失败：$($_.Exception.Message)"
        return
    }

    Write-Step "安装 $Name ..."
    try {
        Start-Process -FilePath $tmp -ArgumentList $Args -Wait -ErrorAction Stop
        Write-Ok "$Name 安装完成"
    } catch {
        Write-Err "$Name 安装失败：$($_.Exception.Message)"
    }

    Remove-Item $tmp -Force -ErrorAction SilentlyContinue
}

# ── 打开网站函数 ──────────────────────────────────────────────
function Open-Sites {
    param([string[]]$Urls)
    Write-Step "打开上课网站..."
    foreach ($url in $Urls) {
        Start-Process $url
        Start-Sleep -Milliseconds 400
    }
    Write-Ok "网站已在浏览器中打开"
}

# ════════════════════════════════════════════════════════════
#  主流程
# ════════════════════════════════════════════════════════════
Write-Header

# ── 1. 安装软件 ───────────────────────────────────────────────
Write-Host "  [ 软件安装 ]" -ForegroundColor Cyan
Write-Host ""

Install-App `
    -Name "WPS Office" `
    -Url "https://wdl1.pcfg.cache.wpscdn.cn/wpsdl/wpsoffice/download/wps_office_free.exe" `
    -Args "/S /v/qn" `
    -DetectPath "C:\Program Files (x86)\Kingsoft\WPS Office"

Install-App `
    -Name "PixPin" `
    -Url "https://download.pixpinapp.com/PixPin_latest.exe" `
    -Args "/S" `
    -DetectPath "$env:LOCALAPPDATA\PixPin\PixPin.exe"

Install-App `
    -Name "OCS 超星客户端" `
    -Url "https://mooc1.chaoxing.com/softdownload/ocsclient/OCS.exe" `
    -Args "/S" `
    -DetectPath "C:\Program Files (x86)\OCS\OCS.exe"

Write-Host ""

# ── 2. 打开网站 ───────────────────────────────────────────────
Write-Host "  [ 打开网站 ]" -ForegroundColor Cyan
Write-Host ""

Open-Sites @(
    "https://mooc1.chaoxing.com",
    "http://10.174.234.251:85/"
    # 在这里继续添加其他网址
)

Write-Host ""
Write-Host "  ╔══════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "  ║     全部完成！祝上课顺利 (・ω・)ノ          ║" -ForegroundColor Cyan
Write-Host "  ╚══════════════════════════════════════════╝" -ForegroundColor Cyan
Write-Host ""
Write-Host "  按任意键关闭..." -ForegroundColor DarkGray
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
