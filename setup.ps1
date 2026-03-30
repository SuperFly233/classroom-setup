# ============================================================
#  机房一键装软件 · setup.ps1
#  by Ikokei · https://github.com/SuperFly233
# ============================================================

# ── TLS 兼容修复（老机器必备）────────────────────────────────
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$Host.UI.RawUI.WindowTitle = "机房一键装软件 · Ikokei"

# ════════════════════════════════════════════════════════════
#  工具函数
# ════════════════════════════════════════════════════════════

function Write-Banner {
    Clear-Host
    Write-Host ""
    Write-Host "  ╔══════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "  ║       机房一键装软件  ·  by Ikokei               ║" -ForegroundColor Cyan
    Write-Host "  ║       github.com/SuperFly233/classroom-setup     ║" -ForegroundColor DarkCyan
    Write-Host "  ╚══════════════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host ""
}

function Write-Section($title) {
    Write-Host ""
    Write-Host "  ┌─ $title " -ForegroundColor Cyan -NoNewline
    Write-Host ("─" * [Math]::Max(2, 44 - $title.Length)) -ForegroundColor DarkCyan
    Write-Host ""
}

function Write-Step($msg)  { Write-Host "  ▶  $msg" -ForegroundColor Yellow }
function Write-Ok($msg)    { Write-Host "  ✓  $msg" -ForegroundColor Green }
function Write-Skip($msg)  { Write-Host "  ·  $msg" -ForegroundColor DarkGray }
function Write-Info($msg)  { Write-Host "  ℹ  $msg" -ForegroundColor Cyan }
function Write-Err($msg)   { Write-Host "  ✗  $msg" -ForegroundColor Red }
function Write-Warn($msg)  { Write-Host "  ⚠  $msg" -ForegroundColor Yellow }

# 格式化文件大小
function Format-Size($bytes) {
    if ($bytes -ge 1GB) { return "{0:F2} GB" -f ($bytes / 1GB) }
    if ($bytes -ge 1MB) { return "{0:F1} MB" -f ($bytes / 1MB) }
    if ($bytes -ge 1KB) { return "{0:F0} KB" -f ($bytes / 1KB) }
    return "$bytes B"
}

# 格式化速度
function Format-Speed($bps) {
    if ($bps -ge 1MB) { return "{0:F1} MB/s" -f ($bps / 1MB) }
    if ($bps -ge 1KB) { return "{0:F0} KB/s" -f ($bps / 1KB) }
    return "$bps B/s"
}

# 格式化剩余时间
function Format-ETA($seconds) {
    if ($seconds -le 0 -or $seconds -gt 3600) { return "计算中…" }
    if ($seconds -ge 60) { return "{0}m {1:D2}s" -f [int]($seconds/60), ($seconds%60) }
    return "{0}s" -f [int]$seconds
}

# ════════════════════════════════════════════════════════════
#  带进度条的下载函数 (主线程安全版)
# ════════════════════════════════════════════════════════════
function Invoke-Download {
    param(
        [string]$Name,
        [string]$Url,
        [string]$OutFile
    )

    Write-Host ""
    $startTick = [DateTime]::Now

    try {
        # 使用 WebRequest 同步请求，避免多线程崩溃
        $request = [System.Net.WebRequest]::Create($Url)
        $request.Method = "GET"
        $response = $request.GetResponse()
        
        $totalBytes = $response.ContentLength
        $stream = $response.GetResponseStream()

        # 每次读取 8KB
        $buffer = New-Object byte[] 8192
        $fileStream = [System.IO.File]::Create($OutFile)

        $downloadedBytes = 0
        $lastTick = $startTick
        $lastBytes = 0
        $speed = 0

        # 主线程循环读取流
        while ($true) {
            $read = $stream.Read($buffer, 0, $buffer.Length)
            if ($read -le 0) { break } # 读取完毕
            
            $fileStream.Write($buffer, 0, $read)
            $downloadedBytes += $read

            $now = [DateTime]::Now
            $elapsed = ($now - $lastTick).TotalSeconds

            # 每 0.3 秒刷新一次屏幕，防止输出过快导致卡顿
            if ($elapsed -ge 0.3) {
                $delta = $downloadedBytes - $lastBytes
                $speed = $delta / $elapsed
                $lastBytes = $downloadedBytes
                $lastTick = $now

                $pct = 0
                if ($totalBytes -gt 0) {
                    $pct = [math]::Floor(($downloadedBytes / $totalBytes) * 100)
                }

                $remaining = $totalBytes - $downloadedBytes
                $eta = if ($speed -gt 0) { $remaining / $speed } else { 0 }

                $recvStr  = Format-Size $downloadedBytes
                $totalStr = if ($totalBytes -gt 0) { Format-Size $totalBytes } else { "未知大小" }
                $speedStr = if ($speed -gt 0) { Format-Speed $speed } else { "—" }
                $etaStr   = Format-ETA $eta

                # 进度条体
                $barWidth = 36
                $filled   = [int]($barWidth * $pct / 100)
                $empty    = $barWidth - $filled
                $bar      = ("█" * $filled) + ("░" * $empty)

                $status = "  $recvStr / $totalStr   $speedStr   ETA $etaStr"

                Write-Progress `
                    -Activity "  下载 $Name" `
                    -Status $status `
                    -PercentComplete $pct `
                    -CurrentOperation "[$bar] $pct%"
            }
        }
    }
    catch {
        # 抛出异常让外层的 Install-App 捕获
        throw $_ 
    }
    finally {
        # 确保清理释放资源，解除文件占用
        if ($null -ne $fileStream) { $fileStream.Dispose() }
        if ($null -ne $stream)     { $stream.Dispose() }
        if ($null -ne $response)   { $response.Dispose() }
    }

    Write-Progress -Activity "  下载 $Name" -Completed

    # 下载完成摘要
    $totalTime = ([DateTime]::Now - $startTick).TotalSeconds
    $totalBytesToReport = if ($totalBytes -gt 0) { $totalBytes } else { $downloadedBytes }
    $avgSpeed  = if ($totalTime -gt 0) { $totalBytesToReport / $totalTime } else { 0 }
    
    Write-Ok ("下载完成  " + (Format-Size $totalBytesToReport) + "  均速 " + (Format-Speed $avgSpeed) + "  用时 {0:F1}s" -f $totalTime)
}

# ════════════════════════════════════════════════════════════
#  安装函数
# ════════════════════════════════════════════════════════════
function Install-App {
    param(
        [string]$Name,
        [string]$Url,
        [string]$Args       = "/S",
        [string]$DetectPath = ""
    )

    Write-Host "  ┄┄ $Name " -ForegroundColor DarkCyan -NoNewline
    Write-Host ("┄" * [Math]::Max(2, 36 - $Name.Length)) -ForegroundColor DarkCyan

    # 检测已安装
    $expanded = [System.Environment]::ExpandEnvironmentVariables($DetectPath)
    if ($DetectPath -and (Test-Path $expanded)) {
        Write-Skip "$Name 已安装，跳过"
        Write-Host ""
        return
    }

    $tmp = "$env:TEMP\ikokei_$Name.exe"

    # 下载
    Write-Step "开始下载…"
    try {
        Invoke-Download -Name $Name -Url $Url -OutFile $tmp
    }
    catch {
        Write-Err "下载失败：$($_.Exception.Message)"
        Write-Warn "你可以手动从以下地址下载后运行："
        Write-Host "     $Url" -ForegroundColor DarkGray
        Write-Host ""
        return
    }

    # 校验文件
    if (-not (Test-Path $tmp) -or (Get-Item $tmp).Length -lt 1KB) {
        Write-Err "文件无效，跳过安装"
        Write-Host ""
        return
    }

    # 安装
    Write-Step "正在安装，请稍候…"
    try {
        $proc = Start-Process -FilePath $tmp -ArgumentList $Args -Wait -PassThru -ErrorAction Stop
        if ($proc.ExitCode -eq 0 -or $proc.ExitCode -eq 3010) {
            if ($proc.ExitCode -eq 3010) {
                Write-Ok "$Name 安装完成（需重启生效）"
            } else {
                Write-Ok "$Name 安装完成"
            }
        } else {
            Write-Warn "$Name 安装程序退出码：$($proc.ExitCode)（可能已安装或需手动确认）"
        }
    }
    catch {
        Write-Err "安装失败：$($_.Exception.Message)"
    }

    Remove-Item $tmp -Force -ErrorAction SilentlyContinue
    Write-Host ""
}

# ════════════════════════════════════════════════════════════
#  主流程
# ════════════════════════════════════════════════════════════
Write-Banner

# ── 环境检测 ──────────────────────────────────────────────────
Write-Section "环境检测"

$psVer = $PSVersionTable.PSVersion
Write-Info "PowerShell $($psVer.Major).$($psVer.Minor)"

$os = (Get-WmiObject Win32_OperatingSystem -ErrorAction SilentlyContinue)
if ($os) { Write-Info "系统：$($os.Caption)  $($os.OSArchitecture)" }

$disk = Get-PSDrive C -ErrorAction SilentlyContinue
if ($disk) {
    $free = $disk.Free
    $freeStr = Format-Size $free
    if ($free -lt 500MB) {
        Write-Warn "C 盘剩余空间不足 500MB（$freeStr），安装可能失败"
    } else {
        Write-Info "C 盘可用空间：$freeStr"
    }
}

# 网络连通性检测
Write-Step "检测网络连通性…"
try {
    $ping = Test-Connection "8.8.8.8" -Count 1 -Quiet -ErrorAction Stop
    if ($ping) { Write-Ok "网络正常" }
    else { Write-Warn "网络可能不稳定" }
} catch {
    Write-Warn "网络检测失败，继续尝试…"
}

# ── 软件安装 ──────────────────────────────────────────────────
Write-Section "软件安装"

Install-App `
    -Name "WPS Office" `
    -Url "https://wdl1.pcfg.cache.wpscdn.cn/wpsdl/wpsoffice/download/wps_office_free.exe" `
    -Args "/S /v/qn" `
    -DetectPath "C:\Program Files (x86)\Kingsoft\WPS Office"

Install-App `
    -Name "PixPin" `
    -Url "https://download.pixpinapp.com/PixPin_latest.exe" `
    -Args "/S" `
    -DetectPath "%LOCALAPPDATA%\PixPin\PixPin.exe"

Install-App `
    -Name "OCS 超星客户端" `
    -Url "https://mooc1.chaoxing.com/softdownload/ocsclient/OCS.exe" `
    -Args "/S" `
    -DetectPath "C:\Program Files (x86)\OCS\OCS.exe"

# ── 打开网站 ──────────────────────────────────────────────────
Write-Section "打开网站"

$sites = @(
    @{ Name = "超星学习通";   Url = "https://mooc1.chaoxing.com" },
    @{ Name = "内网实训平台"; Url = "http://10.174.234.251:85/"  }
)

foreach ($site in $sites) {
    Write-Step "打开 $($site.Name)…"
    Start-Process $site.Url
    Start-Sleep -Milliseconds 500
}
Write-Ok "已在浏览器中打开 $($sites.Count) 个网站"

# ── 完成 ──────────────────────────────────────────────────────
Write-Host ""
Write-Host "  ╔══════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "  ║       全部完成！祝上课顺利  (・ω・)ノ              ║" -ForegroundColor Cyan
Write-Host "  ╚══════════════════════════════════════════════════╝" -ForegroundColor Cyan
Write-Host ""
Write-Host "  按任意键关闭…" -ForegroundColor DarkGray
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
