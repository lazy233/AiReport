# PPT 解析平台：在项目根目录运行本脚本启动本地 Flask（调试模式）
$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest
Set-Location -LiteralPath $PSScriptRoot

# 可选：从 .env 读取环境变量（格式见 .env.example）
$envFile = Join-Path $PSScriptRoot ".env"
if (Test-Path -LiteralPath $envFile) {
    Get-Content -LiteralPath $envFile | ForEach-Object {
        $line = $_.Trim()
        if ($line -eq "" -or $line.StartsWith("#")) { return }
        $idx = $line.IndexOf("=")
        if ($idx -lt 1) { return }
        $name = $line.Substring(0, $idx).Trim()
        $value = $line.Substring($idx + 1).Trim()
        [Environment]::SetEnvironmentVariable($name, $value, "Process")
    }
}

# PostgreSQL：仅在本机 5432 尚不可连时尝试拉起（服务多为「自动」时已就绪，无需提权）
function Test-PgPortOpen {
    try {
        $tcp = New-Object System.Net.Sockets.TcpClient
        $tcp.Connect("127.0.0.1", 5432)
        $tcp.Close()
        return $true
    } catch {
        return $false
    }
}

if (Test-PgPortOpen) {
    Write-Host "127.0.0.1:5432 已可连接，跳过数据库启动步骤。" -ForegroundColor DarkGray
} else {
    $pgSvcs = @(Get-Service -ErrorAction SilentlyContinue | Where-Object {
            $_.Name -like "postgresql*" -or ($_.DisplayName -like "*PostgreSQL*")
        })
    if ($pgSvcs.Count -gt 0) {
        $running = @($pgSvcs | Where-Object { $_.Status -eq "Running" })
        if ($running.Count -gt 0) {
            Write-Host ("PostgreSQL 服务已在运行: " + (($running | ForEach-Object { $_.Name }) -join ", ")) -ForegroundColor DarkGray
        }
        # 只启动第一个未运行服务，避免多实例同时拉起 / 重复占端口
        $firstStopped = @($pgSvcs | Where-Object { $_.Status -ne "Running" }) | Select-Object -First 1
        if ($firstStopped) {
            try {
                Start-Service -InputObject $firstStopped -ErrorAction Stop
                Write-Host "已启动 PostgreSQL 服务: $($firstStopped.Name)" -ForegroundColor Cyan
                Start-Sleep -Seconds 2
            } catch {
                # 当前账户可能无权操作服务（属正常策略），改试 pg_ctl / Docker，不强制要求管理员
            }
        }
    }

    Start-Sleep -Seconds 1
    if (-not (Test-PgPortOpen)) {
        $regRoot = "HKLM:\SOFTWARE\PostgreSQL\Installations"
        if (Test-Path -LiteralPath $regRoot) {
            # 本脚本只对注册表第一项有效安装执行一次 pg_ctl（端口已通则不应再调）
            $pgCtlInvoke = $null
            $pgCtlDataDir = $null
            foreach ($inst in @(Get-ChildItem -LiteralPath $regRoot -ErrorAction SilentlyContinue)) {
                $prop = Get-ItemProperty -LiteralPath $inst.PSPath -ErrorAction SilentlyContinue
                $baseDir = $prop."Base Directory"
                $dataDir = $prop."Data Directory"
                if (-not $baseDir -or -not $dataDir) { continue }
                $pgCtlExe = Join-Path $baseDir "bin\pg_ctl.exe"
                if (-not (Test-Path -LiteralPath $pgCtlExe)) { continue }
                $pgCtlInvoke = $pgCtlExe
                $pgCtlDataDir = $dataDir
                break
            }
            if ($pgCtlInvoke -and $pgCtlDataDir) {
                $logFile = Join-Path $env:TEMP "ppt-report-pg_ctl.log"
                # 不用直接调用 & pg_ctl，以免 stderr 在 ErrorAction=Stop 时终止整段脚本
                $proc = Start-Process -FilePath $pgCtlInvoke -ArgumentList @(
                    "start", "-w", "-D", $pgCtlDataDir, "-l", $logFile
                ) -Wait -NoNewWindow -PassThru
                if ($proc.ExitCode -eq 0) {
                    Write-Host "已通过 pg_ctl 启动 PostgreSQL（日志: $logFile）" -ForegroundColor Cyan
                }
            }
        }
    }

    Start-Sleep -Seconds 1
    if (-not (Test-PgPortOpen)) {
        $docker = Get-Command docker -ErrorAction SilentlyContinue
        if ($docker) {
            $fmt = docker ps -a --format "{{.Names}}|{{.Image}}|{{.State}}" 2>$null
            foreach ($row in $fmt) {
                if (-not $row) { continue }
                $p = $row -split "\|", 3
                if ($p.Count -lt 3) { continue }
                if ($p[2] -notmatch "^(exited|created)$") { continue }
                if ($p[1] -notmatch "postgres") { continue }
                $null = Start-Process -FilePath $docker.Source -ArgumentList "start", $p[0] -Wait -NoNewWindow
                break
            }
        }
    }

    Start-Sleep -Seconds 1
    if (-not (Test-PgPortOpen)) {
        Write-Host "未能自动连通 127.0.0.1:5432，Flask 仍会启动（数据库不可用时会用内存缓存）。请在本机打开 PostgreSQL 服务或 Docker 容器后再试连接。" -ForegroundColor DarkYellow
    }
}


$venvPython = Join-Path $PSScriptRoot ".venv\Scripts\python.exe"
$appPy = Join-Path $PSScriptRoot "app.py"
if (-not (Test-Path -LiteralPath $appPy)) {
    Write-Error "未找到 app.py，请在项目根目录执行 start.ps1。"
}

Write-Host "正在启动 Flask（调试）http://127.0.0.1:5000/ - 按 Ctrl+C 停止" -ForegroundColor Green
# 避免使用 PATH 里 Microsoft Store 的 python 占位程序（WindowsApps），其会秒退。
if (Test-Path -LiteralPath $venvPython) {
    & $venvPython $appPy
} elseif (Get-Command py -ErrorAction SilentlyContinue) {
    & py -3 $appPy
} else {
    $winPy = Get-Command python -ErrorAction SilentlyContinue
    if ($winPy -and ($winPy.Source -notmatch "\\WindowsApps\\")) {
        & $winPy.Source $appPy
    } else {
        Write-Error "未检测到可用的 Python。请安装 Python（确保可使用 py -3）或在根目录创建 .venv。"
    }
}
