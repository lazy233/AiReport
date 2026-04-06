# 本地开发：启动 PostgreSQL（若未监听 5432），再启动 Flask（同时提供接口与静态页面）
$ErrorActionPreference = "Stop"
Set-Location -LiteralPath $PSScriptRoot

function Test-PgPort {
    try {
        $c = New-Object System.Net.Sockets.TcpClient
        $c.Connect("127.0.0.1", 5432)
        $c.Close()
        return $true
    } catch {
        return $false
    }
}

function Invoke-StartPostgreSqlServicesElevated {
    $elevPs = Join-Path $env:WINDIR "System32\WindowsPowerShell\v1.0\powershell.exe"
    if (-not (Test-Path -LiteralPath $elevPs)) {
        Write-Warning "未找到 $elevPs，无法弹出 UAC 提权启动 PostgreSQL。"
        return
    }
    $helper = Join-Path $env:TEMP "ppt-report-start-postgres-$(Get-Random).ps1"
    @'
$ErrorActionPreference = "Stop"
Get-Service -ErrorAction SilentlyContinue | Where-Object {
    ($_.Name -like "postgresql*" -or $_.DisplayName -like "*PostgreSQL*") -and $_.Status -ne "Running"
} | ForEach-Object {
    Start-Service -InputObject $_
    Write-Host ("已启动: " + $_.Name)
}
'@ | Set-Content -LiteralPath $helper -Encoding UTF8
    try {
        Write-Host "PostgreSQL 需要管理员权限启动，请在 UAC 窗口中选「是」…" -ForegroundColor Yellow
        $p = Start-Process -FilePath $elevPs -Verb RunAs -Wait -PassThru `
            -ArgumentList @("-NoProfile", "-ExecutionPolicy", "Bypass", "-File", $helper)
        if ($p.ExitCode -ne 0) {
            Write-Warning "提权脚本退出码: $($p.ExitCode)"
        }
    } finally {
        Remove-Item -LiteralPath $helper -Force -ErrorAction SilentlyContinue
    }
}

if (-not (Test-PgPort)) {
    $pgSvcs = @(Get-Service -ErrorAction SilentlyContinue | Where-Object {
        $_.Name -like "postgresql*" -or $_.DisplayName -like "*PostgreSQL*"
    })
    if ($pgSvcs.Count -eq 0) {
        Write-Warning "未检测到 Windows 版 PostgreSQL 服务（名称通常含 postgresql）。若用 Docker/WSL，请先手动启动 postgres 容器。"
    } else {
        foreach ($s in @($pgSvcs | Where-Object { $_.Status -ne "Running" })) {
            try {
                Start-Service -InputObject $s -ErrorAction Stop
                Write-Host "已启动 PostgreSQL 服务: $($s.Name)" -ForegroundColor Cyan
            } catch {
                Write-Host "当前会话无权启动 $($s.Name)，将尝试 UAC 提权…" -ForegroundColor DarkYellow
            }
        }
        Start-Sleep -Seconds 2
    }

    if (-not (Test-PgPort)) {
        $names = @($pgSvcs | ForEach-Object { $_.Name } | Sort-Object -Unique)
        $stillStopped = @()
        foreach ($n in $names) {
            $svc = Get-Service -Name $n -ErrorAction SilentlyContinue
            if ($svc -and $svc.Status -ne "Running") {
                $stillStopped += $svc
            }
        }
        if ($stillStopped.Count -gt 0) {
            Invoke-StartPostgreSqlServicesElevated
            Start-Sleep -Seconds 2
        }
    }

    if (-not (Test-PgPort)) {
        Write-Warning "127.0.0.1:5432 仍不可连（Connection refused 即服务端未监听）。请先确保 PostgreSQL 已运行，并已在库中创建 ppt_report_platform（见 .env.example）。"
    }
}

$venvPython = Join-Path $PSScriptRoot ".venv\Scripts\python.exe"
$appPy = Join-Path $PSScriptRoot "app.py"
if (-not (Test-Path -LiteralPath $appPy)) {
    Write-Error "未找到 app.py，请在项目根目录执行 start.ps1。"
}

Write-Host "Flask（前后端一体）: http://127.0.0.1:5000/ — Ctrl+C 停止" -ForegroundColor Green
if (Test-Path -LiteralPath $venvPython) {
    & $venvPython $appPy
} elseif (Get-Command py -ErrorAction SilentlyContinue) {
    & py -3 $appPy
} else {
    Write-Error "未检测到 Python：请在根目录创建 .venv，或安装 Python 后使用 py -3。"
}
