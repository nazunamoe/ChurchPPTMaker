# ChurchPPTMaker 빌드 스크립트
#   .\build.ps1                  이 PC 의 CPU 용 exe (x64 PC → x64, ARM64 PC → ARM64 — ARM64 Python 이 없으면 x64)
#   .\build.ps1 -Arch arm64      ARM64(Windows on Arm) 용 → dist\ChurchPPTMaker-arm64.exe (ARM64 PC + ARM64 Python 필요)
#   .\build.ps1 -Arch x64        x64 용 → dist\ChurchPPTMaker.exe (ARM64 PC 에서도 x64 Python 이 있으면 가능)
#   .\build.ps1 -Python <python.exe>   그 Python 으로 빌드 (CPU 는 그 Python 을 따름 — GitHub Actions 에서 사용)
#   .\build.ps1 -Clean           build\, dist\ 를 지우고 새로 빌드
#   .\build.ps1 -Version 1.2.0   version.py 의 버전을 1.2.0 으로 바꾼 뒤 빌드
# 버전은 version.py 하나에서 관리 → 화면 제목과 exe 파일 속성(파일/제품 버전)에 함께 반영된다.
# CPU: PyInstaller 는 다른 CPU 용 exe 를 만들지 못하므로(교차 빌드 없음) 그 CPU 용 Python 으로 CPU 별 가상환경
# (.venv = x64, .venv-arm64 = ARM64)을 만들어 빌드한다. x64 PC 에서는 ARM64 Python 을 실행할 수 없어 ARM64 빌드 불가 →
# GitHub Actions(.github\workflows\build.yml, windows-11-arm)로 만든다.
param(
    [switch]$Clean,
    [string]$Version,
    [ValidateSet('auto', 'x64', 'arm64')][string]$Arch = 'auto',
    [string]$Python
)

$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot

$Platforms = @{ x64 = 'win-amd64'; arm64 = 'win-arm64' }   # Python sysconfig.get_platform() 값
$MinPython = [version]'3.10'

function Get-OsArch {
    # PC 의 CPU. 에뮬레이션으로 실행된 프로세스(ARM64 PC 의 x64 PowerShell 등)는 PROCESSOR_ARCHITECTURE 가
    # 에뮬레이션 CPU 로 보이므로 시스템 환경 변수(레지스트리)를 읽는다
    $key = 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Environment'
    $v = (Get-ItemProperty -Path $key -Name PROCESSOR_ARCHITECTURE -ErrorAction SilentlyContinue).PROCESSOR_ARCHITECTURE
    if (-not $v) { $v = $env:PROCESSOR_ARCHITECTURE }
    if ($v -eq 'ARM64') { 'arm64' } else { 'x64' }
}

function Get-PythonInfo([string]$exe) {
    # 그 Python 의 CPU·버전·설치 폴더(base_prefix — 가상환경이면 만든 Python 의 폴더). tkinter 가 없으면(설치 때
    # tcl/tk 를 뺌) 쓸 수 없으므로 null. 표준 입력은 비워 둔다 — 'python' 별칭이 설치할지 묻느라 멈추지 않게
    if (-not $exe -or -not (Test-Path -LiteralPath $exe -PathType Leaf)) { return $null }
    $ErrorActionPreference = 'Continue'   # Windows PowerShell 5.1 은 Stop 이면 네이티브 명령의 stderr 출력만으로도 중단됨
    try {
        $out = @($null | & $exe -c 'import sys, sysconfig, tkinter; print(sysconfig.get_platform(), *sys.version_info[:2]); print(sys.base_prefix)' 2>$null)
    } catch { return $null }
    if ($LASTEXITCODE -ne 0 -or $out.Count -lt 2) { return $null }
    $f = "$($out[0])".Trim() -split '\s+'
    if ($f.Count -ne 3) { return $null }
    $cpu = $null
    foreach ($k in $Platforms.Keys) { if ($Platforms[$k] -eq $f[0]) { $cpu = $k } }
    [pscustomobject]@{ Path = $exe; Platform = $f[0]; Arch = $cpu; Version = [version]"$($f[1]).$($f[2])"
                       BasePrefix = "$($out[1])".Trim().TrimEnd('\') }
}

function Find-Python([string]$cpu) {
    # 설치된 Python(py 런처 목록 + PATH 의 python) 중 그 CPU 용 최신 버전
    $paths = New-Object System.Collections.Generic.List[string]
    if (Get-Command py -ErrorAction SilentlyContinue) {
        $ErrorActionPreference = 'Continue'
        foreach ($line in @(py -0p 2>$null)) {   # 예: " -V:3.13-arm64 *        C:\...\python.exe"
            if ("$line" -match '([A-Za-z]:\\.*\.exe)\s*$') { $paths.Add($Matches[1]) }
        }
    }
    foreach ($cmd in @(Get-Command python, python3 -CommandType Application -All -ErrorAction SilentlyContinue)) {
        $paths.Add($cmd.Source)
    }
    $best = $null
    foreach ($p in ($paths | Sort-Object -Unique)) {
        $info = Get-PythonInfo $p
        if ($info -and $info.Arch -eq $cpu -and $info.Version -ge $MinPython -and
            (-not $best -or $info.Version -gt $best.Version)) { $best = $info }
    }
    $best
}

function Get-VenvDir([string]$cpu) {
    if ($cpu -eq 'arm64') { Join-Path $PSScriptRoot '.venv-arm64' } else { Join-Path $PSScriptRoot '.venv' }
}

function Get-PeArch([string]$path) {
    # exe 의 CPU (PE 헤더의 Machine 값) — 다른 CPU 용 Python 으로 빌드되지 않았는지 마지막에 확인
    $buf = New-Object byte[] 4096
    $fs = [System.IO.File]::OpenRead($path)
    try { $n = $fs.Read($buf, 0, $buf.Length) } finally { $fs.Dispose() }
    if ($n -lt 64) { return 'unknown' }
    $pe = [BitConverter]::ToInt32($buf, 0x3C)
    if ($pe -lt 0 -or $pe + 6 -gt $n) { return 'unknown' }
    switch ([BitConverter]::ToUInt16($buf, $pe + 4)) {
        0x8664 { 'x64' }
        0xAA64 { 'arm64' }
        0x014C { 'x86' }
        default { 'unknown' }
    }
}

# ---- 어느 CPU 용으로, 어느 Python 으로 만들지
$osArch = Get-OsArch
$given = $null
if ($Python) {
    $given = Get-PythonInfo $Python
    if (-not $given -or -not $given.Arch -or $given.Version -lt $MinPython) {
        throw "쓸 수 없는 Python 입니다 (x64·ARM64 용 $MinPython 이상, tcl/tk 포함): $Python"
    }
    if ($Arch -ne 'auto' -and $Arch -ne $given.Arch) { throw "-Python 은 $($given.Arch) 용인데 -Arch 는 $Arch 입니다." }
    $candidates = @($given.Arch)
} elseif ($Arch -ne 'auto') {
    $candidates = @($Arch)
} elseif ($osArch -eq 'arm64') {
    $candidates = @('arm64', 'x64')   # ARM64 Python 이 없으면 x64 (Windows 11 on Arm 은 x64 exe 를 에뮬레이션으로 실행)
} else {
    $candidates = @('x64')
}

$target = $null
$create = $null   # 가상환경을 새로 만들 Python (null 이면 있는 가상환경 사용)
foreach ($cpu in $candidates) {
    $venvInfo = Get-PythonInfo (Join-Path (Get-VenvDir $cpu) 'Scripts\python.exe')
    # 예전 .venv 가 다른 CPU 용일 수 있음(ARM64 PC 에서 ARM64 Python 으로 만든 .venv 등) → 그 CPU 용이 아니면 다시 만듦
    $venvOk = $venvInfo -and $venvInfo.Arch -eq $cpu -and $venvInfo.Version -ge $MinPython -and
              (-not $given -or $venvInfo.BasePrefix -eq $given.BasePrefix)
    if ($venvOk) { $target = $cpu; break }
    $py = if ($given) { $given } else { Find-Python $cpu }
    if ($py) { $target = $cpu; $create = $py; break }
}

if (-not $target) {
    $want = $candidates[0]
    if ($want -eq 'arm64' -and $osArch -ne 'arm64') {
        throw ("ARM64 용 exe 는 ARM64 PC(Windows 11 on Arm)에서만 만들 수 있습니다. PyInstaller 는 다른 CPU 용으로 " +
               "빌드하지 못하고, x64 PC 에서는 ARM64 Python 을 실행할 수 없습니다.`n" +
               "GitHub 저장소의 [Actions] → [build] → [Run workflow] 로 x64·ARM64 exe 를 함께 만들 수 있습니다.")
    }
    $kinds = ($candidates | ForEach-Object { if ($_ -eq 'arm64') { "'Windows installer (ARM64)'" } else { "'Windows installer (64-bit)'" } }) -join ' 또는 '
    throw ("$($candidates -join '·') 용 Python $MinPython 이상을 찾지 못했습니다. " +
           "https://www.python.org/downloads/windows/ 에서 $kinds 를 받아 설치하세요 (tcl/tk 는 기본 설치에 들어 있음).")
}

$suffix = if ($target -eq 'arm64') { '-arm64' } else { '' }
$appName = "ChurchPPTMaker$suffix"
$venvDir = Get-VenvDir $target
$venvPy = Join-Path $venvDir 'Scripts\python.exe'

if ($create) {
    Write-Host "[1/3] 가상환경($(Split-Path -Leaf $venvDir)) 생성 — $target, Python $($create.Version)" -ForegroundColor Cyan
    Write-Host "      $($create.Path)"
    & $create.Path -m venv --clear $venvDir   # --clear: 다른 CPU·Python 으로 만든 예전 가상환경이 있으면 비우고 새로
    if ($LASTEXITCODE -ne 0) { throw '가상환경 생성 실패' }
} else {
    Write-Host "[1/3] 가상환경($(Split-Path -Leaf $venvDir)) 사용 — $target" -ForegroundColor Cyan
}
if ($Arch -eq 'auto' -and -not $Python -and $target -ne $osArch) {
    Write-Host ("      ARM64 Python 이 없어 x64 용으로 만듭니다 (Windows 11 on Arm 에서는 에뮬레이션으로 실행됨). " +
                "ARM64 용은 python.org 의 'Windows installer (ARM64)' 를 설치한 뒤 다시 빌드하세요.") -ForegroundColor Yellow
}

Write-Host '[2/3] 패키지 설치' -ForegroundColor Cyan
& $venvPy -m pip install --disable-pip-version-check -q -r requirements-build.txt
if ($LASTEXITCODE -ne 0) { throw '패키지 설치 실패' }

if ($Clean) {
    Remove-Item -Recurse -Force build, dist -ErrorAction SilentlyContinue
}

if ($Version) {
    & $venvPy version.py --set $Version
    if ($LASTEXITCODE -ne 0) { throw '버전 변경 실패' }
}
$versionInfo = Join-Path $PSScriptRoot "build\version_info$suffix.txt"
$appVersion = (& $venvPy version.py --write-info $versionInfo $appName)
if ($LASTEXITCODE -ne 0) { throw '버전 정보 생성 실패' }

Write-Host "[3/3] PyInstaller 빌드 (v$appVersion, $target)" -ForegroundColor Cyan
& $venvPy -m PyInstaller --noconfirm --clean --onefile --windowed `
    --name $appName --icon assets/app.ico --add-data "assets/app.ico;assets" `
    --version-file $versionInfo --collect-data customtkinter main.py
if ($LASTEXITCODE -ne 0) { throw 'PyInstaller 빌드 실패' }

$exe = Join-Path $PSScriptRoot "dist\$appName.exe"
$exeArch = Get-PeArch $exe
if ($exeArch -ne $target) { throw "만든 exe 가 $target 용이 아닙니다 ($exeArch): $exe" }
$sizeMB = [math]::Round((Get-Item $exe).Length / 1MB, 1)
Write-Host "완료: $exe (v$appVersion, $target, $sizeMB MB)" -ForegroundColor Green
