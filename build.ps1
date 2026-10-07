# ChurchPPTMaker 빌드 스크립트
#   .\build.ps1            가상환경(.venv) 준비 → 패키지 설치 → dist\ChurchPPTMaker.exe 생성
#   .\build.ps1 -Clean     build\, dist\ 를 지우고 새로 빌드
param([switch]$Clean)

$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot

$venvPy = Join-Path $PSScriptRoot '.venv\Scripts\python.exe'

if (-not (Test-Path $venvPy)) {
    Write-Host '[1/3] 가상환경(.venv) 생성' -ForegroundColor Cyan
    if (Get-Command py -ErrorAction SilentlyContinue) { py -3 -m venv .venv }
    elseif (Get-Command python -ErrorAction SilentlyContinue) { python -m venv .venv }
    else { throw 'Python 3.10 이상이 필요합니다. https://www.python.org 에서 설치하세요.' }
    if ($LASTEXITCODE -ne 0) { throw '가상환경 생성 실패' }
} else {
    Write-Host '[1/3] 가상환경(.venv) 사용' -ForegroundColor Cyan
}

Write-Host '[2/3] 패키지 설치' -ForegroundColor Cyan
& $venvPy -m pip install --disable-pip-version-check -q -r requirements-build.txt
if ($LASTEXITCODE -ne 0) { throw '패키지 설치 실패' }

if ($Clean) {
    Remove-Item -Recurse -Force build, dist -ErrorAction SilentlyContinue
}

Write-Host '[3/3] PyInstaller 빌드' -ForegroundColor Cyan
& $venvPy -m PyInstaller --noconfirm --clean --onefile --windowed `
    --name ChurchPPTMaker --icon assets/app.ico --add-data "assets/app.ico;assets" `
    --collect-data customtkinter main.py
if ($LASTEXITCODE -ne 0) { throw 'PyInstaller 빌드 실패' }

$exe = Join-Path $PSScriptRoot 'dist\ChurchPPTMaker.exe'
$sizeMB = [math]::Round((Get-Item $exe).Length / 1MB, 1)
Write-Host "완료: $exe ($sizeMB MB)" -ForegroundColor Green
