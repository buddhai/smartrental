# Streamlit Cloud 배포 보조 스크립트
# 사용: powershell -ExecutionPolicy Bypass -File deploy_github.ps1

$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

function Resolve-GhExe {
    $cmd = Get-Command gh -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }

    $candidates = @(
        "$env:ProgramFiles\GitHub CLI\gh.exe",
        "${env:ProgramFiles(x86)}\GitHub CLI\gh.exe",
        "$env:LOCALAPPDATA\Programs\GitHub CLI\gh.exe"
    )
    foreach ($p in $candidates) {
        if (Test-Path $p) {
            $dir = Split-Path $p -Parent
            if ($env:Path -notlike "*$dir*") {
                $env:Path = "$env:Path;$dir"
            }
            return $p
        }
    }
    return $null
}

$Gh = Resolve-GhExe
if (-not $Gh) {
    Write-Host "GitHub CLI(gh)를 찾을 수 없습니다."
    Write-Host "설치: winget install GitHub.cli"
    Write-Host "설치 후 PowerShell을 새로 열거나 아래를 실행하세요:"
    Write-Host '  $env:Path += ";C:\Program Files\GitHub CLI"'
    exit 1
}

Write-Host "gh 경로: $Gh"

$auth = & $Gh auth status 2>&1
if ($LASTEXITCODE -ne 0) {
    Write-Host "GitHub 로그인이 필요합니다. 아래 명령을 실행하세요:"
    Write-Host "  & `"$Gh`" auth login"
    exit 1
}

$repoName = Read-Host "GitHub 저장소 이름 (예: smartquote-rental)"
if (-not $repoName) { $repoName = "smartquote-rental" }

$visibility = Read-Host "공개 여부 [private/public] (기본: private)"
if (-not $visibility) { $visibility = "private" }

if ($visibility -eq "public") {
    & $Gh repo create $repoName --public --source=. --remote=origin --push
} else {
    & $Gh repo create $repoName --private --source=. --remote=origin --push
}

if ($LASTEXITCODE -eq 0) {
    $url = & $Gh repo view --json url -q .url
    Write-Host ""
    Write-Host "GitHub 푸시 완료: $url"
    Write-Host ""
    Write-Host "다음: https://share.streamlit.io 에서"
    Write-Host "  Main file path = streamlit_app.py"
    Write-Host "  Branch = main"
} else {
    Write-Host "저장소 생성 실패. 이미 remote가 있으면:"
    Write-Host "  git push -u origin main"
}
