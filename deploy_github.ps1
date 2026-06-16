# Streamlit Cloud 배포 보조 스크립트
# 사용: powershell -ExecutionPolicy Bypass -File deploy_github.ps1

$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

$gh = Get-Command gh -ErrorAction SilentlyContinue
if (-not $gh) {
    Write-Host "GitHub CLI(gh)가 없습니다. winget install GitHub.cli 후 다시 실행하세요."
    exit 1
}

$auth = gh auth status 2>&1
if ($LASTEXITCODE -ne 0) {
    Write-Host "GitHub 로그인이 필요합니다. 아래 명령을 실행하세요:"
    Write-Host "  gh auth login"
    exit 1
}

$repoName = Read-Host "GitHub 저장소 이름 (예: smartquote-rental)"
if (-not $repoName) { $repoName = "smartquote-rental" }

$visibility = Read-Host "공개 여부 [private/public] (기본: private)"
if (-not $visibility) { $visibility = "private" }

if ($visibility -eq "public") {
    gh repo create $repoName --public --source=. --remote=origin --push
} else {
    gh repo create $repoName --private --source=. --remote=origin --push
}

if ($LASTEXITCODE -eq 0) {
    $url = gh repo view --json url -q .url
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
