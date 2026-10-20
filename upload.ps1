# GitHub 저장소 만들기 + 파일 올리기 + 비밀값 등록 + 시험 실행 (한 번만 실행)
$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot
$repo = "alsgur035-cmd/dart-earnings-alert"

Write-Host "[1/4] 파일 묶는 중..." -ForegroundColor Cyan
if (-not (Test-Path .git)) { git init -q -b main }
git config user.name "alsgur035-cmd"
git config user.email "alsgur035-cmd@users.noreply.github.com"
git add -A
git commit -q -m "DART 실적 공시 텔레그램 알림"

Write-Host "[2/4] GitHub 공개 저장소 만들고 올리는 중..." -ForegroundColor Cyan
gh repo create $repo --public --source . --push --description "DART 실적 공시 텔레그램 알림"

Write-Host "[3/4] 비밀값 등록" -ForegroundColor Cyan
$envFile = Join-Path $PSScriptRoot "..\..\02_설정파일\.env"
$dart = (Get-Content $envFile -Encoding UTF8 | Where-Object { $_ -match '^DART_API_KEY=' }) -replace '^DART_API_KEY=', '' -replace '["\s]', ''
gh secret set DART_API_KEY --repo $repo --body $dart
Write-Host "  DART 키 등록 완료 (.env에서 읽음)"
gh secret set TELEGRAM_CHAT_ID --repo $repo --body "8450555165"
Write-Host "  채팅 ID 등록 완료"
Write-Host "  이제 텔레그램 봇 토큰을 붙여넣고 Enter (화면에 안 보여도 정상):" -ForegroundColor Yellow
gh secret set TELEGRAM_BOT_TOKEN --repo $repo

Write-Host "[4/4] 시험 실행 요청..." -ForegroundColor Cyan
Start-Sleep -Seconds 5
gh workflow run alert.yml --repo $repo
Write-Host ""
Write-Host "끝! 1~3분 안에 폰으로 알림이 오면 성공입니다." -ForegroundColor Green
Write-Host "진행 상황: https://github.com/$repo/actions"
