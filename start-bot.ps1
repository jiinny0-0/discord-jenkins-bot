# 봇 실행: 저장소 밖의 .env 원본을 가져온 뒤 bot.py를 띄운다.
#   powershell -ExecutionPolicy Bypass -File .\start-bot.ps1
$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

$dataDir = if ($env:BOT_DATA_DIR) { $env:BOT_DATA_DIR } else { "C:\test-automation\bot-data" }
$envSource = Join-Path $dataDir ".env"

if (-not (Test-Path $envSource)) {
    Write-Host "[실행 실패] $envSource 가 없습니다. env.example을 참고해 만들어 주세요."
    exit 1
}
Copy-Item $envSource ".env" -Force

if (-not (Test-Path ".venv\Scripts\python.exe")) {
    Write-Host "venv가 없어 새로 만듭니다."
    python -m venv .venv
    .\.venv\Scripts\python.exe -m pip install -r requirements.txt
}

$Host.UI.RawUI.WindowTitle = "discord-jenkins-bot"
.\.venv\Scripts\python.exe .\bot.py
