# 봇 실행: 저장소 밖의 .env 원본을 가져온 뒤 bot.py를 띄운다.
# 봇이 꺼지면(부팅 직후 네트워크·Jenkins 미준비, 일시 장애 등) 30초 뒤 다시 띄운다. 완전히 끄려면 이 창을 닫는다.
#   powershell -ExecutionPolicy Bypass -File .\start-bot.ps1
# 기록: bot-data\bot-start.log (시작·종료 시각과 종료 코드), bot-data\bot.log (봇 로그, 꺼진 이유)
$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

$dataDir = if ($env:BOT_DATA_DIR) { $env:BOT_DATA_DIR } else { "C:\test-automation\bot-data" }
$envSource = Join-Path $dataDir ".env"
$startLog = Join-Path $dataDir "bot-start.log"
$retrySeconds = 30

function Write-StartLog([string]$message) {
    $line = "{0} {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $message
    Write-Host $line
    try { Add-Content -Path $startLog -Value $line -Encoding UTF8 } catch { }
}

if (-not (Test-Path $envSource)) {
    Write-Host "[실행 실패] $envSource 가 없습니다. env.example을 참고해 만들어 주세요."
    exit 1
}

if (-not (Test-Path ".venv\Scripts\python.exe")) {
    Write-Host "venv가 없어 새로 만듭니다."
    python -m venv .venv
    .\.venv\Scripts\python.exe -m pip install -r requirements.txt
}

$Host.UI.RawUI.WindowTitle = "discord-jenkins-bot"
while ($true) {
    try {
        Copy-Item $envSource ".env" -Force  # 다시 띄울 때 bot-data\.env 수정분도 반영
    } catch {
        Write-StartLog "[경고] .env 복사 실패, 기존 .env로 시작: $($_.Exception.Message)"
    }
    Write-StartLog "봇 시작"
    & .\.venv\Scripts\python.exe .\bot.py
    $code = $LASTEXITCODE
    Write-StartLog "봇 종료 (종료 코드 $code, 원인은 bot.log). ${retrySeconds}초 뒤 다시 시작합니다. 완전히 끄려면 이 창을 닫으세요."
    Start-Sleep -Seconds $retrySeconds
}
