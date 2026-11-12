# 봇을 끄고(start-bot.ps1 재시도 루프 포함) 작업 스케줄러로 다시 띄운다.
# Jenkins(SYSTEM)나 관리자 PowerShell에서 실행:
#   powershell -NoProfile -ExecutionPolicy Bypass -File .\restart-bot.ps1             끄고 다시 켜기
#   powershell -NoProfile -ExecutionPolicy Bypass -File .\restart-bot.ps1 -StopOnly   끄기만 (패키지 설치 전 등)
#
# 커넥터 조작이나 BAT 테스트가 진행 중이면 끝날 때까지 기다린다 (최대 -WaitMinutes분, 넘으면 재시작하지 않고 실패).
# 다시 켜는 작업(register-autostart.ps1로 등록)은 봇 계정이 로그인돼 있어야 그 데스크톱에서 뜬다.
param(
    [switch]$StopOnly,
    [int]$WaitMinutes = 30,
    [string]$TaskName = "discord-jenkins-bot"
)
$ErrorActionPreference = "Stop"

function Write-Step([string]$message) {
    Write-Host ("{0} {1}" -f (Get-Date -Format "HH:mm:ss"), $message)
}

# 파일 이름은 경로 구분자·따옴표·공백 바로 뒤에 올 때만 맞춘다.
# ('restart-bot.ps1' 안에도 'start-bot.ps1'이 들어 있어서, 그냥 포함 여부로 찾으면 이 스크립트가 자기 자신을 끈다)
$BotScriptRe = '(?i)[\\/"\s]bot\.py\b'
$LoopScriptRe = '(?i)[\\/"\s]start-bot\.ps1\b'
$ConnectorRe = '(?i)[\\/"\s]connector\.py\b'

function Get-OtherProcesses {
    Get-CimInstance Win32_Process | Where-Object { $_.ProcessId -ne $PID -and $_.CommandLine }
}

function Get-BotProcesses {
    Get-OtherProcesses | Where-Object {
        ($_.Name -eq "python.exe" -and $_.CommandLine -match $BotScriptRe) -or
        ($_.Name -eq "powershell.exe" -and $_.CommandLine -match $LoopScriptRe)
    }
}

function Get-BusyReason {
    $procs = @(Get-OtherProcesses)
    if ($procs | Where-Object { $_.Name -eq "python.exe" -and $_.CommandLine -match $ConnectorRe }) { return "커넥터 조작 중" }
    if ($procs | Where-Object { $_.CommandLine -match "bat(\.exe)?.{0,2} run " }) { return "BAT 테스트 실행 중" }
    return $null
}

# 1) 진행 중인 작업이 끝날 때까지 대기
$deadline = (Get-Date).AddMinutes($WaitMinutes)
while ($reason = Get-BusyReason) {
    if ((Get-Date) -gt $deadline) {
        Write-Step "[실패] ${WaitMinutes}분 동안 '$reason' 상태가 끝나지 않아 재시작하지 않았습니다."
        exit 2
    }
    Write-Step "대기: $reason (끝나면 진행)"
    Start-Sleep -Seconds 30
}

# 2) 끄기: 재시도 루프(powershell)를 먼저 끝내야 python을 꺼도 다시 뜨지 않는다
$targets = @(Get-BotProcesses | Sort-Object { if ($_.Name -eq "powershell.exe") { 0 } else { 1 } })
if ($targets.Count -eq 0) {
    Write-Step "실행 중인 봇 없음"
} else {
    foreach ($p in $targets) {
        Write-Step "종료: $($p.Name) (PID $($p.ProcessId))"
        try { Stop-Process -Id $p.ProcessId -Force } catch { }
    }
    for ($i = 0; $i -lt 20 -and @(Get-BotProcesses).Count -gt 0; $i++) { Start-Sleep -Milliseconds 500 }
    if (@(Get-BotProcesses).Count -gt 0) {
        Write-Step "[실패] 봇 프로세스가 종료되지 않았습니다."
        exit 1
    }
    Write-Step "봇 종료 완료"
}
if ($StopOnly) { exit 0 }

# 3) 켜기
try {
    Start-ScheduledTask -TaskName $TaskName
} catch {
    Write-Step "[실패] 작업 '$TaskName'을(를) 시작하지 못했습니다: $($_.Exception.Message) (register-autostart.ps1로 등록했는지 확인)"
    exit 1
}
for ($i = 0; $i -lt 30; $i++) {
    Start-Sleep -Seconds 1
    $bot = Get-BotProcesses | Where-Object { $_.Name -eq "python.exe" }
    if ($bot) {
        Write-Step "봇 시작 완료 (PID $(@($bot)[0].ProcessId))"
        exit 0
    }
}
Write-Step "[실패] 30초 안에 봇이 뜨지 않았습니다. 봇 계정이 로그인돼 있는지, bot-data\bot-start.log를 확인하세요."
exit 1
