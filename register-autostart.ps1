# 로그온할 때 봇을 관리자 권한으로 자동 실행하도록 작업 스케줄러에 등록한다.
# 관리자 PowerShell에서, 봇을 띄울 계정(데스크톱에 로그인하는 계정)으로 실행:
#   powershell -ExecutionPolicy Bypass -File .\register-autostart.ps1            등록 (같은 이름이 있으면 덮어씀)
#   powershell -ExecutionPolicy Bypass -File .\register-autostart.ps1 -StartNow  등록 후 바로 실행
#   powershell -ExecutionPolicy Bypass -File .\register-autostart.ps1 -Remove    등록 해제
#
# 관리자 권한이 필요한 이유: U+SASE 앱이 관리자 권한으로만 실행되고(WinError 740),
# 일반 권한 프로그램은 관리자 권한 창을 조작할 수 없다.
param(
    [switch]$StartNow,
    [switch]$Remove,
    [string]$TaskName = "discord-jenkins-bot"
)
$ErrorActionPreference = "Stop"

$principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "[실패] 관리자 PowerShell에서 실행해 주세요."
    exit 1
}

if ($Remove) {
    if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
        Write-Host "[완료] 작업 '$TaskName' 등록을 해제했습니다."
    } else {
        Write-Host "작업 '$TaskName'이(가) 없습니다."
    }
    exit 0
}

$startScript = Join-Path $PSScriptRoot "start-bot.ps1"
if (-not (Test-Path $startScript)) {
    Write-Host "[실패] $startScript 가 없습니다."
    exit 1
}
$user = [Security.Principal.WindowsIdentity]::GetCurrent().Name

$action = New-ScheduledTaskAction -Execute "powershell.exe" `
    -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$startScript`"" `
    -WorkingDirectory $PSScriptRoot
# 이 계정이 로그온할 때 실행. 커넥터 화면을 조작해야 하므로 로그인된 데스크톱 세션(Interactive)에서 돈다.
$trigger = New-ScheduledTaskTrigger -AtLogOn -User $user
$taskPrincipal = New-ScheduledTaskPrincipal -UserId $user -LogonType Interactive -RunLevel Highest
# ExecutionTimeLimit 0: 기본값(72시간)이면 봇이 3일 뒤 강제 종료된다.
# IgnoreNew: 이미 실행 중이면 새로 띄우지 않는다 (봇이 둘이면 명령·알림이 두 번씩 처리됨).
$settings = New-ScheduledTaskSettingsSet -ExecutionTimeLimit ([TimeSpan]::Zero) `
    -MultipleInstances IgnoreNew -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries

Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger `
    -Principal $taskPrincipal -Settings $settings -Force | Out-Null
Write-Host "[완료] 작업 '$TaskName' 등록: $user 로그온 시 관리자 권한으로 $startScript 실행"

if ($StartNow) {
    Start-ScheduledTask -TaskName $TaskName
    Write-Host "[완료] 지금 실행했습니다. 'discord-jenkins-bot' 창을 확인해 주세요."
} else {
    Write-Host "바로 띄우려면: Start-ScheduledTask -TaskName $TaskName  (이미 띄운 봇이 있으면 먼저 닫을 것)"
}
