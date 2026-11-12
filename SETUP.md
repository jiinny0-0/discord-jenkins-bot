# discord-jenkins-bot 설치

Discord 슬래시 명령으로 Jenkins의 BAT Job(`aioncloud-sse-bat`)을 실행하고 결과를 채널로 받는 봇입니다.
테스트 전에 커넥터(AIConnector, U+SASE)를 해당 스테이지로 로그인시키는 일도 봇이 합니다.

## 1. 폴더 구조

```
C:\test-automation\
├─ .env                     공용 BAT 설정 (비어 있어도 됨. setup Job이 있는지만 검사)
├─ .env.<stage>             스테이지별 콘솔 계정·Open API 토큰 (qa, real, test, lguplus, lg-preprod, lg-staging)
├─ environments.yaml        스테이지별 URL·워크스페이스·단말
├─ bat-exclude.json         all 실행 제외 목록 (봇이 만듦)
├─ ms-playwright\           BAT용 브라우저 (setup Job이 설치)
├─ bot-data\                봇 설정 원본(.env)과 상태 파일 (봇이 만듦)
├─ aioncloud-sse-bat\       BAT 저장소
└─ discord-jenkins-bot\     이 저장소
```

git으로 받는 것은 두 저장소 폴더뿐입니다. 비밀값과 상태 파일은 저장소 밖에 있어서 업데이트해도 지워지지 않습니다.
설정은 항상 `C:\test-automation\` 쪽 원본을 고칩니다. 본 Job이 실행할 때마다 BAT 폴더 안으로 복사합니다.

## 2. 새 서버에 설치

### 2-1. 준비물

- Python(**시스템 PATH**에 있어야 함. Jenkins가 LocalSystem으로 돌기 때문), Git, Jenkins(Windows 서비스), Chrome, 커넥터 앱
- 기존 서버가 있으면 `C:\test-automation\`의 `.env*`, `environments.yaml`, `bot-data\`를 그대로 복사합니다.

### 2-2. Jenkins

1. 관리자 계정의 **API Token**을 발급합니다 (사용자 → Security → API Token). 봇이 이 계정으로 Job을 실행합니다.
2. 기존 서버의 `C:\ProgramData\Jenkins\.jenkins\jobs\`에서 Job 폴더 3개를 복사하고 Jenkins를 재시작합니다.

   | Job | 하는 일 | 파라미터 |
   |---|---|---|
   | `aioncloud-sse-bat` | 설정 복사 → doctor → sweep → `bat run` → 결과 CSV·`run-report.json`을 아티팩트로 보관 | `STAGE`, `PRODUCT`, `MODE`(run/doctor/sweep), `PRODUCTS`, `OVERRIDE_ID`, `LANG`(en/ko) |
   | `setup` | BAT 새로 clone → venv → `pip install -e .` → Playwright 브라우저 | `RECREATE_VENV` |
   | `update` | `fetch` + `reset --hard`로 코드 갱신 (`reports`·`.venv`는 유지) | `TARGET`(bat/bot/all, 기본 bat) |

   봇은 `MODE`·`PRODUCTS`가 없으면 doctor/sweep/사전 설정 실행을, `OVERRIDE_ID`가 없으면 `/bat custom`을 거절합니다.
3. **setup Job**을 실행합니다. 브라우저 설치가 실패하면 커넥터를 잠시 끄고 다시 실행합니다.
4. 커넥터를 끈 상태에서 콘솔 세션을 받고, 본 Job을 `MODE=doctor`로 실행해 확인합니다.
   ```powershell
   cd C:\test-automation\aioncloud-sse-bat; .\.venv\Scripts\Activate.ps1; bat login --profile qa
   ```

### 2-3. Discord 앱

[Developer Portal](https://discord.com/developers/applications)에서:

1. New Application → **Bot** 탭에서 토큰 발급 → `DISCORD_BOT_TOKEN`
2. **OAuth2 → URL Generator**: scope `bot`, `applications.commands` / 권한 Send Messages, Embed Links, Attach Files, Read Message History → 생성된 URL로 서버에 초대
3. Discord 개발자 모드를 켜고 ID를 복사합니다 (우클릭 → ID 복사).

### 2-4. 봇

1. `git clone <이 저장소 URL> C:\test-automation\discord-jenkins-bot`
2. `env.example`을 `C:\test-automation\bot-data\.env`로 복사해 채웁니다. 봇 폴더의 `.env`는 시작할 때마다 이 파일로 덮어씁니다.

   | 키 | 값 |
   |---|---|
   | `JENKINS_URL`, `JENKINS_JOB` | 예: `http://localhost:8080`, `aioncloud-sse-bat` |
   | `JENKINS_USER`, `JENKINS_API_TOKEN` | 2-2의 계정·토큰 |
   | `JENKINS_LINK_URL` | `/jenkins` 명령이 보여줄 주소 (선택) |
   | `DISCORD_BOT_TOKEN` | 2-3 |
   | `DISCORD_GUILD_ID` | 서버 ID. 여러 개면 쉼표로 구분 (실행은 서버와 상관없이 한 번에 하나) |
   | `BAT_RESULT_DIR` | 본 Job의 `DEST` 경로. 아티팩트가 없는 예전 Build의 결과를 찾을 때 씀 |
   | `CONNECTOR_ALLOWED_USER_IDS` | `/connector` 허용 사용자 ID (선택, 비우면 모두) |
   | `USASE_EXE`, `AICONNECTOR_EXE` | 커넥터 앱이 기본 경로가 아닐 때만 |

3. **로그인된 데스크톱 세션에서, 관리자 권한으로** 실행합니다. 커넥터 화면을 조작하므로 서비스나 Jenkins에서 돌리면 안 됩니다.
   U+SASE 앱이 관리자 권한으로만 실행되기 때문에, 봇이 일반 권한이면 `/connector on`이 `WinError 740`으로 실패합니다.

   자동 실행 등록 (권장, 한 번만): 관리자 PowerShell에서, 봇을 띄울 계정으로 실행합니다.
   ```powershell
   powershell -ExecutionPolicy Bypass -File C:\test-automation\discord-jenkins-bot\register-autostart.ps1 -StartNow
   ```
   그 계정이 로그온할 때마다 봇이 관리자 권한으로 뜹니다 (작업 이름 `discord-jenkins-bot`). 해제는 `-Remove`.
   재부팅 후 아무도 로그인하지 않으면 뜨지 않으므로, 무인 운영이면 서버 자동 로그인도 설정합니다.

   자동 로그인: 재부팅 직후에는 로그인 세션이 없어서 봇이 뜰 수 없으므로 세션을 하나 만들어 줍니다.
   로그인 뒤 화면이 잠기거나 RDP 연결이 끊겨도 커넥터 조작은 동작하므로 잠금 설정은 바꾸지 않아도 됩니다.
   Microsoft Sysinternals의 [Autologon](https://learn.microsoft.com/sysinternals/downloads/autologon)을 씁니다 (설치 없음, zip을 받아 실행).
   비밀번호를 레지스트리에 평문으로 두지 않고 암호화(LSA secret)해서 저장합니다.
   1. `Autologon64.exe`를 관리자 권한으로 실행 → Username(봇 계정), Domain(로컬 계정이면 컴퓨터 이름), Password 입력 → **Enable**
   2. 서버를 재부팅해 자동 로그인 → 봇 창이 뜨는지 확인합니다.
   해제는 Autologon에서 **Disable**.

   직접 실행: 관리자 PowerShell에서
   ```powershell
   powershell -ExecutionPolicy Bypass -File C:\test-automation\discord-jenkins-bot\start-bot.ps1
   ```
4. Discord에서 `/ping` → `/bat-status` → `/connector status`로 확인합니다.

## 3. 업데이트

| 대상 | 방법 |
|---|---|
| BAT 코드 | update Job (`TARGET=bat`). 패키지가 크게 바뀌었으면 setup Job |
| 봇 코드 | update Job (`TARGET=bot` 또는 `all`) 후 **봇 재시작** (아래) |
| 비밀값·설정 | `C:\test-automation\.env.*`, `environments.yaml`, `bot-data\.env` 수정. 봇 `.env`는 봇 재시작 때 반영 |

**봇 재시작** (관리자 PowerShell): 커넥터 조작·BAT 테스트가 진행 중이면 끝날 때까지 기다린 뒤 끄고 다시 켭니다.
```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File C:\test-automation\discord-jenkins-bot\restart-bot.ps1
```
- 끄기만: 같은 명령에 `-StopOnly`. 켜기만: `Start-ScheduledTask -TaskName discord-jenkins-bot`
- `start-bot.ps1`은 봇이 오류로 꺼지면 30초 뒤 다시 띄웁니다. 완전히 끄려면 봇 창을 닫거나 `-StopOnly`.
- 시작·종료 기록은 `bot-data\bot-start.log`, 봇 로그(꺼진 이유 포함)는 `bot-data\bot.log`에 남습니다.

### 3-1. 봇 자동 배포 (GitHub push → 1분 안에 서버 반영)

Jenkins에 Job `bot-auto-update`를 만듭니다 (새로운 Item → Freestyle project).

| 설정 | 값 |
|---|---|
| 소스 코드 관리 | Git. Repository URL `https://github.com/jiinny0-0/discord-jenkins-bot.git`, Branch `*/master` |
| 빌드 유발 | **Poll SCM**, Schedule `* * * * *` (1분마다 새 커밋 확인) |
| Build Steps | Execute Windows batch command: 아래 |

```bat
@echo off
chcp 65001 > nul
rem safe.directory는 환경변수로 넘긴다 (for /f 안에서는 -c 옵션의 '='가 깨짐)
set GIT_CONFIG_COUNT=1
set GIT_CONFIG_KEY_0=safe.directory
set GIT_CONFIG_VALUE_0=*
set BOT_DIR=C:\test-automation\discord-jenkins-bot
set BOT_BRANCH=master
set RUN_PS=powershell -NoProfile -ExecutionPolicy Bypass -File

for /f %%r in ('git -C "%BOT_DIR%" rev-parse HEAD') do set BOT_OLD=%%r
git -C "%BOT_DIR%" fetch origin %BOT_BRANCH% || exit /b 1
git -C "%BOT_DIR%" reset --hard origin/%BOT_BRANCH% || exit /b 1
git -C "%BOT_DIR%" log -1 --oneline

rem 봇 끄기 (커넥터 조작·BAT 테스트 중이면 끝날 때까지 대기)
%RUN_PS% "%BOT_DIR%\restart-bot.ps1" -StopOnly || exit /b 1

git -C "%BOT_DIR%" diff --quiet %BOT_OLD% HEAD -- requirements.txt
if errorlevel 1 (
    echo requirements.txt 변경됨: 패키지 설치
    "%BOT_DIR%\.venv\Scripts\python.exe" -m pip install -q -r "%BOT_DIR%\requirements.txt"
    if errorlevel 1 echo [WARN] 패키지 설치 실패
)

rem 봇 켜기 (작업 스케줄러)
%RUN_PS% "%BOT_DIR%\restart-bot.ps1"
exit /b %ERRORLEVEL%
```

- 봇 계정이 로그인돼 있어야 봇이 그 데스크톱에서 뜹니다 (자동 로그인 설정, 2-4).
- 봇 폴더에서 직접 고친 코드는 `reset --hard`로 지워집니다. 고칠 때는 GitHub에 push합니다.

봇은 반드시 한 번에 하나만 띄웁니다. 같은 토큰으로 둘이 돌면 명령과 알림이 두 번씩 처리됩니다.

## 4. 문제 해결

| 증상 | 확인 |
|---|---|
| 명령어가 Discord에 안 보임 | `DISCORD_GUILD_ID`, 초대 URL의 `applications.commands`. 봇 재시작 후 Discord 새로고침(Ctrl+R) |
| "Jenkins 조회 실패: HTTP 401/403" | `JENKINS_USER`, `JENKINS_API_TOKEN` |
| 완료 알림이 안 옴 | 봇으로 실행한 Build만 실행한 채널로 알림이 감 (Jenkins 화면 실행은 알림 없음, `/bat-result`로 확인). 그 채널의 봇 권한 |
| 결과 파일을 찾지 못함 | Build에 `bat-artifacts` 아티팩트가 있는지. 예전 Build면 `BAT_RESULT_DIR` |
| 커넥터 준비 실패 | 봇이 로그인된 데스크톱 세션에서 도는지, `.env.<stage>` 계정, 화면 잠금. 기록은 `bot-data\connector-last.log` |
| `[실행 실패] doctor 점검 실패` | 콘솔 세션 만료. 커넥터를 끄고 `bat login --profile <stage>` |
| update Job이 "dubious ownership" | 스크립트의 git에 `-c safe.directory=*`가 붙어 있는지 |
| update Job이 "봇 저장소에 원격이 없습니다" | 봇 폴더에 `git remote add origin <URL>` |
