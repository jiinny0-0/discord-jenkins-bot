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
   | `DISCORD_NOTIFY_CHANNEL_ID` | Jenkins 화면에서 직접 실행한 Build의 알림 채널. 여러 개면 쉼표로 구분. 봇으로 실행한 Build는 실행한 채널로만 감 |
   | `BAT_RESULT_DIR` | 본 Job의 `DEST` 경로. 아티팩트가 없는 예전 Build의 결과를 찾을 때 씀 |
   | `CONNECTOR_ALLOWED_USER_IDS` | `/connector` 허용 사용자 ID (선택, 비우면 모두) |
   | `USASE_EXE`, `AICONNECTOR_EXE` | 커넥터 앱이 기본 경로가 아닐 때만 |

3. **로그인된 데스크톱 세션에서** 실행합니다. 커넥터 화면을 조작하므로 서비스나 Jenkins에서 돌리면 안 됩니다.
   ```powershell
   powershell -ExecutionPolicy Bypass -File C:\test-automation\discord-jenkins-bot\start-bot.ps1
   ```
4. Discord에서 `/ping` → `/bat-status` → `/connector status`로 확인합니다.

## 3. 업데이트

| 대상 | 방법 |
|---|---|
| BAT 코드 | update Job (`TARGET=bat`). 패키지가 크게 바뀌었으면 setup Job |
| 봇 코드 | update Job (`TARGET=bot` 또는 `all`) 후 **봇 재시작** (봇 창을 닫고 `start-bot.ps1` 다시 실행) |
| 비밀값·설정 | `C:\test-automation\.env.*`, `environments.yaml`, `bot-data\.env` 수정. 봇 `.env`는 봇 재시작 때 반영 |

봇은 반드시 한 번에 하나만 띄웁니다. 같은 토큰으로 둘이 돌면 명령과 알림이 두 번씩 처리됩니다.

## 4. 문제 해결

| 증상 | 확인 |
|---|---|
| 명령어가 Discord에 안 보임 | `DISCORD_GUILD_ID`, 초대 URL의 `applications.commands`. 봇 재시작 후 Discord 새로고침(Ctrl+R) |
| "Jenkins 조회 실패: HTTP 401/403" | `JENKINS_USER`, `JENKINS_API_TOKEN` |
| 완료 알림이 안 옴 | `DISCORD_NOTIFY_CHANNEL_ID`와 그 채널의 봇 권한 |
| 결과 파일을 찾지 못함 | Build에 `bat-artifacts` 아티팩트가 있는지. 예전 Build면 `BAT_RESULT_DIR` |
| 커넥터 준비 실패 | 봇이 로그인된 데스크톱 세션에서 도는지, `.env.<stage>` 계정, 화면 잠금. 기록은 `bot-data\connector-last.log` |
| `[실행 실패] doctor 점검 실패` | 콘솔 세션 만료. 커넥터를 끄고 `bat login --profile <stage>` |
| update Job이 "dubious ownership" | 스크립트의 git에 `-c safe.directory=*`가 붙어 있는지 |
| update Job이 "봇 저장소에 원격이 없습니다" | 봇 폴더에 `git remote add origin <URL>` |
