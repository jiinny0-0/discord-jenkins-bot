import sys
from dotenv import dotenv_values, load_dotenv
from pathlib import Path

load_dotenv(Path(__file__).resolve().parent / ".env")

import asyncio
import csv
import base64
import io
import json
import logging
import os
import re
import ssl
import statistics
import time
import threading
from datetime import datetime, timedelta
from urllib.parse import quote, urljoin

import aiohttp
import discord
from discord import app_commands

DISCORD_BOT_TOKEN = os.environ["DISCORD_BOT_TOKEN"]
# 명령어를 등록할 Discord 서버 ID. 여러 서버면 쉼표로 구분 (예: 111,222). 실행 큐·Jenkins는 함께 쓴다.
DISCORD_GUILD_IDS = [int(x) for x in os.environ["DISCORD_GUILD_ID"].replace(" ", "").split(",") if x]
if not DISCORD_GUILD_IDS:
    raise RuntimeError("DISCORD_GUILD_ID가 비어 있습니다")
# DISCORD_NOTIFY_CHANNEL_ID는 더 이상 쓰지 않는다. Jenkins 화면에서 실행한 Build는 자동 알림이 없고
# (결과는 /bat-result로 확인), 봇으로 실행한 Build는 실행한 채널로만 알림이 간다.

JENKINS_URL = os.environ["JENKINS_URL"].rstrip("/")
JENKINS_JOB = os.environ["JENKINS_JOB"]
JENKINS_USER = os.environ["JENKINS_USER"]
JENKINS_API_TOKEN = os.environ["JENKINS_API_TOKEN"]

# Build Step이 결과 CSV를 복사하는 폴더 (Build Step의 DEST와 같은 절대 경로)
BAT_RESULT_DIR = os.environ.get("BAT_RESULT_DIR", "").strip()

# all 실행 제외 목록 파일 (bat run이 실행 시작 때 읽음)
BAT_EXCLUDE_FILE = Path(
    os.environ.get("BAT_EXCLUDE_FILE", "").strip() or r"C:\test-automation\bat-exclude.json"
)
CASE_ID_PATTERN = re.compile(r"^BAT-[A-Z]+-\d{3}$")

# 봇이 고쳐 쓰는 상태 파일 폴더 (저장소 밖. 다시 clone해도 예약·사전 설정이 남도록)
BOT_DATA_DIR = Path(os.environ.get("BOT_DATA_DIR", "").strip() or r"C:\test-automation\bot-data")
BOT_DATA_DIR.mkdir(parents=True, exist_ok=True)

JOB_PATH = "/job/" + quote(JENKINS_JOB, safe="")
JOB_API_URL = f"{JENKINS_URL}{JOB_PATH}/api/json"
QUEUE_API_URL = f"{JENKINS_URL}/queue/api/json"
BUILD_API_TREE = (
    "api/json?tree=number,building,result,url,timestamp,duration,"
    "actions[parameters[name,value]]"
)

STAGES = ["qa", "real", "test", "lguplus", "lg-preprod", "lg-staging"]
PRODUCTS = [
    "all", "basic", "casb", "cross", "dlp",
    "fwaas", "genai", "rbi", "swg", "ztna",
]

MAX_CSV_ATTACHMENTS = 10
MAX_CSV_SIZE_BYTES = 8 * 1024 * 1024

# 완료 감시 설정
WATCH_INTERVAL_SECONDS = 15
CATCHUP_LIMIT = 5  # 봇이 꺼져 있던 동안 끝난 Build는 최근 5개까지만 알림
STATE_FILE = BOT_DATA_DIR / "notify-state.json"

# 콘솔 로그 패턴
TIMING_PATTERN = re.compile(r"\[BAT-TIMING\]\s+(START|END)=(\S+)")
EXCLUDE_LINE_PATTERN = re.compile(r"\[EXCLUDE\] items=(\S+)")
COPY_LINE_PATTERN = re.compile(r"복사:\s*(.+?\.csv)\s*$", re.MULTILINE | re.IGNORECASE)

def _log_handlers() -> list[logging.Handler]:
    """콘솔 + 파일(bot-data\\bot.log, 5MB×3개 순환). 봇 창이 닫혀도 꺼진 이유가 남도록."""
    from logging.handlers import RotatingFileHandler

    handlers: list[logging.Handler] = [logging.StreamHandler()]
    try:
        handlers.append(RotatingFileHandler(BOT_DATA_DIR / "bot.log", maxBytes=5 * 1024 * 1024,
                                            backupCount=3, encoding="utf-8"))
    except OSError:
        pass  # 파일을 못 열어도 봇은 띄운다
    return handlers


logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
    handlers=_log_handlers(),
)
log = logging.getLogger(__name__)

try:
    from csv_to_xlsx import csv_to_xlsx  # 같은 폴더의 csv_to_xlsx.py (openpyxl 필요)
except ImportError as exc:
    csv_to_xlsx = None
    log.warning("엑셀 변환 모듈을 불러오지 못해 CSV를 그대로 첨부합니다 (원인: %s)", exc)


class _HideVoiceWarning(logging.Filter):
    """음성 기능을 쓰지 않으므로 PyNaCl/davey 미설치 경고만 숨긴다."""
    def filter(self, record: logging.LogRecord) -> bool:
        return "voice will NOT be supported" not in record.getMessage()


logging.getLogger("discord.client").addFilter(_HideVoiceWarning())


# ============================================================
# 언어(en/ko): 응답 언어는 lang 옵션 > 채널 설정(/set) > en 순서로 정한다.
# 문구는 tr(lang, 한국어, 영어)로 두 언어를 나란히 적는다.
# ============================================================
LANGS = ("en", "ko")
DEFAULT_LANG = "en"
CHANNEL_LANG_FILE = BOT_DATA_DIR / "channel-lang.json"
channel_langs: dict[str, str] | None = None  # {채널 ID: 언어}. 처음 쓸 때 읽는다


def tr(lang: str | None, ko: str, en: str, **kw) -> str:
    """lang에 맞는 문구. kw가 있으면 format으로 채운다."""
    text = ko if lang == "ko" else en
    return text.format(**kw) if kw else text


def norm_lang(value) -> str | None:
    """Choice·문자열·None 어느 것이든 'en'/'ko' 또는 None."""
    value = getattr(value, "value", value)
    return value if value in LANGS else None


def load_channel_langs() -> dict[str, str]:
    try:
        data = json.loads(CHANNEL_LANG_FILE.read_text(encoding="utf-8-sig"))
        if isinstance(data, dict):
            return {str(k): v for k, v in data.items() if v in LANGS}
    except FileNotFoundError:
        pass
    except Exception:
        log.exception("채널 언어 파일을 읽지 못함 (기본 언어 사용)")
    return {}


def save_channel_langs() -> None:
    tmp = CHANNEL_LANG_FILE.with_suffix(".tmp")
    tmp.write_text(json.dumps(channel_langs or {}, indent=2), encoding="utf-8")
    tmp.replace(CHANNEL_LANG_FILE)


def channel_lang(channel_id) -> str:
    """채널에 /set으로 저장한 언어. 없으면 en."""
    global channel_langs
    if channel_langs is None:
        channel_langs = load_channel_langs()
    if not channel_id:
        return DEFAULT_LANG
    return channel_langs.get(str(channel_id)) or DEFAULT_LANG


def set_channel_lang(channel_id, lang: str) -> None:
    channel_lang(None)  # 아직 안 읽었으면 읽어 둔다
    channel_langs[str(channel_id)] = lang
    save_channel_langs()


def lang_for(interaction, explicit=None) -> str:
    """응답 언어: 명령의 lang 옵션 > 채널 설정 > en."""
    return norm_lang(explicit) or channel_lang(getattr(interaction, "channel_id", None))


def req_lang(req: dict) -> str:
    """예약·대기열 항목의 언어: 요청 때 정한 언어 > 채널 설정 > en."""
    return norm_lang(req.get("lang")) or channel_lang(req.get("channel_id"))


def L(en: str, ko: str) -> app_commands.locale_str:
    """명령어 목록에 보이는 설명: 기본은 영어, 한국어 Discord에는 ko."""
    return app_commands.locale_str(en, ko=ko)


class BotTranslator(app_commands.Translator):
    async def translate(self, string: app_commands.locale_str, locale: discord.Locale, context):
        if locale == discord.Locale.korean:
            return string.extras.get("ko")
        return None


LANG_CHOICES = [app_commands.Choice(name="English", value="en"), app_commands.Choice(name="한국어", value="ko")]
LANG_DESCRIBE = L("Reply language (default: this channel's /set language, else English)",
                  "응답 언어 (비우면 이 채널의 /set 언어, 없으면 영어)")


def with_lang(func):
    """명령 함수에 선택 옵션 lang(en/ko) 설명·선택지를 붙인다."""
    return app_commands.choices(lang=LANG_CHOICES)(app_commands.describe(lang=LANG_DESCRIBE)(func))

intents = discord.Intents.default()
TEST_GUILDS = [discord.Object(id=guild_id) for guild_id in DISCORD_GUILD_IDS]

http: aiohttp.ClientSession | None = None
bat_lock = asyncio.Lock()
exclude_lock = asyncio.Lock()
monitor_task: asyncio.Task | None = None
watch_task: asyncio.Task | None = None
state: dict = {}  # notify-state.json 내용 (Build 번호와 채널 ID만, 비밀 정보 없음)


class JenkinsDiscordClient(discord.Client):
    async def setup_hook(self):
        global http, watch_task, state
        timeout = aiohttp.ClientTimeout(total=30)
        # aiohttp 버전과 관계없이 동작하도록 Basic 인증 헤더를 직접 만든다 (값은 로그에 남기지 않음)
        credentials = base64.b64encode(
            f"{JENKINS_USER}:{JENKINS_API_TOKEN}".encode("utf-8")
        ).decode("ascii")
        http = aiohttp.ClientSession(
            timeout=timeout,
            headers={"Authorization": f"Basic {credentials}"},
        )
        await tree.set_translator(BotTranslator())  # 명령어 설명의 한국어 번역 (sync 전에)
        for guild in TEST_GUILDS:
            try:
                await tree.sync(guild=guild)
                log.info("슬래시 명령어 동기화 완료: 서버 %s", guild.id)
            except discord.HTTPException:
                # 봇이 초대되지 않은 서버 등. 다른 서버는 계속 등록한다.
                log.exception("슬래시 명령어 동기화 실패: 서버 %s (봇 초대·권한 확인)", guild.id)
        state = load_state()
        watch_task = asyncio.create_task(watch_builds())
        global schedule, schedule_task
        schedule = load_schedule()
        schedule_task = asyncio.create_task(schedule_loop())


client = JenkinsDiscordClient(intents=intents)
class GuardedTree(app_commands.CommandTree):
    """업데이트 차단(block) 중에는 차단을 건 사람 외의 모든 명령을 막는다."""

    async def interaction_check(self, interaction: discord.Interaction) -> bool:
        if interaction.type != discord.InteractionType.application_command:
            return True
        until = blocked_until()
        if until and interaction.user.id != schedule.get("block_by"):
            lang = lang_for(interaction, getattr(interaction.namespace, "lang", None))
            await interaction.response.send_message(block_text(until, lang), ephemeral=True)
            return False
        return True


tree = GuardedTree(client)


async def jenkins_get(url: str) -> dict:
    assert http is not None
    async with http.get(url) as response:
        if response.status != 200:
            body = await response.text()
            raise RuntimeError(
                f"Jenkins 조회 실패: HTTP {response.status} {body[:300]}"
            )
        return await response.json()


async def jenkins_text(url: str) -> str:
    assert http is not None
    async with http.get(url) as response:
        if response.status != 200:
            body = await response.text()
            raise RuntimeError(
                f"Jenkins 텍스트 조회 실패: HTTP {response.status} {body[:300]}"
            )
        return await response.text(errors="replace")


async def jenkins_bytes(url: str, limit: int) -> bytes:
    assert http is not None
    async with http.get(url) as response:
        if response.status != 200:
            raise RuntimeError(f"Jenkins 파일 조회 실패: HTTP {response.status}")
        if (response.content_length or 0) > limit:
            raise ValueError(f"크기 제한 초과 ({response.content_length / 1024 / 1024:.1f}MB)")
        return await response.read()


# Jenkins Job이 이 폴더에 결과물(CSV, run-report.json)을 모아 'Archive the artifacts'로 보관한다
ARTIFACT_DIR_PREFIX = "bat-artifacts/"


async def fetch_build_artifacts(build_url: str) -> tuple[dict[str, bytes], list[str]]:
    """Build에 보관된 bat-artifacts/ 파일을 내려받는다. 보관 설정 전 Build면 빈 dict."""
    base = build_url.rstrip("/") + "/"
    info = await jenkins_get(base + "api/json?tree=artifacts[fileName,relativePath]")
    files: dict[str, bytes] = {}
    problems: list[str] = []
    for artifact in info.get("artifacts") or []:
        relative = str(artifact.get("relativePath") or "")
        name = str(artifact.get("fileName") or "")
        if not relative.startswith(ARTIFACT_DIR_PREFIX) or not name:
            continue
        try:
            files[name] = await jenkins_bytes(base + "artifact/" + quote(relative), MAX_CSV_SIZE_BYTES)
        except (RuntimeError, ValueError) as exc:
            problems.append(f"{name}: {exc}")
    return files, problems


async def load_result_csvs(build_url: str, console_text: str,
                           lang: str = DEFAULT_LANG) -> tuple[list[tuple[str, bytes]], list[str]]:
    """결과 CSV: Jenkins 아티팩트를 먼저 쓰고, 없으면 예전 방식(로그의 '복사:' 줄 + 결과 폴더)."""
    try:
        files, problems = await fetch_build_artifacts(build_url)
    except Exception:
        log.exception("Jenkins 아티팩트 조회 실패 (결과 폴더 방식으로 대체)")
        files, problems = {}, []
    report = files.get("run-report.json")
    if report:
        try:
            await asyncio.to_thread(record_report_durations, report)
        except Exception:
            log.exception("run-report.json 소요 시간 기록 실패")
    csvs = [(name, data) for name, data in files.items() if name.lower().endswith(".csv")]
    if csvs:
        return csvs[:MAX_CSV_ATTACHMENTS], problems
    names = parse_copied_csv_names(console_text) if console_text else []
    if not console_text:
        return [], [tr(lang, "콘솔 로그를 읽지 못해 CSV를 확인할 수 없습니다.",
                       "Could not read the console log, so the CSV cannot be checked.")]
    return await asyncio.to_thread(_read_csv_files, names, lang)


async def get_job() -> dict:
    return await jenkins_get(JOB_API_URL)


async def get_active_state(job: dict, lang: str = DEFAULT_LANG) -> str | None:
    if job.get("inQueue"):
        return tr(lang, "대기 중", "queued")

    last_build = job.get("lastBuild")
    if last_build and last_build.get("url"):
        try:
            build = await jenkins_get(
                last_build["url"].rstrip("/")
                + "/api/json?tree=number,building,result"
            )
            if build.get("building"):
                return tr(lang, f"실행 중 (빌드 #{build.get('number')})",
                          f"running (build #{build.get('number')})")
        except Exception:
            log.exception("마지막 빌드 상태 조회 실패")

    queue = await jenkins_get(QUEUE_API_URL)
    expected_url = f"{JENKINS_URL}{JOB_PATH}".rstrip("/")
    for item in queue.get("items", []):
        task = item.get("task", {})
        task_url = task.get("url", "")
        if task_url.rstrip("/") == expected_url or task.get("name") == JENKINS_JOB:
            return tr(lang, "실행 대기열에 등록됨", "in the Jenkins queue")
    return None


async def job_parameter_names() -> set[str]:
    job = await jenkins_get(JOB_API_URL + "?tree=property[parameterDefinitions[name]]")
    names: set[str] = set()
    for prop in job.get("property") or []:
        for definition in (prop or {}).get("parameterDefinitions") or []:
            if definition.get("name"):
                names.add(definition["name"])
    return names


async def find_active_target() -> tuple[str, dict] | None:
    """취소 대상: 대기열 항목 또는 실행 중인 Build."""
    queue = await jenkins_get(QUEUE_API_URL + "?tree=items[id,task[name,url]]")
    expected_url = f"{JENKINS_URL}{JOB_PATH}".rstrip("/")
    for item in queue.get("items", []):
        task = item.get("task", {})
        if task.get("url", "").rstrip("/") == expected_url or task.get("name") == JENKINS_JOB:
            return "queue", {"id": item.get("id")}

    job = await jenkins_get(JOB_API_URL + "?tree=lastBuild[number,url,building]")
    last = job.get("lastBuild") or {}
    if last.get("building"):
        build = await jenkins_get(last["url"].rstrip("/") + "/" + BUILD_API_TREE)
        return "build", {"number": last.get("number"), "url": last.get("url"), "build": build}
    return None


async def jenkins_post(url: str) -> int:
    assert http is not None
    async with http.post(url, allow_redirects=False) as response:
        return response.status


async def trigger_job(
    stage: str, product: str, mode: str | None = None, products: list[str] | None = None,
    override_id: str | None = None, lang: str | None = None,
) -> str | None:
    """STAGE, PRODUCT(+ MODE, PRODUCTS, OVERRIDE_ID, LANG) 파라미터를 전달해 빌드를 요청한다."""
    assert http is not None
    url = f"{JENKINS_URL}{JOB_PATH}/buildWithParameters"
    parameters = {"STAGE": stage, "PRODUCT": product}
    if mode:
        parameters["MODE"] = mode
    if products:
        parameters["PRODUCTS"] = ",".join(products)
    if override_id:
        parameters["OVERRIDE_ID"] = override_id  # 값은 파일로 넘기고 파라미터에는 id만
    if lang:
        parameters["LANG"] = lang  # 결과 엑셀 언어 (Job에 LANG 파라미터가 있을 때만 넘어옴)

    async with http.post(url, data=parameters, allow_redirects=False) as response:
        if response.status not in (200, 201, 202, 302):
            body = await response.text()
            raise RuntimeError(
                f"Jenkins 실행 요청 실패: HTTP {response.status} {body[:300]}"
            )
        location = response.headers.get("Location")
        return urljoin(JENKINS_URL + "/", location) if location else None


# /bat custom: 워크스페이스·계정·토큰을 Jenkins 파라미터(빌드 정보에 남음) 대신 이 폴더의 파일로 넘긴다.
# Jenkins가 OVERRIDE_ID로 파일을 찾아 환경변수로 넣고 바로 지운다. BAT은 환경변수를 .env보다 우선한다.
OVERRIDE_DIR = BOT_DATA_DIR / "bat-override"
OVERRIDE_MAX_AGE = 3600  # 초. Jenkins가 못 지운 파일은 다음 custom 실행 때 정리


def remove_override_file(override_id: str) -> None:
    (OVERRIDE_DIR / f"{override_id}.env").unlink(missing_ok=True)


def write_override_file(override: dict) -> str:
    """override: workspace(필수), user·password(선택), token(선택). 파일 id를 돌려준다."""
    OVERRIDE_DIR.mkdir(parents=True, exist_ok=True)
    now = time.time()
    for old in OVERRIDE_DIR.glob("*.env"):
        if now - old.stat().st_mtime > OVERRIDE_MAX_AGE:
            old.unlink(missing_ok=True)

    workspace = override["workspace"]
    lines = [f"BAT_WORKSPACE={workspace}"]
    if override.get("user") and override.get("password"):
        lines += [f"BAT_CONSOLE_USER={override['user']}", f"BAT_CONSOLE_PASSWORD={override['password']}"]
    if override.get("token"):
        lines.append(f"BAT_OPENAPI_TOKEN_{workspace.upper()}={override['token']}")
    override_id = os.urandom(8).hex()
    (OVERRIDE_DIR / f"{override_id}.env").write_text("\n".join(lines) + "\n", encoding="utf-8")
    return override_id


# BAT 결과 상태. FAIL = 치명적 결함, WARN = 확인 필요, BAT_ERROR = 테스트 도구·환경 문제(제품 이슈 아님).
# 예전 BAT는 BAT_ERROR를 ERROR로, WARN 없이 냈다 → 읽을 때 ERROR는 BAT_ERROR로 바꿔 다룬다.
VERDICT_KEYS = ("PASS", "FAIL", "WARN", "UNVERIFIED", "SKIP", "BAT_ERROR")
VERDICT_LINE_PATTERN = re.compile(r"BAT verdict[^\n]*", re.IGNORECASE)
VERDICT_PAIR_PATTERN = re.compile(r"\b(PASS|FAIL|WARN|UNVERIFIED|SKIP|BAT_ERROR|ERROR)\s+(\d+)", re.IGNORECASE)
GATE_PATTERN = re.compile(r"\[BAT-GATE\]\s+(ROLLBACK|REVIEW|OK)\b", re.IGNORECASE)
GATE_REASON_PATTERN = re.compile(r"\[BAT-GATE\]\s+reason:\s*(.+)", re.IGNORECASE)


def norm_status(status: str) -> str:
    status = (status or "").strip().upper()
    return "BAT_ERROR" if status == "ERROR" else status


def parse_verdicts(console_text: str) -> dict[str, int] | None:
    """콘솔 로그의 마지막 'BAT verdict' 집계 행 → {상태: 개수}. 옛 형식(ERROR, WARN 없음)과 새 형식 모두."""
    lines = VERDICT_LINE_PATTERN.findall(console_text or "")
    for line in reversed(lines):
        pairs = VERDICT_PAIR_PATTERN.findall(line)
        if not pairs:
            continue
        counts = dict.fromkeys(VERDICT_KEYS, 0)
        for name, value in pairs:
            counts[norm_status(name)] = int(value)
        return counts
    return None


def parse_gate(console_text: str) -> tuple[str, list[str]] | None:
    """콘솔 로그의 마지막 [BAT-GATE] 판정(ROLLBACK/REVIEW/OK)과 그 뒤의 사유들. 옛 BAT면 None."""
    matches = list(GATE_PATTERN.finditer(console_text or ""))
    if not matches:
        return None
    last = matches[-1]
    reasons = [m.group(1).strip() for m in GATE_REASON_PATTERN.finditer(console_text, last.end())]
    return last.group(1).upper(), reasons


def _parse_iso(value: str) -> datetime:
    # PowerShell의 ToString('o')는 소수 초가 7자리라서 Python이 읽을 수 있게 6자리로 맞춘다.
    return datetime.fromisoformat(re.sub(r"(\.\d{6})\d+", r"\1", value.strip()))


def parse_test_duration(console_text: str) -> timedelta | None:
    """Build Step이 남긴 [BAT-TIMING] START/END 마커로 웹훅과 같은 구간을 계산한다."""
    found: dict[str, str] = {}
    for match in TIMING_PATTERN.finditer(console_text):
        found[match.group(1)] = match.group(2)  # 같은 마커가 여러 번이면 마지막 값
    if "START" not in found or "END" not in found:
        return None
    try:
        span = _parse_iso(found["END"]) - _parse_iso(found["START"])
    except ValueError:
        log.warning("BAT-TIMING 시각 형식을 읽지 못함")
        return None
    return span if span.total_seconds() >= 0 else None


def classify_result(
    jenkins_result: str | None, counts: dict[str, int] | None, lang: str = DEFAULT_LANG,
    gate: str | None = None,
) -> tuple[str, int, str]:
    """알림 상태를 정한다. 롤백 판정 문구는 보여주지 않는다 (롤백 여부는 사람이 결과를 보고 정함).

    gate 행이 있으면 새 BAT(FAIL = 치명적 결함만)라 집계로 정하고, 없으면(옛 BAT) Jenkins 결과와 집계로 정한다.
    """
    result = (jenkins_result or "UNKNOWN").upper()
    if result == "ABORTED":
        return "aborted", 0x95A5A6, tr(lang, "빌드가 중단되었습니다.", "The build was aborted.")
    if gate and counts:
        if counts["FAIL"] > 0:
            return "critical", 0xE74C3C, tr(lang, "치명적 결함(FAIL)이 있습니다.",
                                            "There are critical defects (FAIL).")
        if any(counts[key] > 0 for key in ("WARN", "UNVERIFIED", "BAT_ERROR")):
            return "review", 0x2ECC71, tr(lang, "치명적 결함은 없고, 확인이 필요한 결과가 있습니다.",
                                          "No critical defects; some results need review.")
        return "success", 0x2ECC71, tr(lang, "테스트 실행이 정상적으로 완료되었습니다.",
                                       "The test run completed successfully.")
    if result in ("SUCCESS", "UNSTABLE") and counts:
        if counts["FAIL"] == 0 and counts["BAT_ERROR"] == 0 and counts["UNVERIFIED"] > 0:
            # 실패 없이 UNVERIFIED만 있는 경우
            return "unverified", 0xF1C40F, tr(lang, "일부 테스트가 확인 필요(UNVERIFIED) 상태입니다.",
                                              "Some tests need review (UNVERIFIED).")
        if any(counts[key] > 0 for key in ("FAIL", "WARN", "BAT_ERROR")):
            return "unstable", 0xF1C40F, tr(lang, "일부 테스트에서 실패가 발생했습니다.", "Some tests failed.")
    if result == "SUCCESS":
        return "success", 0x2ECC71, tr(lang, "테스트 실행이 정상적으로 완료되었습니다.",
                                       "The test run completed successfully.")
    if result == "UNSTABLE":
        return "unstable", 0xF1C40F, tr(lang, "일부 테스트에서 실패 또는 확인이 필요한 결과가 발생했습니다.",
                                        "Some tests failed or need review.")
    return "failure", 0xE74C3C, tr(lang, "실행 오류 또는 테스트 실패가 발생했습니다.",
                                   "The run hit an error or tests failed.")


def format_duration(milliseconds: int | float | None, lang: str = DEFAULT_LANG) -> str:
    if milliseconds is None or milliseconds < 0:
        return "-"
    seconds = int(milliseconds // 1000)  # 웹훅과 같이 초 미만은 버림
    hours, remainder = divmod(seconds, 3600)
    minutes, seconds = divmod(remainder, 60)
    if hours:
        return tr(lang, f"{hours}시간 {minutes}분 {seconds}초", f"{hours}h {minutes}m {seconds}s")
    return tr(lang, f"{minutes}분 {seconds}초", f"{minutes}m {seconds}s")


def _build_params(build: dict) -> dict[str, str]:
    params: dict[str, str] = {}
    for action in build.get("actions") or []:
        for param in (action or {}).get("parameters") or []:
            params[str(param.get("name"))] = str(param.get("value") or "")
    return params


def get_build_mode(build: dict) -> str:
    """MODE 파라미터 (없으면 일반 테스트 실행 run)."""
    return (_build_params(build).get("MODE") or "run").strip().lower()


def get_build_parameters(build: dict) -> tuple[str, str]:
    """Build에 실제로 전달된 STAGE, PRODUCT를 읽는다 (Jenkins UI 실행 포함)."""
    params = _build_params(build)
    stage = params.get("STAGE") or "-"
    if params.get("PRODUCTS", "").strip():  # 사전 설정 실행 (여러 제품)
        product = ", ".join(p.strip() for p in params["PRODUCTS"].split(",") if p.strip())
        return stage, product
    product = params.get("PRODUCT") or "all"  # Build Step의 PRODUCT_TAG 규칙과 동일
    return stage, product


def parse_copied_csv_names(console_text: str) -> list[str]:
    """Build Step의 '복사: 파일명' 줄에서 이번 Build가 복사한 CSV 이름을 찾는다."""
    names: list[str] = []
    for match in COPY_LINE_PATTERN.finditer(console_text):
        name = Path(match.group(1).strip()).name  # 경로가 섞여 있어도 파일명만 사용
        if name and name not in names:
            names.append(name)
    return names


def _read_csv_files(names: list[str], lang: str = DEFAULT_LANG) -> tuple[list[tuple[str, bytes]], list[str]]:
    """결과 폴더에서 CSV를 읽는다. 실패 원인은 구분해서 돌려준다."""
    contents: list[tuple[str, bytes]] = []
    problems: list[str] = []

    if not names:
        return contents, [tr(lang, "이번 Build 로그에 CSV 복사 기록이 없습니다.",
                             "This build's log has no record of copying a CSV.")]
    if not BAT_RESULT_DIR:
        return contents, [tr(lang, "BAT_RESULT_DIR 설정이 없어 CSV를 찾을 수 없습니다.",
                             "BAT_RESULT_DIR is not set, so the CSV cannot be found.")]

    base = Path(BAT_RESULT_DIR)
    try:
        if not base.is_dir():
            return contents, [tr(lang, f"결과 폴더가 없습니다: {base}", f"Result folder not found: {base}")]
    except PermissionError:
        return contents, [tr(lang, f"결과 폴더 접근 권한이 없습니다: {base}", f"No permission to access the result folder: {base}")]

    for name in names[:MAX_CSV_ATTACHMENTS]:
        path = base / name
        try:
            size = path.stat().st_size
            if size > MAX_CSV_SIZE_BYTES:
                problems.append(tr(lang, f"{name}: 크기 제한 초과 ({size / 1024 / 1024:.1f}MB)",
                                   f"{name}: exceeds the size limit ({size / 1024 / 1024:.1f}MB)"))
                continue
            contents.append((name, path.read_bytes()))
        except FileNotFoundError:
            problems.append(tr(lang, f"{name}: 파일 없음", f"{name}: file not found"))
        except PermissionError:
            problems.append(tr(lang, f"{name}: 접근 권한 없음", f"{name}: permission denied"))
        except OSError as exc:
            problems.append(tr(lang, f"{name}: 읽기 오류 ({exc.__class__.__name__})",
                               f"{name}: read error ({exc.__class__.__name__})"))

    if len(names) > MAX_CSV_ATTACHMENTS:
        extra = len(names) - MAX_CSV_ATTACHMENTS
        problems.append(tr(lang, f"첨부 개수 제한으로 {extra}개 제외", f"{extra} file(s) left out (attachment limit)"))
    return contents, problems


def excel_note_text(note: str, lang: str) -> str:
    """csv_to_xlsx.py가 돌려주는 한국어 안내를 영어로 바꾼다 (모르는 문구는 그대로)."""
    if lang == "ko":
        return note
    if note == "결과 행이 없어 표 서식은 적용하지 않음":
        return "no result rows, table formatting not applied"
    if note == "상태 열을 찾지 못해 상태 색은 적용하지 않음":
        return "status column not found, status colors not applied"
    m = re.fullmatch(r"엑셀 셀 길이 제한으로 (\d+)개 칸 일부 생략", note)
    if m:
        return f"{m.group(1)} cell(s) truncated (Excel cell length limit)"
    return note


def _convert_to_excel(
    contents: list[tuple[str, bytes]], lang: str = DEFAULT_LANG
) -> tuple[list[tuple[str, bytes]], list[str]]:
    """CSV를 서식 적용된 xlsx로 바꾼다. 실패하면 그 파일은 CSV 그대로 보낸다."""
    converted: list[tuple[str, bytes]] = []
    notes: list[str] = []
    for name, data in contents:
        if csv_to_xlsx is None:
            converted.append((name, data))
            notes.append(tr(lang, f"{name}: 엑셀 변환 모듈이 없어 CSV로 첨부",
                            f"{name}: attached as CSV (Excel converter not available)"))
            continue
        excel_name = Path(name).with_suffix(".xlsx").name
        try:
            excel_bytes, excel_notes = csv_to_xlsx(data)
            if len(excel_bytes) > MAX_CSV_SIZE_BYTES:
                raise ValueError("엑셀 파일 크기 제한 초과")
            converted.append((excel_name, excel_bytes))
            notes.extend(f"{excel_name}: {excel_note_text(note, lang)}" for note in excel_notes)
        except Exception as exc:
            log.exception("엑셀 변환 실패: %s", name)
            converted.append((name, data))
            notes.append(tr(lang, f"{name}: 엑셀 변환 실패 ({exc.__class__.__name__}), CSV로 첨부",
                            f"{name}: Excel conversion failed ({exc.__class__.__name__}), attached as CSV"))
    return converted, notes


# show 옵션 없이 보이는 상태. show:all이면 SKIP·UNVERIFIED·미실행까지 전부 보인다.
SHOW_DEFAULT_STATUSES = ("PASS", "FAIL", "WARN", "BAT_ERROR", "ERROR")  # ERROR = 옛 BAT의 BAT_ERROR
HIDDEN_BY_DEFAULT = ("UNVERIFIED", "SKIP")


def filter_csv_rows(data: bytes, statuses=SHOW_DEFAULT_STATUSES) -> bytes:
    """Status 열이 statuses에 든 행만 남긴 CSV. Status 열이 없으면 그대로 돌려준다."""
    rows = list(csv.reader(io.StringIO(_decode_bytes(data), newline="")))
    if not rows:
        return data
    header = [name.strip().lower() for name in rows[0]]
    if "status" not in header:
        return data
    index = header.index("status")
    kept = [rows[0]] + [r for r in rows[1:] if len(r) > index and r[index].strip().upper() in statuses]
    out = io.StringIO()
    csv.writer(out, lineterminator="\r\n").writerows(kept)
    return out.getvalue().encode("utf-8-sig")


MODE_LABELS = {"doctor": "🩺 bat doctor", "sweep": "🧹 bat sweep"}
MODE_START_PATTERN = re.compile(r"\[BAT-MODE\] start (\w+)")
MODE_END_PATTERN = re.compile(r"\[BAT-MODE\] exit=(-?\d+)")
# "키=값", "키: 값" 형태만 가린다 ("session expired" 같은 일반 문장은 그대로 둠)
SECRET_KV_PATTERN = re.compile(
    r"(?i)\b(token|secret|password|passwd|pwd|cookie|session[_-]?id|authorization|api[_-]?key)"
    r"(\s*[:=]\s*)(\S+)"
)
BEARER_PATTERN = re.compile(r"(?i)\b(bearer)\s+\S+")
LONG_TOKEN_PATTERN = re.compile(r"[A-Za-z0-9_\-.=+/]{32,}")


def mask_secrets(text: str) -> str:
    """Discord에 올리기 전에 토큰·비밀번호처럼 보이는 값을 가린다."""
    text = SECRET_KV_PATTERN.sub(lambda m: f"{m.group(1)}{m.group(2)}***", text)
    text = BEARER_PATTERN.sub(lambda m: f"{m.group(1)} ***", text)
    return LONG_TOKEN_PATTERN.sub("[masked]", text)


def extract_mode_output(
    console_text: str, max_lines: int = 25, max_chars: int = 950, lang: str = DEFAULT_LANG
) -> str | None:
    start = MODE_START_PATTERN.search(console_text)
    if not start:
        return None
    end = MODE_END_PATTERN.search(console_text, start.end())
    body = console_text[start.end(): end.start() if end else len(console_text)]
    lines = [line.rstrip() for line in body.replace("\r", "").split("\n") if line.strip()]
    text = mask_secrets("\n".join(lines[-max_lines:])).replace("```", "`\u200b``")
    if len(text) > max_chars:
        text = "...\n" + text[-max_chars:]
    return text or tr(lang, "(출력 없음)", "(no output)")


# 테스트 실행이 실패했을 때 콘솔에서 원인으로 보여 줄 줄
FAILURE_LINE_PATTERN = re.compile(
    r"(?i)\bFAIL(ED)?\b|\bERROR\b|Problems|Not running|Traceback|Exception|Duplicate login|"
    r"실행 실패|종료 코드"
)
FAILURE_SKIP_PATTERN = re.compile(r"\[BAT-TIMING\]|^Build step |^Finished:")


def extract_failure_reason(
    console_text: str, max_lines: int = 14, max_chars: int = 950, lang: str = DEFAULT_LANG
) -> str:
    """콘솔 로그에서 실패 원인 줄(과 바로 다음 설명 줄)을 골라 온다. 없으면 마지막 줄들."""
    lines = [
        line.rstrip()
        for line in console_text.replace("\r", "").split("\n")
        if line.strip() and not FAILURE_SKIP_PATTERN.search(line.strip())
    ]
    picked: list[str] = []
    for index, line in enumerate(lines):
        if FAILURE_LINE_PATTERN.search(line):
            for extra in lines[index:index + 2]:  # 오류 줄 + 다음 설명 줄
                if extra not in picked:
                    picked.append(extra)
    if not picked:
        picked = lines[-max_lines:]
    text = mask_secrets("\n".join(picked[-max_lines:])).replace("```", "`​``")
    if len(text) > max_chars:
        text = "...\n" + text[-max_chars:]
    return text or tr(lang, "(출력 없음)", "(no output)")


async def send_maintenance_message(channel, build: dict, mode: str, console_text: str,
                                   lang: str = DEFAULT_LANG) -> None:
    build_number = build.get("number", "-")
    result = (build.get("result") or "UNKNOWN").upper()
    stage, _ = get_build_parameters(build)
    if result == "SUCCESS":
        emoji, color, state = "🟢", 0x2ECC71, tr(lang, "정상", "OK")
    elif result == "ABORTED":
        emoji, color, state = "⚪", 0x95A5A6, tr(lang, "중단", "aborted")
    else:
        emoji, color, state = "🔴", 0xE74C3C, tr(lang, "실패", "failed")

    embed = discord.Embed(
        title=f"{emoji} {MODE_LABELS.get(mode, mode)} {state}",
        description=f"Job: {JENKINS_JOB}\nBuild: {build_number}",
        color=color,
        timestamp=discord.utils.utcnow(),
    )
    embed.add_field(name=tr(lang, "스테이지", "Stage"), value=stage, inline=True)
    embed.add_field(name=tr(lang, "Jenkins 결과", "Jenkins result"), value=result, inline=True)
    embed.add_field(name=tr(lang, "소요 시간", "Duration"), value=format_duration(build.get("duration"), lang), inline=True)
    output = extract_mode_output(console_text, lang=lang) if console_text else None
    embed.add_field(
        name=tr(lang, "출력 (마지막 부분, 민감정보 가림)", "Output (last part, secrets masked)"),
        value=f"```\n{output}\n```" if output else tr(
            lang, "출력 구간을 찾지 못했습니다 (Build Step의 [BAT-MODE] 표시 확인).",
            "Could not find the output section (check the [BAT-MODE] markers in the Build Step)."),
        inline=False,
    )
    embed.set_footer(text=f"Jenkins CI • Build #{build_number}")
    await channel.send(embed=embed)


async def send_completion_message(channel, build_url: str, build: dict, options: dict | None = None) -> None:
    show_all = (options or {}).get("show") == "all"
    # 요청 때 정한 언어 > 채널 설정 > en (Jenkins 화면에서 실행한 Build는 options가 없음)
    lang = norm_lang((options or {}).get("lang")) or channel_lang(getattr(channel, "id", None))
    build_number = build.get("number", "-")
    jenkins_result = build.get("result")
    stage, product = get_build_parameters(build)

    console_text = ""
    counts = None
    try:
        console_text = await jenkins_text(build_url.rstrip("/") + "/consoleText")
        counts = parse_verdicts(console_text)
    except Exception:
        log.exception("콘솔 로그 조회 실패")
    gate = parse_gate(console_text)

    mode = get_build_mode(build)
    if mode in MODE_LABELS:
        await send_maintenance_message(channel, build, mode, console_text, lang)
        return

    status, color, summary = classify_result(jenkins_result, counts, lang, gate[0] if gate else None)
    emoji, title = {
        "critical": ("🔴", tr(lang, "치명적 결함 발생", "Critical defects found")),
        "review": ("🟢", tr(lang, "테스트 확인 필요", "Tests need review")),
        "success": ("🟢", tr(lang, "테스트 실행 성공", "Test run passed")),
        "unstable": ("🟡", tr(lang, "테스트 일부 실패", "Some tests failed")),
        "unverified": ("🟡", tr(lang, "테스트 일부 확인 필요", "Some tests need review")),
        "failure": ("🔴", tr(lang, "테스트 실행 실패", "Test run failed")),
        "aborted": ("⚪", tr(lang, "테스트 중단", "Test run aborted")),
    }[status]

    embed = discord.Embed(
        title=f"{emoji} {title}",
        description=(
            f"{emoji} **{summary}**\n\n"
            f"Job: {JENKINS_JOB}\n"
            f"Build: {build_number}"
        ),
        color=color,
        timestamp=discord.utils.utcnow(),
    )
    embed.add_field(name=tr(lang, "프로젝트 / 제품", "Project / Product"), value=product, inline=True)
    embed.add_field(name=tr(lang, "스테이지", "Stage"), value=stage, inline=True)

    # 소요 시간: 웹훅과 같은 구간 (Build Step의 BUILD_START ~ 알림 직전)
    test_span = parse_test_duration(console_text)
    if test_span is not None:
        duration_text = format_duration(test_span.total_seconds() * 1000, lang)
    else:
        duration_text = tr(lang, "측정 불가 (BAT-TIMING 로그 없음)", "Unavailable (no BAT-TIMING log)")
    embed.add_field(name=tr(lang, "소요 시간", "Duration"), value=duration_text, inline=True)

    if counts is not None:
        embed.add_field(name="🟢 PASS", value=str(counts["PASS"]), inline=True)
        # 옛 BAT의 FAIL은 치명 기준이 아니므로 새 BAT(gate 행 있음)일 때만 '(치명)'을 붙인다
        fail_name = tr(lang, "🔴 FAIL (치명)", "🔴 FAIL (critical)") if gate else "🔴 FAIL"
        embed.add_field(name=fail_name, value=str(counts["FAIL"]), inline=True)
        if counts["WARN"] > 0:
            embed.add_field(name=tr(lang, "🟡 WARN (확인 필요)", "🟡 WARN (review)"), value=str(counts["WARN"]), inline=True)
        if counts["BAT_ERROR"] > 0:
            embed.add_field(name=tr(lang, "🧰 BAT_ERROR (제품 이슈 아님)", "🧰 BAT_ERROR (not a product issue)"),
                            value=str(counts["BAT_ERROR"]), inline=True)
        if show_all:
            embed.add_field(name="🟠 UNVERIFIED", value=str(counts["UNVERIFIED"]), inline=True)
            embed.add_field(name="⏭️ SKIP", value=str(counts["SKIP"]), inline=True)
        # show:all이 아니면 SKIP·UNVERIFIED는 개수도 표시하지 않는다 (사용자 요청)
    else:
        embed.add_field(
            name=tr(lang, "테스트 집계", "Test counts"),
            value=tr(lang, "집계 확인 불가 (콘솔 로그에서 BAT verdict를 찾지 못했습니다)",
                     "Counts unavailable (no BAT verdict found in the console log)"),
            inline=False,
        )

    if status == "failure" and console_text:
        embed.add_field(
            name=tr(lang, "❗ 실패 원인 (콘솔 로그 발췌, 민감정보 가림)",
                    "❗ Failure reason (console log excerpt, secrets masked)"),
            value=f"```\n{extract_failure_reason(console_text, lang=lang)}\n```",
            inline=False,
        )

    excluded_match = EXCLUDE_LINE_PATTERN.findall(console_text) if console_text else []
    if excluded_match:
        excluded_items = excluded_match[-1].split(",")
        embed.add_field(name=tr(lang, "🚫 제외 (all 실행)", "🚫 Excluded (all run)"),
                        value=", ".join(excluded_items)[:1024], inline=False)

    # CSV: Jenkins 아티팩트 우선, 없으면 콘솔 로그의 '복사:' 줄 + 결과 폴더
    contents, problems = await load_result_csvs(build_url, console_text, lang)

    if contents and not show_all:
        try:
            contents = [(name, await asyncio.to_thread(filter_csv_rows, data)) for name, data in contents]
        except Exception:
            log.exception("결과 CSV 상태 필터 실패 (전체 첨부)")
    if contents:
        contents, convert_notes = await asyncio.to_thread(_convert_to_excel, contents, lang)
        problems.extend(convert_notes)

    file_lines = [f"• {name}" for name, _ in contents] + [f"⚠️ {p}" for p in problems]
    embed.add_field(name=tr(lang, "📎 테스트 결과 파일", "📎 Test result files"),
                    value="\n".join(file_lines)[:1024], inline=False)

    # 기준이 다른 Jenkins 전체 시간은 footer에 구분해서 표시
    jenkins_time = format_duration(build.get("duration"), lang)
    embed.set_footer(
        text=f"Jenkins CI • Build #{build_number} • "
             + tr(lang, f"Jenkins 빌드 시간 {jenkins_time}", f"Jenkins build time {jenkins_time}")
    )

    files = [discord.File(io.BytesIO(data), filename=name) for name, data in contents]
    try:
        await channel.send(embed=embed, files=files if files else None)
    except discord.HTTPException:
        if not files:
            raise
        # 첨부 때문에 실패해도 결과 메시지는 보낸다
        log.exception("CSV 첨부 전송 실패, 첨부 없이 재전송")
        embed.add_field(name=tr(lang, "⚠️ 첨부 실패", "⚠️ Attachment failed"),
                        value=tr(lang, "결과 파일 첨부 전송에 실패해 결과만 보냅니다.",
                                 "Sending the result files failed, so only the summary is sent."),
                        inline=False)
        await channel.send(embed=embed)


def load_state() -> dict:
    try:
        data = json.loads(STATE_FILE.read_text(encoding="utf-8"))
        if isinstance(data, dict):
            return data
    except FileNotFoundError:
        pass
    except Exception:
        log.exception("알림 상태 파일을 읽지 못함")
    return {}


def save_state() -> None:
    for key_name in ("build_channels", "build_options"):
        records = state.get(key_name) or {}
        if len(records) > 20:  # 오래된 기록은 정리
            keep = sorted(records, key=int)[-20:]
            state[key_name] = {key: records[key] for key in keep}
    tmp = STATE_FILE.with_suffix(".tmp")
    tmp.write_text(json.dumps(state), encoding="utf-8")
    tmp.replace(STATE_FILE)


def remember_channel(channel_id: int, build_number: int | None = None, options: dict | None = None) -> None:
    """/bat를 사용한 채널을 기억한다. build_number가 있으면 그 Build의 완료 알림 채널이 된다.

    options: 그 Build 완료 알림에 쓸 실행 옵션 (show 등). 비밀값은 넣지 않는다.
    """
    state["default_channel_id"] = channel_id
    if build_number is not None:
        state.setdefault("build_channels", {})[str(build_number)] = channel_id
        clean = {k: v for k, v in (options or {}).items() if v}
        if clean:
            state.setdefault("build_options", {})[str(build_number)] = clean
    save_state()


async def resolve_channels(build_number: int) -> list:
    """알림 보낼 채널: 이 Build를 봇으로 실행한 채널 하나뿐.

    Jenkins 화면에서 직접 실행한 Build는 자동 알림을 보내지 않는다 (사용자 요청). 결과는 /bat-result로 본다.
    """
    own = (state.get("build_channels") or {}).get(str(build_number))
    if not own:
        return []
    channel_ids = [own]
    channels = []
    for channel_id in channel_ids:
        channel = client.get_channel(int(channel_id))
        if channel is None:
            try:
                channel = await client.fetch_channel(int(channel_id))
            except discord.HTTPException:
                log.exception("알림 채널 조회 실패: %s (채널 ID 또는 봇 권한 확인 필요)", channel_id)
                continue
        channels.append(channel)
    return channels


async def watch_builds() -> None:
    """봇으로 실행한 Build가 끝나면 실행한 채널로 한 번 알림을 보낸다. Jenkins 화면에서 실행한 Build는 건너뛴다."""
    await client.wait_until_ready()

    while "last_notified" not in state:
        # 첫 실행: 과거 Build는 알리지 않고 지금 시점부터 감시
        try:
            job = await jenkins_get(JOB_API_URL + "?tree=lastCompletedBuild[number]")
            state["last_notified"] = int((job.get("lastCompletedBuild") or {}).get("number", 0))
            save_state()
        except Exception:
            log.exception("초기 Build 번호 조회 실패, 재시도")
            await asyncio.sleep(WATCH_INTERVAL_SECONDS)
    log.info("완료 감시 시작: Build #%s 이후부터 알림", state["last_notified"])

    builds_url = JOB_API_URL + f"?tree=builds[number,building,url]{{0,{CATCHUP_LIMIT}}}"
    while True:
        try:
            data = await jenkins_get(builds_url)
            for item in sorted(data.get("builds") or [], key=lambda b: b.get("number", 0)):
                number = int(item.get("number", 0))
                if number <= int(state["last_notified"]):
                    continue
                if item.get("building"):
                    break  # 번호 순서대로 알림
                build_url = item.get("url") or f"{JENKINS_URL}{JOB_PATH}/{number}/"
                build = await jenkins_get(build_url.rstrip("/") + "/" + BUILD_API_TREE)
                own = (state.get("build_channels") or {}).get(str(number))
                if not own and monitor_task is not None and not monitor_task.done():
                    # 봇 요청의 Build 번호를 아직 확인 중이다. 이 Build가 그 요청일 수 있으니 다음 주기에 다시 본다.
                    break
                channels = await resolve_channels(number)
                if not own:
                    log.info("Build #%s 알림 안 함: 봇으로 실행하지 않은 Build (Jenkins 화면 실행 등)", number)
                elif not channels:
                    log.warning("Build #%s 알림 건너뜀: 실행 채널을 찾지 못함 (채널 삭제 또는 봇 권한)", number)
                options = (state.get("build_options") or {}).get(str(number))
                for channel in channels:
                    try:
                        await send_completion_message(channel, build_url, build, options)
                    except Exception:
                        # 한 채널이 실패해도 다른 채널 알림과 다음 Build는 계속 (같은 알림 중복 방지)
                        log.exception("Build #%s 알림 실패: 채널 %s", number, getattr(channel, "id", "?"))
                state["last_notified"] = number
                (state.get("build_channels") or {}).pop(str(number), None)
                (state.get("build_options") or {}).pop(str(number), None)
                save_state()
        except asyncio.CancelledError:
            raise
        except Exception:
            log.exception("완료 감시 오류 (다음 주기에 재시도)")
        await asyncio.sleep(WATCH_INTERVAL_SECONDS)


async def monitor_build(
    channel, queue_url: str | None, stage: str, product: str, label: str = "BAT run",
    options: dict | None = None,
) -> None:
    """/bat 요청이 실제 Build로 시작됐는지만 확인한다. 완료 알림은 watch_builds가 보낸다."""
    lang = norm_lang((options or {}).get("lang")) or channel_lang(getattr(channel, "id", None))
    try:
        build_url = None
        build_number = None

        # Jenkins 대기열에서 실제 빌드 번호를 확인한다.
        if queue_url:
            for _ in range(720):  # 최대 1시간 대기
                try:
                    item = await jenkins_get(queue_url.rstrip("/") + "/api/json")
                    if item.get("cancelled"):
                        await channel.send(tr(lang, f"{label} 요청이 대기열에서 취소됨: STAGE=`{stage}`",
                                              f"{label} request was cancelled in the queue: STAGE=`{stage}`"))
                        return
                    executable = item.get("executable")
                    if executable:
                        build_number = executable.get("number")
                        build_url = executable.get("url")
                        break
                except Exception:
                    log.exception("대기열 조회 오류")
                await asyncio.sleep(5)
        else:
            # Location 헤더가 없을 때, 새 빌드가 등록되는지 작업 API를 확인한다.
            original_last = (await get_job()).get("lastBuild") or {}
            original_number = original_last.get("number", 0)
            for _ in range(120):
                await asyncio.sleep(5)
                try:
                    job = await get_job()
                    last_build = job.get("lastBuild") or {}
                    if last_build.get("number", 0) > original_number:
                        build_number = last_build.get("number")
                        build_url = last_build.get("url")
                        break
                except Exception:
                    log.exception("새 빌드 확인 오류")

        if not build_url:
            await channel.send(tr(
                lang,
                "Jenkins가 실행 요청을 접수했지만 빌드 번호를 확인하지 못했습니다.\n"
                f"STAGE=`{stage}`, PRODUCT=`{product}`\n"
                "Jenkins에서 직접 확인해 주세요.",
                "Jenkins accepted the request, but the build number could not be confirmed.\n"
                f"STAGE=`{stage}`, PRODUCT=`{product}`\n"
                "Please check Jenkins directly.",
            ))
            return

        if build_number is not None and getattr(channel, "id", None):
            remember_channel(channel.id, int(build_number), options)

        product_line = f"PRODUCT: `{product}`\n" if label.startswith("BAT") else ""
        await channel.send(tr(
            lang,
            f"{label} 시작\nSTAGE: `{stage}`\n{product_line}"
            f"빌드: #{build_number}\n"
            "완료 결과는 이 채널로 전송됩니다.",
            f"{label} started\nSTAGE: `{stage}`\n{product_line}"
            f"Build: #{build_number}\n"
            "The result will be posted in this channel.",
        ))
    except Exception:
        log.exception("빌드 시작 확인 오류")
        try:
            await channel.send(tr(
                lang,
                f"빌드 시작 확인 중 오류 발생. STAGE=`{stage}`, PRODUCT=`{product}`. "
                "Jenkins에서 상태를 확인해 주세요.",
                f"Error while confirming the build start. STAGE=`{stage}`, PRODUCT=`{product}`. "
                "Please check the status in Jenkins.",
            ))
        except Exception:
            log.exception("Discord 알림 전송 실패")


@tree.command(
    name="ping",
    description=L("Check the Discord bot connection", "Discord 봇 연결 상태 확인"),
    guilds=TEST_GUILDS,
)
@with_lang
async def ping(interaction: discord.Interaction, lang: app_commands.Choice[str] | None = None):
    lang = lang_for(interaction, lang)
    await interaction.response.send_message(tr(lang, "Discord 봇 연결 정상", "Discord bot connection OK"), ephemeral=True)


@tree.command(
    name="bat-status",
    description=L("Show the Jenkins BAT job status", "Jenkins BAT 작업 상태 조회"),
    guilds=TEST_GUILDS,
)
@with_lang
async def bat_status(interaction: discord.Interaction, lang: app_commands.Choice[str] | None = None):
    lang = lang_for(interaction, lang)
    await interaction.response.defer(thinking=True)  # 채널 전체에 공개
    try:
        job = await get_job()
        active = await get_active_state(job, lang)
        if active:
            message = tr(lang, f"`{JENKINS_JOB}` 상태: **{active}**", f"`{JENKINS_JOB}` status: **{active}**")
        else:
            completed = job.get("lastCompletedBuild") or {}
            if completed:
                result = completed.get('result') or tr(lang, "미확인", "unknown")
                message = tr(
                    lang,
                    f"`{JENKINS_JOB}` 실행 중인 빌드 없음\n"
                    f"마지막 완료 빌드: #{completed.get('number')}\n"
                    f"결과: {result}",
                    f"`{JENKINS_JOB}` has no running build\n"
                    f"Last completed build: #{completed.get('number')}\n"
                    f"Result: {result}",
                )
            else:
                message = tr(lang, "완료된 빌드 기록이 없습니다.", "No completed builds yet.")
        await interaction.followup.send(message)
    except Exception:
        log.exception("상태 조회 실패")
        await interaction.followup.send(
            tr(lang, "Jenkins 상태 조회 실패. 봇 콘솔 로그를 확인해 주세요.",
               "Failed to get the Jenkins status. Please check the bot console log."),
        )


def job_not_ready_text(param: str, lang: str) -> str:
    return tr(lang, f"Jenkins Job에 {param} 파라미터가 없어 실행하지 않았습니다. Job 설정을 먼저 해 주세요.",
              f"Not started: the Jenkins job has no {param} parameter. Please configure the job first.")


def request_failed_text(what: str, lang: str) -> str:
    """'... 요청 실패. 봇 콘솔 로그를 확인해 주세요.' (what: 'Jenkins 실행', 'bat doctor' 등)"""
    en_what = "Jenkins run" if what == "Jenkins 실행" else what
    return tr(lang, f"{what} 요청 실패. 봇 콘솔 로그를 확인해 주세요.",
              f"{en_what} request failed. Please check the bot console log.")


def monitoring_text(lang: str) -> str:
    return tr(lang, "이미 다른 실행의 시작을 확인하는 중입니다. `/bat-status`로 확인해 주세요.",
              "Already waiting for another run to start. Check with `/bat-status`.")


def rejected_text(active: str, lang: str) -> str:
    return tr(lang, f"실행 요청 거절: 현재 작업 상태는 **{active}**입니다.",
              f"Request rejected: the job is currently **{active}**.")


async def start_job(
    channel, stage: str, product: str, mode: str, label: str, products: list[str] | None = None,
    override: dict | None = None, options: dict | None = None,
) -> str:
    """중복 실행을 확인한 뒤 Jenkins에 요청하고, 시작 확인 작업을 띄운다. 결과 문구를 돌려준다.

    options: 완료 알림에 쓸 실행 옵션 (show, lang 등). Build 번호가 정해지면 그 번호로 저장된다.
    """
    global monitor_task
    lang = norm_lang((options or {}).get("lang")) or channel_lang(getattr(channel, "id", None))
    async with bat_lock:
        if monitor_task is not None and not monitor_task.done():
            return monitoring_text(lang)

        job = await get_job()
        active = await get_active_state(job, lang)
        if active:
            return rejected_text(active, lang)

        param_names = await job_parameter_names()
        has_mode = "MODE" in param_names
        if mode != "run" and not has_mode:
            # MODE 없이 요청하면 Jenkins가 전체 BAT를 돌려 버리므로 막는다
            return job_not_ready_text("MODE", lang)

        if products and "PRODUCTS" not in param_names:
            # PRODUCTS 없이 요청하면 Jenkins가 PRODUCT=all(전체)로 돌려 버리므로 막는다
            return job_not_ready_text("PRODUCTS", lang)

        # LANG(결과 엑셀 언어)은 Job에 파라미터가 있을 때만 넘긴다. 없으면 그냥 실행한다.
        jenkins_lang = lang if mode == "run" and "LANG" in param_names else None

        override_id = None
        if override:
            if "OVERRIDE_ID" not in param_names:
                # 없이 요청하면 기본 워크스페이스로 돌아 버리므로 막는다
                return job_not_ready_text("OVERRIDE_ID", lang)
            override_id = write_override_file(override)
        try:
            queue_url = await trigger_job(
                stage, "all" if products else product, mode if has_mode else None, products, override_id,
                jenkins_lang,
            )
        except Exception:
            if override_id:
                remove_override_file(override_id)  # 요청이 안 갔으면 계정 파일을 남기지 않는다
            raise
        shown_product = ", ".join(products) if products else product
        if channel is not None:
            remember_channel(channel.id)
            monitor_task = asyncio.create_task(
                monitor_build(channel, queue_url, stage, shown_product, label, options)
            )
    product = shown_product
    product_line = f"\nPRODUCT: `{product}`" if mode == "run" else ""
    return tr(lang, f"{label} 요청 전송 완료\nSTAGE: `{stage}`{product_line}",
              f"{label} request sent\nSTAGE: `{stage}`{product_line}")


class ConfirmView(discord.ui.View):
    """요청한 사람만 누를 수 있는 확인/그만두기 버튼."""

    def __init__(self, requester_id: int, confirm_label: str, lang: str = DEFAULT_LANG):
        super().__init__(timeout=60)
        self.requester_id = requester_id
        self.lang = lang
        self.confirmed: bool | None = None
        self.confirm_button.label = confirm_label
        self.cancel_button.label = tr(lang, "취소", "Cancel")

    async def interaction_check(self, interaction: discord.Interaction) -> bool:
        if interaction.user.id != self.requester_id:
            await interaction.response.send_message(
                tr(self.lang, "요청한 분만 누를 수 있습니다.", "Only the requester can press this."), ephemeral=True)
            return False
        return True

    @discord.ui.button(label="OK", style=discord.ButtonStyle.danger)
    async def confirm_button(self, interaction: discord.Interaction, button: discord.ui.Button):
        self.confirmed = True
        await interaction.response.edit_message(content=tr(self.lang, "처리 중...", "Processing..."), view=None)
        self.stop()

    @discord.ui.button(label="Cancel", style=discord.ButtonStyle.secondary)
    async def cancel_button(self, interaction: discord.Interaction, button: discord.ui.Button):
        self.confirmed = False
        await interaction.response.edit_message(
            content=tr(self.lang, "요청을 취소했습니다.", "Request cancelled."), view=None)
        self.stop()


def no_response_text(lang: str) -> str:
    return tr(lang, "응답이 없어 요청을 취소했습니다. (60초)", "No response, so the request was cancelled. (60s)")


async def ask_confirm(interaction: discord.Interaction, question: str, confirm_label: str,
                      lang: str = DEFAULT_LANG) -> bool:
    view = ConfirmView(interaction.user.id, confirm_label, lang)
    await interaction.response.send_message(question, view=view)
    await view.wait()
    if view.confirmed is None:
        await interaction.edit_original_response(content=no_response_text(lang), view=None)
        return False
    return view.confirmed


stage_choices = [app_commands.Choice(name=value, value=value) for value in STAGES]


# ---- 제품별 예상 시간: Jenkins 실행 기록(단일 제품 run, 성공·UNSTABLE)의 최근 3번 중앙값 ----
DURATION_TTL = 300  # 초. 기록은 이 간격으로만 다시 읽는다 (자동완성은 3초 안에 답해야 함)
DURATION_SAMPLES = 3
duration_cache: dict[tuple[str, str], list[float]] = {}  # (stage, product) → 최근순 소요 분
duration_loaded_at = 0.0
duration_task: asyncio.Task | None = None


async def load_durations() -> None:
    global duration_loaded_at
    tree_query = "?tree=allBuilds[result,building,duration,actions[parameters[name,value]]]{0,200}"
    data = await jenkins_get(JOB_API_URL + tree_query)
    cache: dict[tuple[str, str], list[float]] = {}
    for build in data.get("allBuilds") or []:
        params = _build_params(build)
        if (build.get("building") or build.get("result") not in ("SUCCESS", "UNSTABLE")
                or get_build_mode(build) != "run" or params.get("PRODUCTS", "").strip()):
            continue
        key = (params.get("STAGE") or "-", params.get("PRODUCT") or "all")
        cache.setdefault(key, []).append(build.get("duration", 0) / 60000)
    duration_cache.clear()
    duration_cache.update(cache)
    try:
        report_cache = await asyncio.to_thread(load_report_durations)
        report_duration_cache.clear()
        report_duration_cache.update(report_cache)
    except Exception as exc:  # 보조 자료라 실패해도 Jenkins 기록만으로 계속
        log.warning("BAT 리포트 소요 시간 읽기 실패: %s", exc)
    duration_loaded_at = time.monotonic()


# ---- 보조 자료: BAT run-report.json의 케이스별 duration_sec를 제품별로 합한 값 ----
# 단독 실행 기록이 없는 제품(all 실행에만 섞여 돈 swg·dlp·cross 등)도 예상 시간을 보여주기 위함.
# update Job이 reports 폴더를 지워도 남도록 bot-data에 run_id별로 모아 둔다.
BAT_REPORTS_DIR = Path(os.environ.get("BAT_REPORTS_DIR", "").strip()
                       or r"C:\test-automation\aioncloud-sse-bat\reports")
REPORT_DURATIONS_FILE = BOT_DATA_DIR / "product-durations.json"
REPORT_DURATIONS_KEEP = 60          # 최근 실행 몇 개까지 보관
REPORT_OVERHEAD_MIN = 0.5           # 단독 실행 시 doctor·sweep 등 테스트 밖 시간 (basic 실측 기준)
report_duration_cache: dict[tuple[str, str], list[float]] = {}  # (stage, product) → 최근순 테스트 분


report_durations_lock = threading.Lock()  # 완료 알림과 주기 갱신이 동시에 파일을 고칠 수 있음


def _summarize_report(report: dict) -> tuple[str, dict] | None:
    """run-report.json → (run_id, {stage, started_at, products: {제품: 테스트 분}})."""
    run_id = str(report.get("run_id") or "")
    products: dict[str, float] = {}
    for case in report.get("cases") or []:
        product = str(case.get("product") or "")
        seconds = case.get("duration_sec")
        if product and isinstance(seconds, (int, float)):
            products[product] = products.get(product, 0.0) + seconds / 60
    if not run_id or not products:
        return None
    return run_id, {
        "stage": str(report.get("profile") or "-"),
        "started_at": str(report.get("started_at") or ""),
        "products": {p: round(m, 2) for p, m in products.items()},
    }


def _merge_report_durations(reports: list[dict]) -> dict:
    """보관 파일에 새 실행을 합쳐 저장하고 전체 내용을 돌려준다."""
    with report_durations_lock:
        try:
            stored = json.loads(REPORT_DURATIONS_FILE.read_text(encoding="utf-8"))
        except (FileNotFoundError, ValueError):
            stored = {}
        changed = False
        for report in reports:
            summary = _summarize_report(report)
            if summary and summary[0] not in stored:
                stored[summary[0]] = summary[1]
                changed = True
        if changed:
            newest = sorted(stored.items(), key=lambda kv: kv[1].get("started_at", ""), reverse=True)
            stored = dict(newest[:REPORT_DURATIONS_KEEP])
            REPORT_DURATIONS_FILE.write_text(json.dumps(stored, ensure_ascii=False, indent=2), encoding="utf-8")
        return stored


def record_report_durations(data: bytes) -> None:
    """Jenkins 아티팩트로 받은 run-report.json을 보관 파일에 더하고, 다음 자동완성 때 다시 읽게 한다."""
    global duration_loaded_at
    _merge_report_durations([json.loads(data.decode("utf-8-sig"))])
    duration_loaded_at = 0.0


def load_report_durations() -> dict[tuple[str, str], list[float]]:
    """보관 파일 + (아티팩트 보관 전 실행 대비) 로컬 reports/*/*/run-report.json → (stage, product) → 최근순 분."""
    local: list[dict] = []
    if BAT_REPORTS_DIR.is_dir():
        for path in BAT_REPORTS_DIR.glob("*/*/run-report.json"):
            try:
                local.append(json.loads(path.read_text(encoding="utf-8")))
            except (OSError, ValueError):
                continue
    stored = _merge_report_durations(local)
    cache: dict[tuple[str, str], list[float]] = {}
    for _, entry in sorted(stored.items(), key=lambda kv: kv[1].get("started_at", ""), reverse=True):
        for product, minutes in (entry.get("products") or {}).items():
            cache.setdefault((entry.get("stage") or "-", product), []).append(float(minutes))
    return cache


async def refresh_durations(wait: float = 0) -> None:
    """기록이 오래됐으면 백그라운드로 다시 읽는다. wait초까지는 읽기를 기다린다 (처음 한 번)."""
    global duration_task
    if time.monotonic() - duration_loaded_at < DURATION_TTL:
        return
    if duration_task is None or duration_task.done():
        duration_task = asyncio.create_task(load_durations())
        duration_task.add_done_callback(lambda t: t.cancelled() or not t.exception()
                                        or log.warning("예상 시간 기록 조회 실패: %s", t.exception()))
    if wait:
        await asyncio.wait({duration_task}, timeout=wait)


def expected_minutes(stage: str | None, product: str) -> int | None:
    """1) 그 제품 단독 Jenkins 빌드 시간, 2) 없으면 BAT 리포트의 그 제품 케이스 시간 합 + 준비 시간.
    각각 같은 스테이지 기록을 먼저 쓰고, 없으면 모든 스테이지 기록. 둘 다 없으면 None."""
    def pick(cache: dict[tuple[str, str], list[float]]) -> list[float]:
        samples = cache.get((stage, product)) if stage else None
        if not samples:
            samples = [m for (s, p), values in cache.items() if p == product for m in values]
        return samples or []

    samples = pick(duration_cache)
    if samples:
        return max(1, round(statistics.median(samples[:DURATION_SAMPLES])))
    samples = pick(report_duration_cache)
    if samples:
        return max(1, round(statistics.median(samples[:DURATION_SAMPLES]) + REPORT_OVERHEAD_MIN))
    return None


async def product_autocomplete(interaction: discord.Interaction, current: str):
    await refresh_durations(wait=2 if not duration_loaded_at else 0)
    stage = getattr(interaction.namespace, "stage", None)
    lang = lang_for(interaction, getattr(interaction.namespace, "lang", None))
    choices = []
    for product in PRODUCTS:
        if current.lower() not in product:
            continue
        minutes = expected_minutes(stage, product)
        if minutes:
            label = tr(lang, f"{product} [ 예상시간 {minutes}분 ]", f"{product} [ est. {minutes} min ]")
        else:
            label = product
        choices.append(app_commands.Choice(name=label, value=product))
    return choices


WORKSPACE_NAME_PATTERN = re.compile(r"^[A-Za-z0-9_-]{1,64}$")


class BatCustomModal(discord.ui.Modal, title="BAT custom workspace run"):
    # 입력칸은 언어에 맞는 이름으로 __init__에서 만든다 (만든 뒤 label을 바꾸면 discord.py가 경고를 낸다)

    def __init__(self, stage: str, product: str, options: dict | None = None):
        self.options = options or {}
        self.lang = lang = norm_lang(self.options.get("lang")) or DEFAULT_LANG
        super().__init__(title=tr(lang, "BAT 사용자 지정 워크스페이스 실행", "BAT custom workspace run"))
        self.stage = stage
        self.product = product
        self.workspace = discord.ui.TextInput(label=tr(lang, "워크스페이스", "Workspace"), required=True)
        self.email = discord.ui.TextInput(
            label=tr(lang, "이메일 (비우면 .env 기본 계정)", "Email (blank: default account in .env)"), required=False)
        self.password = discord.ui.TextInput(
            label=tr(lang, "비밀번호 (비우면 .env 기본 계정)", "Password (blank: default account in .env)"), required=False)
        self.token = discord.ui.TextInput(
            label=tr(lang, "REST 토큰 (선택)", "REST token (optional)"), required=False, style=discord.TextStyle.paragraph,
            placeholder=tr(lang, "비우면 .env의 BAT_OPENAPI_TOKEN_<워크스페이스> 또는 콘솔 로그인으로 발급",
                           "Blank: BAT_OPENAPI_TOKEN_<workspace> in .env, or issued via console login"),
        )
        for item in (self.workspace, self.email, self.password, self.token):
            self.add_item(item)

    async def on_submit(self, interaction: discord.Interaction):
        lang = self.lang
        workspace = self.workspace.value.strip()
        user, password = self.email.value.strip(), self.password.value
        if not WORKSPACE_NAME_PATTERN.match(workspace):
            await interaction.response.send_message(tr(
                lang, "워크스페이스 이름은 영문·숫자·`-`·`_`만 쓸 수 있습니다.",
                "Workspace names may only contain letters, digits, `-` and `_`."), ephemeral=True)
            return
        if bool(user) != bool(password):
            await interaction.response.send_message(tr(
                lang, "이메일과 비밀번호는 둘 다 넣거나 둘 다 비워 주세요.",
                "Fill in both email and password, or leave both blank."), ephemeral=True)
            return
        if any("\n" in v or "\r" in v for v in (user, password, self.token.value.strip())):
            await interaction.response.send_message(tr(
                lang, "값에 줄바꿈이 들어 있습니다. 다시 입력해 주세요.",
                "A value contains a line break. Please enter it again."), ephemeral=True)
            return
        override = {"workspace": workspace, "user": user, "password": password, "token": self.token.value.strip()}
        try:
            await request_custom_run(interaction, self.stage, self.product, override, self.options)
        finally:
            override.clear()


def reservation_soon_text(r: dict, minutes: int, lang: str) -> str:
    """바로 실행하려는데 다른 사람의 예약이 곧일 때 묻는 문구."""
    at = show_time(res_time(r), lang)
    return tr(
        lang,
        f"⏰ **{r.get('plan', '-')}** 사전 설정 실행 예약 {minutes}분 전입니다. "
        f"({at} · 예약자 {r['user_name']})\n"
        "그래도 진행할까요? 진행하면 예약된 사전 설정은 10분 미뤄집니다.",
        f"⏰ The **{r.get('plan', '-')}** preset run is reserved in {minutes} min. "
        f"({at} · reserved by {r['user_name']})\n"
        "Proceed anyway? If you proceed, the reserved preset run is postponed by 10 minutes.",
    )


def ran_first_text(user_name: str, r: dict) -> str:
    """예약을 미룬 이유 (예약 항목의 언어)."""
    return tr(req_lang(r), f"{user_name}님이 먼저 실행", f"{user_name} started a run first")


async def request_custom_run(
    interaction: discord.Interaction, stage: str, product: str, override: dict, options: dict | None = None
) -> None:
    """/bat custom 실행. 계정·토큰이 디스크에 오래 남지 않도록 대기열·예약에는 넣지 않는다."""
    lang = norm_lang((options or {}).get("lang")) or lang_for(interaction)
    await interaction.response.defer(thinking=True)  # 채널 전체에 공개 (입력값은 보이지 않음)
    workspace = override["workspace"]
    try:
        if "OVERRIDE_ID" not in await job_parameter_names():
            await interaction.followup.send(job_not_ready_text("OVERRIDE_ID", lang))
            return
        busy = await current_busy(lang)
        if busy:
            await interaction.followup.send(tr(
                lang,
                f"⏳ 지금은 다른 실행이 진행 중입니다. ({busy})\n"
                "custom 실행은 계정 정보 때문에 대기열에 넣지 않습니다. 끝난 뒤 다시 실행해 주세요.",
                f"⏳ Another run is in progress. ({busy})\n"
                "Custom runs are not queued because they carry account details. Please try again when it finishes.",
            ))
            return
        async with run_gate:
            r = next_reservation_within(RESERVE_GUARD)
            if r is not None:
                minutes = max(1, round((res_time(r) - datetime.now()).total_seconds() / 60))
                if not await confirm_followup(
                    interaction, reservation_soon_text(r, minutes, lang), tr(lang, "진행", "Proceed"), lang,
                ):
                    return
                await postpone(r, ran_first_text(interaction.user.display_name, r))
            message = await prepare_and_start(
                interaction.followup.send, stage,
                lambda: start_job(interaction.channel, stage, product, "run",
                                  tr(lang, f"BAT 실행 (워크스페이스 {workspace})", f"BAT run (workspace {workspace})"),
                                  override=override, options=options),
                override, lang,
            )
    except Exception:
        log.exception("BAT custom 실행 요청 실패")
        message = request_failed_text("Jenkins 실행", lang)
    await interaction.followup.send(tr(lang, f"{message}\n워크스페이스: `{workspace}` (사용자 지정)",
                                       f"{message}\nWorkspace: `{workspace}` (custom)"))


SHOW_CHOICES = [app_commands.Choice(name=L("all (everything incl. SKIP·UNVERIFIED·not run)",
                                           "all (SKIP·UNVERIFIED·미실행까지 전부)"), value="all")]
SHOW_DESCRIBE = L("Blank: show only PASS·FAIL·WARN·BAT_ERROR (message and Excel). all: show everything",
                  "비우면 PASS·FAIL·WARN·BAT_ERROR만 표시 (메시지·엑셀). all: 전부 표시")
STAGE_DESCRIBE = L("Target stage", "실행할 스테이지")


def unknown_product_text(product: str, lang: str) -> str:
    return tr(lang, f"알 수 없는 제품입니다: `{product}`\n목록에서 골라 주세요: {', '.join(PRODUCTS)}",
              f"Unknown product: `{product}`\nPlease pick one from the list: {', '.join(PRODUCTS)}")


@tree.command(name="bat", description=L("Run the Jenkins BAT test", "Jenkins BAT 테스트 실행"), guilds=TEST_GUILDS)
@app_commands.choices(stage=stage_choices, show=SHOW_CHOICES)
@app_commands.describe(
    stage=STAGE_DESCRIBE,
    product=L("Product to test (brackets: estimated time from recent runs)",
              "테스트할 제품 (괄호 안은 최근 실행 기준 예상 시간)"),
    custom=L("True: enter workspace, account and REST token in a form (only you see it)",
             "True: 워크스페이스·계정·REST 토큰을 입력창(본인만 보임)으로 직접 입력"),
    show=SHOW_DESCRIBE,
)
@app_commands.autocomplete(product=product_autocomplete)
@with_lang
async def bat(
    interaction: discord.Interaction,
    stage: app_commands.Choice[str],
    product: str,
    custom: bool = False,
    show: app_commands.Choice[str] | None = None,
    lang: app_commands.Choice[str] | None = None,
):
    lang = lang_for(interaction, lang)
    if product not in PRODUCTS:
        await interaction.response.send_message(unknown_product_text(product, lang), ephemeral=True)
        return
    show_value = show.value if show else None
    if custom:
        await interaction.response.send_modal(
            BatCustomModal(stage.value, product, {"show": show_value, "lang": lang}))
        return
    await interaction.response.defer(thinking=True)  # 채널 전체에 공개
    req = new_request(interaction, "bat", stage.value, product=product, show=show_value, lang=lang)
    try:
        message = await request_run(interaction, req)
    except Exception:
        log.exception("BAT 실행 요청 실패")
        message = request_failed_text("Jenkins 실행", lang)
    if message:
        await interaction.followup.send(message)


@tree.command(name="bat-doctor", description=L("Run the bat doctor check (auth, REST, log lookup)",
                                               "bat doctor 점검 실행 (인증·REST·로그 조회)"), guilds=TEST_GUILDS)
@app_commands.choices(stage=stage_choices)
@app_commands.describe(stage=STAGE_DESCRIBE)
@with_lang
async def bat_doctor(interaction: discord.Interaction, stage: app_commands.Choice[str],
                     lang: app_commands.Choice[str] | None = None):
    lang = lang_for(interaction, lang)
    await interaction.response.defer(thinking=True)
    if await current_busy(lang):
        await interaction.followup.send(busy_text(lang))
        return
    try:
        message = await start_job(interaction.channel, stage.value, "all", "doctor", "bat doctor",
                                  options={"lang": lang})
    except Exception:
        log.exception("bat doctor 요청 실패")
        message = request_failed_text("bat doctor", lang)
    await interaction.followup.send(message)


@tree.command(name="bat-sweep", description=L("Run bat sweep --delete (delete leftover BAT resources)",
                                              "bat sweep --delete 실행 (이전 BAT 리소스 삭제)"), guilds=TEST_GUILDS)
@app_commands.choices(stage=stage_choices)
@app_commands.describe(stage=STAGE_DESCRIBE)
@with_lang
async def bat_sweep(interaction: discord.Interaction, stage: app_commands.Choice[str],
                    lang: app_commands.Choice[str] | None = None):
    lang = lang_for(interaction, lang)
    confirmed = await ask_confirm(
        interaction,
        tr(lang,
           f"⚠️ **{stage.value}** 스테이지에서 `bat sweep --delete`를 실행할까요?\n"
           "이전 실행이 남긴 BAT 리소스를 **삭제**합니다.",
           f"⚠️ Run `bat sweep --delete` on the **{stage.value}** stage?\n"
           "This **deletes** BAT resources left by earlier runs."),
        tr(lang, "삭제 실행", "Delete"),
        lang,
    )
    if not confirmed:
        return
    if await current_busy(lang):
        await interaction.edit_original_response(content=busy_text(lang))
        return
    try:
        message = await start_job(interaction.channel, stage.value, "all", "sweep", "bat sweep",
                                  options={"lang": lang})
    except Exception:
        log.exception("bat sweep 요청 실패")
        message = request_failed_text("bat sweep", lang)
    await interaction.edit_original_response(content=f"{interaction.user.display_name}: {message}")


@tree.command(name="bat-cancel", description=L("Cancel a queued or running build", "대기 중이거나 실행 중인 Build 취소"),
              guilds=TEST_GUILDS)
@with_lang
async def bat_cancel(interaction: discord.Interaction, lang: app_commands.Choice[str] | None = None):
    lang = lang_for(interaction, lang)
    try:
        target = await find_active_target()
    except Exception:
        log.exception("취소 대상 조회 실패")
        await interaction.response.send_message(tr(
            lang, "Jenkins 상태 조회 실패. 봇 콘솔 로그를 확인해 주세요.",
            "Failed to get the Jenkins status. Please check the bot console log."))
        return
    if target is None:
        await interaction.response.send_message(tr(
            lang, "취소할 Build가 없습니다. (대기 중이거나 실행 중인 Build 없음)",
            "Nothing to cancel. (No queued or running build)"))
        return

    kind, info = target
    if kind == "queue":
        question = tr(lang, "⚠️ 대기열에 있는 실행 요청을 취소할까요?", "⚠️ Cancel the queued run request?")
    else:
        stage, product = get_build_parameters(info["build"])
        mode = get_build_mode(info["build"])
        what = MODE_LABELS.get(mode, tr(lang, f"BAT 실행 (PRODUCT: {product})", f"BAT run (PRODUCT: {product})"))
        question = tr(
            lang,
            f"⚠️ 실행 중인 **Build #{info['number']}**을 중단할까요?\n"
            f"{what} / STAGE: {stage}\n"
            "중간에 멈추면 워크스페이스에 테스트 리소스가 남을 수 있습니다. (다음 실행의 sweep이 정리)",
            f"⚠️ Stop the running **Build #{info['number']}**?\n"
            f"{what} / STAGE: {stage}\n"
            "Stopping midway may leave test resources in the workspace. (The next run's sweep cleans them up)",
        )
    if not await ask_confirm(interaction, question, tr(lang, "중단", "Stop"), lang):
        return

    try:
        # 확인을 기다리는 사이 상태가 바뀌었으면 엉뚱한 Build를 멈추지 않는다
        current = await find_active_target()
        same = current is not None and current[0] == kind and (
            current[1].get("id") == info.get("id") if kind == "queue"
            else current[1].get("number") == info.get("number")
        )
        if not same:
            result = tr(lang, "그 사이 상태가 바뀌어 취소하지 않았습니다. `/bat-status`로 확인해 주세요.",
                        "The state changed in the meantime, so nothing was cancelled. Check with `/bat-status`.")
        elif kind == "queue":
            await jenkins_post(f"{JENKINS_URL}/queue/cancelItem?id={info['id']}")
            await asyncio.sleep(1)
            still = await find_active_target()
            if still and still[0] == "queue" and still[1].get("id") == info["id"]:
                result = tr(lang, "대기열 취소 요청을 보냈지만 아직 남아 있습니다. Jenkins에서 확인해 주세요.",
                            "Sent the queue cancel request, but the item is still there. Please check Jenkins.")
            else:
                result = tr(lang, "대기열의 실행 요청을 취소했습니다.", "Cancelled the queued run request.")
        else:
            status = await jenkins_post(info["url"].rstrip("/") + "/stop")
            if status >= 400:
                result = tr(lang, f"중단 요청 실패 (HTTP {status}). Jenkins에서 확인해 주세요.",
                            f"Stop request failed (HTTP {status}). Please check Jenkins.")
            else:
                result = tr(lang, f"Build #{info['number']} 중단 요청을 보냈습니다. 멈추면 중단 알림이 전송됩니다.",
                            f"Sent a stop request for Build #{info['number']}. An aborted notice will follow when it stops.")
    except Exception:
        log.exception("취소 처리 실패")
        result = tr(lang, "취소 처리 중 오류. 봇 콘솔 로그를 확인해 주세요.",
                    "Error while cancelling. Please check the bot console log.")
    await interaction.edit_original_response(content=f"{interaction.user.display_name}: {result}")


# ============================================================
# /bat-result: Build의 결과 CSV에서 상태별 케이스와 사유 보기
# ============================================================
RESULT_STATUS_ORDER = ["PASS", "FAIL", "WARN", "BAT_ERROR", "UNVERIFIED", "SKIP"]
SHOW_ONLY_IF_PRESENT = ("WARN", "BAT_ERROR")  # 없으면 줄 자체를 생략
STATUS_HINTS = {  # 처음 보는 사람용 짧은 설명
    "FAIL": ("치명적 결함", "critical defect"),
    "WARN": ("기대와 다름, 확인 필요", "differs from expected, review"),
    "BAT_ERROR": ("테스트 도구·환경 문제, 제품 이슈 아님", "test tool/environment problem, not a product issue"),
    "UNVERIFIED": ("확인 못 함", "could not verify"),
}
REASON_MAX_CHARS = 60  # 사유 표시 길이 (넘으면 잘라서 … 표시)
# 짧은 사유 전용 열이 생기면 그 열을 우선 사용하고, 없으면 Actual Result 사용
REASON_COLUMNS = ["reason", "short reason", "failure reason", "사유", "원인", "actual result"]


def _decode_bytes(data: bytes, lang: str = DEFAULT_LANG) -> str:
    for encoding in ("utf-8-sig", "cp949"):
        try:
            return data.decode(encoding)
        except UnicodeDecodeError:
            continue
    raise ValueError(tr(lang, "CSV 인코딩을 확인할 수 없습니다", "Could not detect the CSV encoding"))


def short_reason(text: str, limit: int = REASON_MAX_CHARS, lang: str = DEFAULT_LANG) -> str:
    """여러 줄·공백을 한 줄로 합치고 limit자 안팎으로 자른다 (단어 중간은 피함)."""
    line = " ".join((text or "").split())
    if not line:
        return tr(lang, "사유 없음", "no reason")
    if len(line) <= limit:
        return line
    cut = line[:limit]
    space = cut.rfind(" ")
    if space >= int(limit * 0.6):
        cut = cut[:space]
    return cut.rstrip(" ,.(:") + "…"


def summarize_result_csv(data: bytes, lang: str = DEFAULT_LANG) -> dict[str, list[tuple[str, str]]]:
    """CSV → {상태: [(케이스 ID, 사유), ...]}"""
    rows = list(csv.DictReader(io.StringIO(_decode_bytes(data, lang), newline="")))
    if not rows:
        return {}
    headers = {name.strip().lower(): name for name in rows[0].keys() if name}
    case_col = headers.get("case id") or headers.get("case_id") or headers.get("id")
    status_col = headers.get("status") or headers.get("result") or headers.get("verdict")
    reason_col = next((headers[c] for c in REASON_COLUMNS if c in headers), None)
    if not case_col or not status_col:
        raise ValueError(tr(lang, "CSV에서 Case ID 또는 Status 열을 찾지 못했습니다",
                            "Could not find a Case ID or Status column in the CSV"))

    grouped: dict[str, list[tuple[str, str]]] = {}
    for row in rows:
        status = norm_status(row.get(status_col) or "")
        if not status:
            continue
        case_id = (row.get(case_col) or "").strip() or tr(lang, "(ID 없음)", "(no ID)")
        if reason_col:
            reason = short_reason(row.get(reason_col) or "", lang=lang)
        else:
            reason = tr(lang, "사유 열 없음", "no reason column")
        grouped.setdefault(status, []).append((case_id, reason))
    return grouped


def format_result_summary(
    grouped: dict[str, list[tuple[str, str]]], show_all: bool = False, lang: str = DEFAULT_LANG,
    gate: tuple[str, list[str]] | None = None,
) -> list[str]:
    """PASS/FAIL/WARN/BAT_ERROR를 표시한다. show_all이면 UNVERIFIED/SKIP도 케이스별로 표시한다.

    미실행 등 다른 상태는 표시하지 않는다. show_all이 아니면 숨긴 상태는 개수도 표시하지 않는다
    (완료 알림과 같음, 사용자 요청).
    """
    # gate는 새 BAT인지(FAIL = 치명) 구분하는 데만 쓴다. 롤백 판정 문구는 보여주지 않는다.
    lines: list[str] = []
    for status in RESULT_STATUS_ORDER:
        items = grouped.get(status, [])
        if status in SHOW_ONLY_IF_PRESENT and not items:
            continue
        if status in HIDDEN_BY_DEFAULT and not show_all:
            continue
        hint = STATUS_HINTS.get(status)
        if status == "FAIL" and not gate:
            hint = None  # 옛 BAT의 FAIL은 치명 기준이 아님
        suffix = f" ({tr(lang, *hint)})" if hint else ""
        lines.append(tr(lang, f"**{status}** {len(items)}개{suffix}", f"**{status}** {len(items)}{suffix}"))
        if status != "PASS":
            lines.extend(f"- {case_id}({reason})" for case_id, reason in items)
    return lines


def split_messages(lines: list[str], limit: int = 1900) -> list[str]:
    """Discord 메시지 길이 제한(2000자)에 맞게 나눈다."""
    chunks: list[str] = []
    current = ""
    for line in lines:
        if current and len(current) + len(line) + 1 > limit:
            chunks.append(current)
            current = ""
        current = f"{current}\n{line}" if current else line
    if current:
        chunks.append(current)
    return chunks


@tree.command(name="bat-result", description=L("Show a build's results by status, with cases and reasons",
                                               "Build 결과를 상태별 케이스와 사유로 보기"), guilds=TEST_GUILDS)
@app_commands.describe(build=L("Build number (blank: last completed build)", "Build 번호 (비우면 마지막 완료 Build)"),
                       show=L("Blank: only PASS·FAIL·WARN·BAT_ERROR. all: everything", "비우면 PASS·FAIL·WARN·BAT_ERROR만. all: 전부"),
                       file=L("True: also attach the result Excel, written in the lang language",
                              "True면 결과 엑셀도 첨부 (lang 언어로 작성)"))
@app_commands.choices(show=SHOW_CHOICES)
@with_lang
async def bat_result(interaction: discord.Interaction, build: int | None = None,
                     show: app_commands.Choice[str] | None = None,
                     file: bool = False,
                     lang: app_commands.Choice[str] | None = None):
    lang = lang_for(interaction, lang)
    await interaction.response.defer(thinking=True)  # 채널 전체에 공개
    try:
        if build is None:
            job = await jenkins_get(JOB_API_URL + "?tree=lastCompletedBuild[number]")
            build = (job.get("lastCompletedBuild") or {}).get("number")
            if not build:
                await interaction.followup.send(tr(lang, "완료된 Build가 없습니다.", "No completed builds yet."))
                return

        build_url = f"{JENKINS_URL}{JOB_PATH}/{build}/"
        try:
            info = await jenkins_get(build_url + BUILD_API_TREE)
        except RuntimeError:
            await interaction.followup.send(tr(lang, f"Build #{build}을(를) 찾을 수 없습니다.",
                                               f"Build #{build} not found."))
            return
        if info.get("building"):
            await interaction.followup.send(tr(lang, f"Build #{build}은(는) 아직 실행 중입니다.",
                                               f"Build #{build} is still running."))
            return
        if get_build_mode(info) in MODE_LABELS:
            mode = get_build_mode(info)
            await interaction.followup.send(tr(lang, f"Build #{build}은(는) 테스트 실행이 아닌 {mode} 작업입니다.",
                                               f"Build #{build} is a {mode} job, not a test run."))
            return

        try:
            console_text = await jenkins_text(build_url + "consoleText")
        except RuntimeError:
            console_text = ""
        contents, problems = await load_result_csvs(build_url, console_text, lang)
        if not contents:
            reason = problems[0] if problems else tr(lang, "결과 CSV를 찾지 못했습니다.", "Result CSV not found.")
            await interaction.followup.send(f"Build #{build}: {reason}")
            return

        grouped: dict[str, list[tuple[str, str]]] = {}
        for _, data in contents:
            for status, items in summarize_result_csv(data, lang).items():
                grouped.setdefault(status, []).extend(items)
    except ValueError as exc:
        await interaction.followup.send(f"Build #{build}: {exc}")
        return
    except Exception:
        log.exception("결과 조회 실패")
        await interaction.followup.send(tr(lang, "결과 조회 실패. 봇 콘솔 로그를 확인해 주세요.",
                                           "Failed to get the results. Please check the bot console log."))
        return

    stage, product = get_build_parameters(info)
    result = info.get('result') or '-'
    header = tr(lang, f"📊 **Build #{build} 결과** ({stage} / {product} / Jenkins {result})",
                f"📊 **Build #{build} results** ({stage} / {product} / Jenkins {result})")
    show_all = bool(show and show.value == "all")
    gate = parse_gate(console_text)
    lines = [header, *format_result_summary(grouped, show_all=show_all, lang=lang, gate=gate)]
    files: list[tuple[str, bytes]] = []
    if file:
        try:
            files, file_problems = await build_result_files(contents, lang, show_all)
        except Exception:
            log.exception("결과 파일 만들기 실패")
            files, file_problems = [], [tr(lang, "결과 파일을 만들지 못했습니다. 봇 콘솔 로그를 확인해 주세요.",
                                           "Could not build the result file. Please check the bot console log.")]
        lines += [f"⚠️ {p}" for p in file_problems]
    chunks = split_messages(lines)
    for index, chunk in enumerate(chunks):
        last = index == len(chunks) - 1
        await interaction.followup.send(
            chunk, files=[discord.File(io.BytesIO(data), filename=name) for name, data in files] if last and files else [])


# ============================================================
# /bat-result file:True: 결과 파일을 원하는 언어(lang)로 다시 만들어 첨부 (테스트 재실행 없음)
# ============================================================
# 결과 CSV에서 실행마다 달라지는 값은 상태뿐이고 나머지는 BAT의 고정 설명이라, relang_csv.py가
# BAT 자신의 결과 파일 작성 코드로 다른 언어 CSV를 다시 만든다 (BAT venv python으로 실행).
BAT_DIR = Path(os.environ.get("BAT_DIR", "").strip() or r"C:\test-automation\aioncloud-sse-bat")
RELANG_SCRIPT = Path(__file__).resolve().parent / "relang_csv.py"
RELANG_TIMEOUT_SEC = 60


def relang_csv(data: bytes, lang: str) -> bytes:
    """결과 CSV → lang 언어 CSV. 실패하면 RuntimeError(사유)."""
    import subprocess
    import tempfile

    python = BAT_DIR / ".venv" / "Scripts" / "python.exe"
    if not python.exists():
        raise RuntimeError(f"BAT venv not found: {python}")
    with tempfile.TemporaryDirectory() as tmp:
        src, out = Path(tmp) / "src.csv", Path(tmp) / "out.csv"
        src.write_bytes(data)
        done = subprocess.run(
            [str(python), str(RELANG_SCRIPT), "--bat-dir", str(BAT_DIR), "--src", str(src),
             "--lang", lang, "--out", str(out)],
            capture_output=True, timeout=RELANG_TIMEOUT_SEC,
            env={**os.environ, "PYTHONIOENCODING": "utf-8"},
        )
        if done.returncode != 0 or not out.exists():
            message = done.stderr.decode("utf-8", "replace").strip().splitlines()
            raise RuntimeError(message[-1] if message else f"exit {done.returncode}")
        return out.read_bytes()


async def build_result_files(contents: list[tuple[str, bytes]], lang: str,
                             show_all: bool) -> tuple[list[tuple[str, bytes]], list[str]]:
    """결과 CSV → lang 언어로 다시 쓰고(relang_csv), show 기준으로 거른 뒤 엑셀로 바꾼다."""
    problems: list[str] = []
    converted: list[tuple[str, bytes]] = []
    for name, data in contents:
        try:
            converted.append((name, await asyncio.to_thread(relang_csv, data, lang)))
        except Exception as exc:
            log.exception("결과 파일 언어 변환 실패: %s", name)
            problems.append(tr(lang, f"{name}: 언어 변환 실패 ({exc}), 원본 언어로 첨부",
                               f"{name}: language conversion failed ({exc}), attached in the original language"))
            converted.append((name, data))
    if not show_all:
        converted = [(name, await asyncio.to_thread(filter_csv_rows, data)) for name, data in converted]
    files, notes = await asyncio.to_thread(_convert_to_excel, converted, lang)
    return files, problems + notes


# ============================================================
# /bat-except add·delete·list: all 실행 제외 목록
# ============================================================
class ExcludeFileError(Exception):
    """제외 목록·사전 설정 파일 내용이 JSON 형식이 아님."""

    def __init__(self, path, lineno: int):
        super().__init__(f"{path} line {lineno}: invalid format")
        self.path = path
        self.lineno = lineno

    def text(self, lang: str) -> str:
        return tr(lang, f"{self.path} {self.lineno}번째 줄 형식 오류", f"{self.path} line {self.lineno}: invalid format")


def load_excludes() -> list[dict]:
    try:
        # utf-8-sig: PowerShell 등으로 저장해 BOM이 붙은 파일도 읽는다
        text = BAT_EXCLUDE_FILE.read_text(encoding="utf-8-sig")
    except FileNotFoundError:
        return []
    if not text.strip():
        return []  # 빈 파일은 빈 목록으로 본다
    try:
        data = json.loads(text)
    except json.JSONDecodeError as exc:
        raise ExcludeFileError(BAT_EXCLUDE_FILE, exc.lineno) from exc
    items = data.get("items", []) if isinstance(data, dict) else []
    return [item for item in items if isinstance(item, dict) and item.get("id")]


def save_excludes(items: list[dict]) -> None:
    BAT_EXCLUDE_FILE.parent.mkdir(parents=True, exist_ok=True)
    tmp = BAT_EXCLUDE_FILE.with_suffix(".tmp")
    tmp.write_text(json.dumps({"items": items}, ensure_ascii=False, indent=2), encoding="utf-8")
    tmp.replace(BAT_EXCLUDE_FILE)


def normalize_exclude_id(value: str) -> str | None:
    """제품 이름(all 제외) 또는 케이스 ID(BAT-SWG-001 형식)만 허용한다."""
    text = value.strip()
    if text.lower() in PRODUCTS and text.lower() != "all":
        return text.lower()
    if CASE_ID_PATTERN.match(text.upper()):
        return text.upper()
    return None


def exclude_broken_text(exc: ExcludeFileError, lang: str, changed: bool = True) -> str:
    if changed:
        return tr(lang, f"제외 목록 파일이 손상되어 수정하지 않았습니다: {exc.text(lang)}",
                  f"The exclusion list file is broken, so nothing was changed: {exc.text(lang)}")
    return tr(lang, f"제외 목록 파일이 손상되었습니다: {exc.text(lang)}",
              f"The exclusion list file is broken: {exc.text(lang)}")


def exclude_save_failed_text(lang: str) -> str:
    return tr(lang, "제외 목록 저장 실패. 봇 콘솔 로그를 확인해 주세요.",
              "Failed to save the exclusion list. Please check the bot console log.")


except_group = app_commands.Group(name="bat-except", description=L("Manage the all-run exclusion list",
                                                                   "all 실행 제외 목록 관리"))


@except_group.command(name="add", description=L("Add a product or case ID to exclude from all runs",
                                                "all 실행에서 제외할 제품 또는 케이스 ID 추가"))
@app_commands.describe(item=L("Product name (e.g. dlp) or case ID (e.g. BAT-SWG-001)",
                              "제품 이름(예: dlp) 또는 케이스 ID(예: BAT-SWG-001)"))
@with_lang
async def except_add(interaction: discord.Interaction, item: str, lang: app_commands.Choice[str] | None = None):
    lang = lang_for(interaction, lang)
    target = normalize_exclude_id(item)
    if target is None:
        await interaction.response.send_message(tr(
            lang, f"`{item}`은(는) 제품 이름이나 케이스 ID 형식(BAT-XXX-000)이 아닙니다.",
            f"`{item}` is not a product name or a case ID (BAT-XXX-000)."), ephemeral=True
        )
        return
    async with exclude_lock:
        try:
            items = await asyncio.to_thread(load_excludes)
            if any(entry["id"] == target for entry in items):
                await interaction.response.send_message(tr(
                    lang, f"`{target}`은(는) 이미 제외 목록에 있습니다.", f"`{target}` is already excluded."))
                return
            items.append({
                "id": target,
                "by": interaction.user.display_name,
                "at": datetime.now().astimezone().strftime("%Y-%m-%d %H:%M"),
            })
            await asyncio.to_thread(save_excludes, items)
        except ExcludeFileError as exc:
            await interaction.response.send_message(exclude_broken_text(exc, lang))
            return
        except Exception:
            log.exception("제외 목록 저장 실패")
            await interaction.response.send_message(exclude_save_failed_text(lang))
            return
    await interaction.response.send_message(tr(
        lang, f"🚫 제외 목록에 `{target}` 추가. 다음 all 실행부터 적용됩니다. (현재 {len(items)}개)",
        f"🚫 Added `{target}` to the exclusion list. Applies from the next all run. (now {len(items)})"
    ))


@except_group.command(name="delete", description=L("Remove an item from the exclusion list", "제외 목록에서 항목 삭제"))
@app_commands.describe(item=L("Item to remove", "삭제할 항목"))
@with_lang
async def except_remove(interaction: discord.Interaction, item: str, lang: app_commands.Choice[str] | None = None):
    lang = lang_for(interaction, lang)
    target = normalize_exclude_id(item) or item.strip()
    async with exclude_lock:
        try:
            items = await asyncio.to_thread(load_excludes)
            remaining = [entry for entry in items if entry["id"] != target]
            if len(remaining) == len(items):
                await interaction.response.send_message(tr(
                    lang, f"`{target}`은(는) 제외 목록에 없습니다.", f"`{target}` is not in the exclusion list."))
                return
            await asyncio.to_thread(save_excludes, remaining)
        except ExcludeFileError as exc:
            await interaction.response.send_message(exclude_broken_text(exc, lang))
            return
        except Exception:
            log.exception("제외 목록 저장 실패")
            await interaction.response.send_message(exclude_save_failed_text(lang))
            return
    await interaction.response.send_message(tr(
        lang, f"✅ 제외 목록에서 `{target}` 제거. 다음 all 실행부터 다시 실행됩니다. (현재 {len(remaining)}개)",
        f"✅ Removed `{target}` from the exclusion list. It runs again from the next all run. (now {len(remaining)})"
    ))


@except_remove.autocomplete("item")
async def except_remove_autocomplete(interaction: discord.Interaction, current: str):
    try:
        items = await asyncio.to_thread(load_excludes)
    except Exception:
        return []
    return [
        app_commands.Choice(name=entry["id"], value=entry["id"])
        for entry in items
        if current.lower() in entry["id"].lower()
    ][:25]


@except_add.autocomplete("item")
async def except_add_autocomplete(interaction: discord.Interaction, current: str):
    return [
        app_commands.Choice(name=product, value=product)
        for product in PRODUCTS
        if product != "all" and current.lower() in product
    ][:25]


@except_group.command(name="list", description=L("Show the current all-run exclusion list", "현재 all 실행 제외 목록 보기"))
@with_lang
async def except_list(interaction: discord.Interaction, lang: app_commands.Choice[str] | None = None):
    lang = lang_for(interaction, lang)
    try:
        items = await asyncio.to_thread(load_excludes)
    except ExcludeFileError as exc:
        await interaction.response.send_message(exclude_broken_text(exc, lang, changed=False))
        return
    except Exception:
        log.exception("제외 목록 읽기 실패")
        await interaction.response.send_message(tr(
            lang, "제외 목록을 읽지 못했습니다. 봇 콘솔 로그를 확인해 주세요.",
            "Could not read the exclusion list. Please check the bot console log."))
        return
    if not items:
        await interaction.response.send_message(tr(
            lang, "제외 목록이 비어 있습니다. all 실행 시 전체 테스트가 실행됩니다.",
            "The exclusion list is empty. An all run runs every test."))
        return
    lines = [f"• `{e['id']}` — {e.get('by', '-')} ({e.get('at', '-')})" for e in items]
    await interaction.response.send_message(
        tr(lang, "🚫 **all 실행 제외 목록**\n", "🚫 **All-run exclusion list**\n") + "\n".join(lines)[:1900]
    )


tree.add_command(except_group, guilds=TEST_GUILDS)


# ============================================================
# 개인 사전 설정: /bat-preset add(만들기·수정)·list·delete, /bat-plan <이름>(실행)
# 사람마다 여러 개, 사전 설정마다 STAGE와 제품 목록을 고정
# ============================================================
MY_PLAN_FILE = BOT_DATA_DIR / "my-plans.json"
PLAN_PRODUCTS = [product for product in PRODUCTS if product != "all"]
PLAN_NAME_MAX = 30
MAX_PLANS_PER_USER = 25
my_plan_lock = asyncio.Lock()


def load_my_plans() -> dict:
    """{사용자 ID: {사전 설정 이름: {stage, products, updated_at}}}"""
    try:
        text = MY_PLAN_FILE.read_text(encoding="utf-8-sig")
    except FileNotFoundError:
        return {}
    if not text.strip():
        return {}
    try:
        data = json.loads(text)
    except json.JSONDecodeError as exc:
        raise ExcludeFileError(MY_PLAN_FILE, exc.lineno) from exc
    users = data.get("users", {}) if isinstance(data, dict) else {}
    return users if isinstance(users, dict) else {}


def save_my_plans(users: dict) -> None:
    tmp = MY_PLAN_FILE.with_suffix(".tmp")
    tmp.write_text(json.dumps({"users": users}, ensure_ascii=False, indent=2), encoding="utf-8")
    tmp.replace(MY_PLAN_FILE)


def normalize_plan_name(name: str) -> str | None:
    text = " ".join(name.split())
    if not text or len(text) > PLAN_NAME_MAX:
        return None
    return text


def describe_plan(plan: dict) -> str:
    return f"{plan.get('stage', '-')}: {', '.join(plan.get('products', [])) or '-'}"


def plan_broken_text(exc: ExcludeFileError, lang: str, changed: bool | None = None) -> str:
    """changed: None=읽기만, True=수정 안 함, 'save'=저장 안 함."""
    if changed == "save":
        return tr(lang, f"사전 설정 파일이 손상되어 저장하지 않았습니다: {exc.text(lang)}",
                  f"The preset file is broken, so nothing was saved: {exc.text(lang)}")
    if changed:
        return tr(lang, f"사전 설정 파일이 손상되어 수정하지 않았습니다: {exc.text(lang)}",
                  f"The preset file is broken, so nothing was changed: {exc.text(lang)}")
    return tr(lang, f"사전 설정 파일이 손상되었습니다: {exc.text(lang)}", f"The preset file is broken: {exc.text(lang)}")


async def own_plan_autocomplete(interaction: discord.Interaction, current: str):
    try:
        users = await asyncio.to_thread(load_my_plans)
    except Exception:
        return []
    plans = users.get(str(interaction.user.id), {})
    return [
        app_commands.Choice(
            name=f"{name} [{plan.get('stage', '-')} - {', '.join(plan.get('products', [])) or '-'}]"[:100],
            value=name,
        )
        for name, plan in plans.items()
        if current.lower() in name.lower()
    ][:25]


class PlanEditorView(discord.ui.View):
    """스테이지 1개 + 제품 여러 개를 고르고 저장하는 편집 화면 (본인만 보임)."""

    def __init__(self, owner_id: int, name: str, stage: str | None, products: list[str], lang: str = DEFAULT_LANG):
        super().__init__(timeout=300)
        self.owner_id = owner_id
        self.name = name
        self.lang = lang
        self.stage = stage if stage in STAGES else None
        self.products = [p for p in PLAN_PRODUCTS if p in products]
        self.message: discord.InteractionMessage | None = None

        self.stage_select = discord.ui.Select(
            placeholder=tr(lang, "① 스테이지 선택", "① Choose a stage"),
            min_values=1,
            max_values=1,
            options=[discord.SelectOption(label=s, value=s, default=(s == self.stage)) for s in STAGES],
            row=0,
        )
        self.stage_select.callback = self.on_stage
        self.product_select = discord.ui.Select(
            placeholder=tr(lang, "② 제품 선택 (여러 개 가능)", "② Choose products (multiple allowed)"),
            min_values=1,
            max_values=len(PLAN_PRODUCTS),
            options=[
                discord.SelectOption(label=p, value=p, default=(p in self.products))
                for p in PLAN_PRODUCTS
            ],
            row=1,
        )
        self.product_select.callback = self.on_products
        self.add_item(self.stage_select)
        self.add_item(self.product_select)
        self.save_button.label = tr(lang, "저장", "Save")
        self.cancel_button.label = tr(lang, "취소", "Cancel")

    def summary(self) -> str:
        lang = self.lang
        none = tr(lang, "선택 안 함", "not chosen")
        return tr(
            lang,
            f"📝 사전 설정 **{self.name}** 편집\n"
            f"STAGE: `{self.stage or none}`\n"
            f"PRODUCT: `{', '.join(self.products) or none}`\n"
            "선택 후 [저장]을 눌러 주세요. (5분 이내)",
            f"📝 Editing preset **{self.name}**\n"
            f"STAGE: `{self.stage or none}`\n"
            f"PRODUCT: `{', '.join(self.products) or none}`\n"
            "Choose, then press [Save]. (within 5 minutes)",
        )

    async def interaction_check(self, interaction: discord.Interaction) -> bool:
        if interaction.user.id != self.owner_id:
            await interaction.response.send_message(
                tr(self.lang, "본인 사전 설정만 편집할 수 있습니다.", "You can only edit your own presets."),
                ephemeral=True)
            return False
        return True

    async def on_stage(self, interaction: discord.Interaction):
        self.stage = self.stage_select.values[0]
        for option in self.stage_select.options:
            option.default = option.value == self.stage
        await interaction.response.edit_message(content=self.summary(), view=self)

    async def on_products(self, interaction: discord.Interaction):
        chosen = set(self.product_select.values)
        self.products = [p for p in PLAN_PRODUCTS if p in chosen]
        for option in self.product_select.options:
            option.default = option.value in chosen
        await interaction.response.edit_message(content=self.summary(), view=self)

    @discord.ui.button(label="Save", style=discord.ButtonStyle.success, row=2)
    async def save_button(self, interaction: discord.Interaction, button: discord.ui.Button):
        lang = self.lang
        if not self.stage or not self.products:
            await interaction.response.send_message(
                tr(lang, "스테이지와 제품을 모두 선택해 주세요.", "Please choose both a stage and products."),
                ephemeral=True)
            return
        try:
            async with my_plan_lock:
                users = await asyncio.to_thread(load_my_plans)
                plans = users.setdefault(str(self.owner_id), {})
                if self.name not in plans and len(plans) >= MAX_PLANS_PER_USER:
                    await interaction.response.send_message(tr(
                        lang,
                        f"사전 설정은 한 사람당 {MAX_PLANS_PER_USER}개까지입니다. `/bat-preset delete`로 정리해 주세요.",
                        f"Up to {MAX_PLANS_PER_USER} presets per person. Clean up with `/bat-preset delete`.",
                    ), ephemeral=True)
                    return
                plans[self.name] = {
                    "stage": self.stage,
                    "products": self.products,
                    "updated_at": datetime.now().astimezone().strftime("%Y-%m-%d %H:%M"),
                }
                await asyncio.to_thread(save_my_plans, users)
        except ExcludeFileError as exc:
            await interaction.response.send_message(plan_broken_text(exc, lang, "save"), ephemeral=True)
            return
        except Exception:
            log.exception("사전 설정 저장 실패")
            await interaction.response.send_message(tr(
                lang, "사전 설정 저장 실패. 봇 콘솔 로그를 확인해 주세요.",
                "Failed to save the preset. Please check the bot console log."), ephemeral=True)
            return
        self.stop()
        await interaction.response.edit_message(
            content=tr(
                lang,
                f"✅ 사전 설정 **{self.name}** 저장\n"
                f"STAGE: `{self.stage}`\nPRODUCT: `{', '.join(self.products)}`\n"
                f"실행: `/bat-plan run name:{self.name}`",
                f"✅ Saved preset **{self.name}**\n"
                f"STAGE: `{self.stage}`\nPRODUCT: `{', '.join(self.products)}`\n"
                f"Run: `/bat-plan run name:{self.name}`",
            ),
            view=None,
        )

    @discord.ui.button(label="Cancel", style=discord.ButtonStyle.secondary, row=2)
    async def cancel_button(self, interaction: discord.Interaction, button: discord.ui.Button):
        self.stop()
        await interaction.response.edit_message(
            content=tr(self.lang, "사전 설정 편집을 취소했습니다. (저장 안 됨)", "Preset editing cancelled. (not saved)"),
            view=None)

    async def on_timeout(self):
        if self.message is not None:
            try:
                await self.message.edit(
                    content=tr(self.lang, "시간이 지나 편집을 닫았습니다. (저장 안 됨)", "Editing timed out. (not saved)"),
                    view=None)
            except Exception:
                pass


my_plan_group = app_commands.Group(name="bat-preset", description=L("Manage my presets", "내 사전 설정 관리"))


@my_plan_group.command(name="add", description=L("Create or edit my preset (choose stage and products)",
                                                 "내 사전 설정 만들기·수정 (스테이지와 제품 선택)"))
@app_commands.describe(name=L("Preset name (an existing name is edited)", "사전 설정 이름 (있는 이름이면 수정)"))
@app_commands.autocomplete(name=own_plan_autocomplete)
@with_lang
async def my_plan(interaction: discord.Interaction, name: str, lang: app_commands.Choice[str] | None = None):
    lang = lang_for(interaction, lang)
    plan_name = normalize_plan_name(name)
    if plan_name is None:
        await interaction.response.send_message(tr(
            lang, f"사전 설정 이름은 1~{PLAN_NAME_MAX}자로 입력해 주세요.",
            f"Preset names must be 1-{PLAN_NAME_MAX} characters."), ephemeral=True)
        return
    try:
        users = await asyncio.to_thread(load_my_plans)
    except ExcludeFileError as exc:
        await interaction.response.send_message(plan_broken_text(exc, lang), ephemeral=True)
        return
    existing = users.get(str(interaction.user.id), {}).get(plan_name) or {}
    view = PlanEditorView(interaction.user.id, plan_name, existing.get("stage"), existing.get("products", []), lang)
    await interaction.response.send_message(view.summary(), view=view, ephemeral=True)
    view.message = await interaction.original_response()


@my_plan_group.command(name="list", description=L("Show my presets", "내 사전 설정 목록 보기"))
@with_lang
async def my_plan_list(interaction: discord.Interaction, lang: app_commands.Choice[str] | None = None):
    lang = lang_for(interaction, lang)
    try:
        users = await asyncio.to_thread(load_my_plans)
    except ExcludeFileError as exc:
        await interaction.response.send_message(plan_broken_text(exc, lang), ephemeral=True)
        return
    plans = users.get(str(interaction.user.id), {})
    if not plans:
        await interaction.response.send_message(tr(
            lang, "저장된 사전 설정이 없습니다. `/bat-preset add`로 만들어 주세요.",
            "No saved presets. Create one with `/bat-preset add`."), ephemeral=True)
        return
    lines = [
        f"• **{name}** — {describe_plan(plan)} ({plan.get('updated_at', '-')})"
        for name, plan in plans.items()
    ]
    await interaction.response.send_message(
        tr(lang, "📋 **내 사전 설정** (`/bat-plan run`으로 실행)\n", "📋 **My presets** (run with `/bat-plan run`)\n")
        + "\n".join(lines)[:1900], ephemeral=True
    )


@my_plan_group.command(name="delete", description=L("Delete my preset", "내 사전 설정 삭제"))
@app_commands.describe(name=L("Preset to delete", "삭제할 사전 설정"))
@app_commands.autocomplete(name=own_plan_autocomplete)
@with_lang
async def my_plan_delete(interaction: discord.Interaction, name: str, lang: app_commands.Choice[str] | None = None):
    lang = lang_for(interaction, lang)
    plan_name = normalize_plan_name(name) or name
    try:
        async with my_plan_lock:
            users = await asyncio.to_thread(load_my_plans)
            plans = users.get(str(interaction.user.id), {})
            if plan_name not in plans:
                await interaction.response.send_message(tr(
                    lang, f"`{plan_name}` 사전 설정이 없습니다.", f"No preset named `{plan_name}`."), ephemeral=True)
                return
            del plans[plan_name]
            await asyncio.to_thread(save_my_plans, users)
    except ExcludeFileError as exc:
        await interaction.response.send_message(plan_broken_text(exc, lang, True), ephemeral=True)
        return
    except Exception:
        log.exception("사전 설정 삭제 실패")
        await interaction.response.send_message(tr(
            lang, "사전 설정 삭제 실패. 봇 콘솔 로그를 확인해 주세요.",
            "Failed to delete the preset. Please check the bot console log."), ephemeral=True)
        return
    await interaction.response.send_message(tr(lang, f"🗑️ 사전 설정 **{plan_name}** 삭제",
                                               f"🗑️ Deleted preset **{plan_name}**"), ephemeral=True)


tree.add_command(my_plan_group, guilds=TEST_GUILDS)


# ============================================================
# 실행 대기열·예약: /bat·/bat-plan 동시 실행 막기, /bat-plan run·list·cancel, 업데이트 block(DM)
# 대기열과 예약은 bat-schedule.json에 저장해 봇을 다시 켜도 남는다.
# ============================================================
SCHEDULE_FILE = BOT_DATA_DIR / "bat-schedule.json"
SCHEDULE_INTERVAL_SECONDS = 20
RESERVE_GUARD = timedelta(minutes=10)  # 예약 10분 전부터는 다른 실행을 확인받고, 10분 전에 바쁘면 알린다
RESERVE_CHECK = timedelta(minutes=1)  # 1분 전에도 바쁘면 미룬다
POSTPONE_STEP = timedelta(minutes=10)
MAX_POSTPONE = 6
RESERVE_MAX_DAYS = 7
TIME_FORMAT = "%Y-%m-%d %H:%M"
WEEKDAYS = "월화수목금토일"
WEEKDAYS_EN = ("Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun")

run_gate = asyncio.Lock()  # 실행 요청 하나가 확인·커넥터 준비·Jenkins 요청까지 끝낼 때까지 잡는다
schedule: dict = {"reservations": [], "queue": []}
schedule_task: asyncio.Task | None = None


def load_schedule() -> dict:
    try:
        data = json.loads(SCHEDULE_FILE.read_text(encoding="utf-8"))
        if isinstance(data, dict):
            return {"reservations": list(data.get("reservations") or []), "queue": list(data.get("queue") or []),
                    "block_until": data.get("block_until"), "block_by": data.get("block_by")}
    except FileNotFoundError:
        pass
    except Exception:
        log.exception("예약 파일을 읽지 못함 (빈 상태로 시작)")
    return {"reservations": [], "queue": []}


def save_schedule() -> None:
    tmp = SCHEDULE_FILE.with_suffix(".tmp")
    tmp.write_text(json.dumps(schedule, ensure_ascii=False, indent=2), encoding="utf-8")
    tmp.replace(SCHEDULE_FILE)


def new_request(interaction: discord.Interaction, kind: str, stage: str, **extra) -> dict:
    """extra: product/products/plan, show, lang (대기열·예약 실행과 알림에 쓰는 응답 언어)."""
    return {
        "id": os.urandom(3).hex(), "kind": kind, "stage": stage,
        "user_id": interaction.user.id, "user_name": interaction.user.display_name,
        "channel_id": interaction.channel.id if interaction.channel else None,
        **extra,
    }


def describe_req(req: dict, lang: str = DEFAULT_LANG) -> str:
    if req["kind"] == "plan":
        products = ", ".join(req["products"])
        return tr(lang, f"**{req['plan']}** 사전 설정 · STAGE `{req['stage']}` · PRODUCT `{products}`",
                  f"**{req['plan']}** preset · STAGE `{req['stage']}` · PRODUCT `{products}`")
    return tr(lang, f"BAT 실행 · STAGE `{req['stage']}` · PRODUCT `{req['product']}`",
              f"BAT run · STAGE `{req['stage']}` · PRODUCT `{req['product']}`")


def req_options(req: dict) -> dict:
    """요청에 저장된 실행 옵션 중 완료 알림에 쓰는 것."""
    return {"show": req.get("show"), "lang": req_lang(req)}


def start_req(channel, req: dict):
    lang = req_lang(req)
    if req["kind"] == "plan":
        return start_job(channel, req["stage"], "all", "run",
                         tr(lang, f"BAT 사전 설정 실행 ({req['plan']})", f"BAT preset run ({req['plan']})"),
                         products=req["products"], options=req_options(req))
    return start_job(channel, req["stage"], req["product"], "run", tr(lang, "BAT 실행", "BAT run"),
                     options=req_options(req))


def res_time(item: dict) -> datetime:
    return datetime.strptime(item["at"], TIME_FORMAT)


def show_time(at: datetime, lang: str = DEFAULT_LANG) -> str:
    if lang == "ko":
        return f"{at:%m/%d}({WEEKDAYS[at.weekday()]}) {at:%H:%M}"
    return f"{at:%m/%d}({WEEKDAYS_EN[at.weekday()]}) {at:%H:%M}"


async def reserve_plan(req: dict, at: datetime, lang: str = DEFAULT_LANG) -> str:
    now = datetime.now()
    if at < now + timedelta(minutes=2):
        return tr(lang, f"{show_time(at, lang)}은(는) 이미 지났거나 너무 가깝습니다. 2분 이후 시각을 골라 주세요.",
                  f"{show_time(at, lang)} has passed or is too close. Please pick a time at least 2 minutes ahead.")
    if at > now + timedelta(days=RESERVE_MAX_DAYS):
        return tr(lang, f"예약은 {RESERVE_MAX_DAYS}일 이내로만 할 수 있습니다.",
                  f"Reservations can only be made within {RESERVE_MAX_DAYS} days.")
    schedule["reservations"].append({**req, "at": at.strftime(TIME_FORMAT), "postponed": 0, "warned": False})
    save_schedule()
    # 예약 정보와 안내를 나눠 보이게: 안내는 빈 줄 뒤 작은 회색 글씨(-#)
    products = ", ".join(req["products"])
    return tr(
        lang,
        "⏰ **예약 완료**\n"
        f"사전 설정: **{req['plan']}**\n"
        f"STAGE: `{req['stage']}`\n"
        f"PRODUCT: `{products}`\n"
        f"실행 시각: **{show_time(at, lang)}**\n"
        f"예약자: {req['user_name']}\n"
        "\n"
        "-# 예약 시각 1분 전에 다른 실행이 진행 중이면 10분씩 미뤄집니다. (최대 6번, 이후 취소)",
        "⏰ **Reserved**\n"
        f"Preset: **{req['plan']}**\n"
        f"STAGE: `{req['stage']}`\n"
        f"PRODUCT: `{products}`\n"
        f"Run at: **{show_time(at, lang)}**\n"
        f"Reserved by: {req['user_name']}\n"
        "\n"
        "-# If another run is in progress 1 minute before, it is postponed by 10 minutes. (up to 6 times, then cancelled)",
    )


class ReserveView(discord.ui.View):
    """날짜·시·분을 따로 골라 사전 설정 실행을 예약하는 선택창 (본인만 보임)."""

    def __init__(self, owner_id: int, req: dict, lang: str = DEFAULT_LANG):
        super().__init__(timeout=180)
        self.owner_id = owner_id
        self.req = req
        self.lang = lang
        self.date: str | None = None
        self.hour: str | None = None
        self.minute: str | None = None
        self.message: discord.InteractionMessage | None = None
        today = datetime.now().date()
        dates = [today + timedelta(days=i) for i in range(RESERVE_MAX_DAYS)]

        def date_label(i: int, d) -> str:
            if lang == "ko":
                return f"{d:%m/%d}({WEEKDAYS[d.weekday()]})" + (" 오늘" if i == 0 else " 내일" if i == 1 else "")
            return f"{d:%m/%d}({WEEKDAYS_EN[d.weekday()]})" + (" today" if i == 0 else " tomorrow" if i == 1 else "")

        self.date_select = discord.ui.Select(placeholder=tr(lang, "날짜", "Date"), row=0, options=[
            discord.SelectOption(label=date_label(i, d), value=d.isoformat()) for i, d in enumerate(dates)
        ])
        self.hour_select = discord.ui.Select(placeholder=tr(lang, "시", "Hour"), row=1, options=[
            discord.SelectOption(label=tr(lang, f"{h:02d}시", f"{h:02d}h"), value=f"{h:02d}") for h in range(24)
        ])
        self.minute_select = discord.ui.Select(placeholder=tr(lang, "분", "Minute"), row=2, options=[
            discord.SelectOption(label=tr(lang, f"{m:02d}분", f":{m:02d}"), value=f"{m:02d}") for m in range(0, 60, 10)
        ])
        for select, attr in ((self.date_select, "date"), (self.hour_select, "hour"), (self.minute_select, "minute")):
            select.callback = self._picker(select, attr)
            self.add_item(select)
        self.reserve_button.label = tr(lang, "예약", "Reserve")
        self.cancel_button.label = tr(lang, "취소", "Cancel")

    def summary(self) -> str:
        lang = self.lang
        if self.date and self.hour and self.minute:
            picked = show_time(datetime.strptime(f"{self.date} {self.hour}:{self.minute}", TIME_FORMAT), lang)
        else:
            picked = " / ".join(v or label for v, label in (
                (self.date, tr(lang, "날짜", "date")),
                (self.hour and tr(lang, f"{self.hour}시", f"{self.hour}h"), tr(lang, "시", "hour")),
                (self.minute and tr(lang, f"{self.minute}분", f":{self.minute}"), tr(lang, "분", "minute")),
            ))
        return tr(
            lang,
            f"⏰ 예약 시각 선택\n{describe_req(self.req, lang)}\n"
            f"선택: **{picked}**\n고른 뒤 [예약]을 눌러 주세요. (3분 이내)",
            f"⏰ Choose a time to reserve\n{describe_req(self.req, lang)}\n"
            f"Selected: **{picked}**\nThen press [Reserve]. (within 3 minutes)",
        )

    def _picker(self, select: discord.ui.Select, attr: str):
        async def callback(interaction: discord.Interaction):
            setattr(self, attr, select.values[0])
            for option in select.options:  # 고른 값이 선택창에 계속 보이도록
                option.default = option.value == select.values[0]
            await interaction.response.edit_message(content=self.summary(), view=self)
        return callback

    async def interaction_check(self, interaction: discord.Interaction) -> bool:
        if interaction.user.id != self.owner_id:
            await interaction.response.send_message(
                tr(self.lang, "예약한 분만 고를 수 있습니다.", "Only the person reserving can choose."), ephemeral=True)
            return False
        return True

    @discord.ui.button(label="Reserve", style=discord.ButtonStyle.primary, row=3)
    async def reserve_button(self, interaction: discord.Interaction, button: discord.ui.Button):
        lang = self.lang
        if not (self.date and self.hour and self.minute):
            await interaction.response.send_message(
                tr(lang, "날짜·시·분을 모두 골라 주세요.", "Please choose the date, hour and minute."), ephemeral=True)
            return
        at = datetime.strptime(f"{self.date} {self.hour}:{self.minute}", TIME_FORMAT)
        message = await reserve_plan(self.req, at, lang)
        if not message.startswith("⏰"):
            await interaction.response.send_message(message, ephemeral=True)
            return
        self.stop()
        await interaction.response.edit_message(
            content=tr(lang, f"예약 완료 : {show_time(at, lang)}", f"Reserved: {show_time(at, lang)}"), view=None)
        await interaction.followup.send(message)  # 채널 전체에 공개

    @discord.ui.button(label="Cancel", style=discord.ButtonStyle.secondary, row=3)
    async def cancel_button(self, interaction: discord.Interaction, button: discord.ui.Button):
        self.stop()
        await interaction.response.edit_message(
            content=tr(self.lang, "예약을 취소했습니다. (예약 안 됨)", "Reservation cancelled. (not reserved)"), view=None)

    async def on_timeout(self):
        if self.message is not None:
            try:
                await self.message.edit(
                    content=tr(self.lang, "시간이 지나 예약 창을 닫았습니다. (예약 안 됨)",
                               "The reservation window timed out. (not reserved)"), view=None)
            except discord.HTTPException:
                pass


def next_reservation_within(window: timedelta) -> dict | None:
    now = datetime.now()
    upcoming = [r for r in schedule["reservations"] if res_time(r) - now <= window]
    return min(upcoming, key=res_time) if upcoming else None


async def current_busy(lang: str = DEFAULT_LANG) -> str | None:
    """테스트·doctor·sweep·커넥터 작업 중 하나라도 진행 중이면 그 이유."""
    if run_gate.locked():
        return tr(lang, "다른 실행 요청을 처리하는 중입니다.", "Another run request is being processed.")
    if connector_lock.locked():
        return tr(lang, "커넥터 작업 중입니다.", "A connector task is in progress.")
    return await bat_busy_reason(lang)


async def channel_by_id(channel_id):
    if not channel_id:
        return None
    channel = client.get_channel(int(channel_id))
    if channel is None:
        try:
            channel = await client.fetch_channel(int(channel_id))
        except discord.HTTPException:
            log.exception("채널 조회 실패: %s", channel_id)
            return None
    return channel


async def notify(item: dict, text: str) -> None:
    channel = await channel_by_id(item.get("channel_id"))
    if channel is not None:
        await channel.send(text)


async def postpone(r: dict, reason: str) -> None:
    """예약을 10분 미룬다. 6번을 넘기면 취소한다. 미룬 뒤에는 새 시각 1분 전에만 다시 확인한다.

    reason은 예약 항목의 언어(req_lang(r))로 만들어 넘긴다.
    """
    lang = req_lang(r)
    r["postponed"] = int(r.get("postponed", 0)) + 1
    if r["postponed"] > MAX_POSTPONE:
        if r in schedule["reservations"]:
            schedule["reservations"].remove(r)
        save_schedule()
        await notify(r, tr(
            lang,
            f"❌ <@{r['user_id']}> 예약 취소: {describe_req(r, lang)}\n"
            f"{MAX_POSTPONE}번(60분) 미룬 뒤에도 실행할 수 없었습니다. ({reason})",
            f"❌ <@{r['user_id']}> Reservation cancelled: {describe_req(r, lang)}\n"
            f"It still could not run after being postponed {MAX_POSTPONE} times (60 min). ({reason})",
        ))
        return
    at = res_time(r) + POSTPONE_STEP
    r["at"] = at.strftime(TIME_FORMAT)
    r["warned"] = True
    save_schedule()
    await notify(r, tr(
        lang,
        f"⏳ <@{r['user_id']}> 예약을 10분 미뤘습니다 → **{show_time(at, lang)}** ({r['postponed']}/{MAX_POSTPONE})\n"
        f"{describe_req(r, lang)}\n이유: {reason}",
        f"⏳ <@{r['user_id']}> Reservation postponed by 10 minutes → **{show_time(at, lang)}** "
        f"({r['postponed']}/{MAX_POSTPONE})\n"
        f"{describe_req(r, lang)}\nReason: {reason}",
    ))


async def confirm_followup(interaction: discord.Interaction, question: str, confirm_label: str,
                           lang: str = DEFAULT_LANG) -> bool:
    view = ConfirmView(interaction.user.id, confirm_label, lang)
    message = await interaction.followup.send(question, view=view, wait=True)
    await view.wait()
    if view.confirmed is None:
        await message.edit(content=f"{question}\n\n{no_response_text(lang)}", view=None)
        return False
    return view.confirmed


async def request_run(interaction: discord.Interaction, req: dict) -> str | None:
    """/bat·/bat-plan 바로 실행 요청. 실행 중이면 대기열 등록을 묻고, 예약 10분 전이면 진행 여부를 묻는다.

    보낼 결과 문구를 돌려준다 (확인 질문으로 이미 답했으면 None).
    """
    lang = req_lang(req)
    busy = await current_busy(lang)
    if busy:
        ahead = len(schedule["queue"])
        if not await confirm_followup(
            interaction,
            tr(lang,
               f"⏳ 지금은 다른 실행이 진행 중이라 바로 실행할 수 없습니다. ({busy})\n"
               f"대기열에 등록해 드릴까요? 앞 작업이 끝나면 자동으로 실행합니다. (현재 대기 {ahead}건)\n"
               f"{describe_req(req, lang)}",
               f"⏳ Another run is in progress, so this cannot start now. ({busy})\n"
               f"Add it to the queue? It runs automatically when the earlier work finishes. ({ahead} waiting)\n"
               f"{describe_req(req, lang)}"),
            tr(lang, "대기열 등록", "Add to queue"),
            lang,
        ):
            return None
        schedule["queue"].append(req)
        save_schedule()
        position = len(schedule["queue"])
        return tr(lang,
                  f"📥 대기열 등록 완료 ({position}번째): {describe_req(req, lang)}\n"
                  "앞 작업이 끝나면 자동으로 실행하고 이 채널에 알려 드립니다. 취소: `/bat-plan cancel`",
                  f"📥 Added to the queue (#{position}): {describe_req(req, lang)}\n"
                  "It runs automatically when the earlier work finishes and is announced here. "
                  "Cancel: `/bat-plan cancel`")
    async with run_gate:
        r = next_reservation_within(RESERVE_GUARD)
        if r is not None:
            minutes = max(1, round((res_time(r) - datetime.now()).total_seconds() / 60))
            if not await confirm_followup(
                interaction, reservation_soon_text(r, minutes, lang), tr(lang, "진행", "Proceed"), lang,
            ):
                return None
            await postpone(r, ran_first_text(interaction.user.display_name, r))
        return await prepare_and_start(interaction.followup.send, req["stage"],
                                       lambda: start_req(interaction.channel, req), None, lang)


async def run_scheduled(item: dict, intro: tuple[str, str]) -> None:
    """예약·대기열 항목을 실행한다. 시작 직전에 다시 바쁘면 예약은 미루고 대기열은 맨 앞에 되돌린다.

    intro: (한국어, 영어) 시작 안내. 항목의 언어로 보낸다.
    """
    lang = req_lang(item)
    channel = await channel_by_id(item.get("channel_id"))
    if channel is None:
        log.warning("예약·대기열 실행 건너뜀: 채널 없음 (%s)", item.get("id"))
        return
    async with run_gate:
        busy = await bat_busy_reason(lang)
        if busy:
            if "at" in item:
                schedule["reservations"].append(item)
                await postpone(item, busy)
            else:
                schedule["queue"].insert(0, item)
                save_schedule()
            return
        await channel.send(f"{tr(lang, *intro)}\n{describe_req(item, lang)}")
        try:
            message = await prepare_and_start(channel.send, item["stage"], lambda: start_req(channel, item),
                                              None, lang)
        except Exception:
            log.exception("예약·대기열 실행 실패")
            message = request_failed_text("Jenkins 실행", lang)
        await channel.send(message)


async def schedule_tick() -> None:
    now = datetime.now()
    await expire_block(now)
    for r in sorted(schedule["reservations"], key=res_time):
        at = res_time(r)
        if now < at - RESERVE_GUARD or (now < at - RESERVE_CHECK and r.get("warned")):
            continue
        lang = req_lang(r)
        busy = await tick_busy(lang)
        if now < at - RESERVE_CHECK:  # 10분 전: 바쁘면 알림만
            r["warned"] = True
            save_schedule()
            if busy:
                await notify(r, tr(
                    lang,
                    f"⏰ <@{r['user_id']}> **{r['plan']}** 사전 설정 실행 예약 10분 전입니다 ({show_time(at, lang)}).\n"
                    f"지금 다른 실행이 진행 중이라 1분 전에 다시 확인하고, 그때도 진행 중이면 10분 미룹니다. ({busy})",
                    f"⏰ <@{r['user_id']}> Your **{r['plan']}** preset run is reserved in 10 minutes "
                    f"({show_time(at, lang)}).\n"
                    "Another run is in progress, so it will be checked again 1 minute before "
                    f"and postponed by 10 minutes if still busy. ({busy})",
                ))
            continue
        if busy:  # 1분 전부터: 바쁘면 미룬다
            await postpone(r, busy)
            return
        if now >= at:
            schedule["reservations"].remove(r)
            save_schedule()
            asyncio.create_task(run_scheduled(r, (
                f"⏰ <@{r['user_id']}> 예약한 사전 설정 실행을 시작합니다.",
                f"⏰ <@{r['user_id']}> Starting your reserved preset run.",
            )))
            return
    # 대기열: 곧 실행할 예약이 없고 비어 있을 때 하나씩
    if schedule["queue"] and next_reservation_within(RESERVE_GUARD) is None and not await tick_busy():
        item = schedule["queue"].pop(0)
        save_schedule()
        asyncio.create_task(run_scheduled(item, (
            f"📤 <@{item['user_id']}> 대기열에 등록한 실행을 시작합니다.",
            f"📤 <@{item['user_id']}> Starting your queued run.",
        )))


# ---- 업데이트 차단: 봇에게 DM으로 'block 10'(10분) / 'block off'(해제) / 'block'(상태) ----
# 명령어 목록에 나오지 않도록 슬래시 명령이 아니라 DM으로 받는다. 누구나 쓸 수 있고(아는 사람만 씀),
# 차단을 건 사람은 차단 중에도 명령을 쓸 수 있다 (업데이트한 것을 직접 테스트하도록).
# DM에는 채널 언어가 없으므로 기본 언어(en)로 답한다.
BLOCK_MAX_MINUTES = 720


def block_help(lang: str) -> str:
    return tr(lang, "사용법: `block 10` (10분 차단) · `block off` (해제) · `block` (상태)",
              "Usage: `block 10` (block for 10 min) · `block off` (release) · `block` (status)")


def blocked_until() -> datetime | None:
    value = schedule.get("block_until")
    if not value:
        return None
    try:
        until = datetime.strptime(value, TIME_FORMAT + ":%S")
    except ValueError:
        return None
    return until if until > datetime.now() else None


def block_text(until: datetime, lang: str = DEFAULT_LANG) -> str:
    return tr(lang, f"🛠️ {until:%H:%M}까지 업데이트 중입니다. 잠시만 기다려주세요.",
              f"🛠️ Updating until {until:%H:%M}. Please wait a moment.")


async def tick_busy(lang: str = DEFAULT_LANG) -> str | None:
    until = blocked_until()
    if until:
        return tr(lang, f"{until:%H:%M}까지 업데이트 중", f"updating until {until:%H:%M}")
    return await current_busy(lang)


async def dm_blocker(text: str) -> None:
    uid = schedule.get("block_by")
    if not uid:
        return
    try:
        await (await client.fetch_user(int(uid))).send(text)
    except discord.HTTPException:
        log.exception("block DM 전송 실패")


async def expire_block(now: datetime) -> None:
    value = schedule.get("block_until")
    if value and blocked_until() is None:
        await dm_blocker(tr(DEFAULT_LANG, f"🔓 block 종료 (시간 만료, {now:%H:%M})",
                            f"🔓 block ended (time expired, {now:%H:%M})"))
        schedule["block_until"] = None
        schedule["block_by"] = None
        save_schedule()


async def handle_block_dm(message: discord.Message) -> None:
    m = re.fullmatch(r"block(?:\s+(\S+))?", message.content.strip(), re.IGNORECASE)
    if not m:  # block이 아닌 DM에는 반응하지 않는다 (기능이 드러나지 않게)
        return
    lang = channel_lang(message.channel.id)
    arg = (m.group(1) or "").lower()
    now = datetime.now()
    if not arg:
        until = blocked_until()
        if until:
            left_min = int((until - now).total_seconds() // 60) + 1
            left = tr(lang, f"{until:%H:%M}까지 (남은 {left_min}분)", f"until {until:%H:%M} ({left_min} min left)")
        else:
            left = tr(lang, "차단 안 함", "not blocked")
        await message.channel.send(tr(lang, f"block 상태: {left}", f"block status: {left}") + f"\n{block_help(lang)}")
        return
    if arg in ("off", "cancel", "해제", "취소"):
        was = blocked_until()
        schedule["block_until"] = None
        schedule["block_by"] = None
        save_schedule()
        await message.channel.send(tr(lang, "🔓 block 해제", "🔓 block released") if was
                                   else tr(lang, "차단 중이 아닙니다.", "Not blocked."))
        return
    if not arg.isdigit() or not 1 <= int(arg) <= BLOCK_MAX_MINUTES:
        await message.channel.send(tr(lang, f"분은 1~{BLOCK_MAX_MINUTES} 사이 숫자로 넣어 주세요.",
                                      f"Minutes must be a number from 1 to {BLOCK_MAX_MINUTES}.")
                                   + f"\n{block_help(lang)}")
        return
    until = now + timedelta(minutes=int(arg))
    schedule["block_until"] = until.strftime(TIME_FORMAT + ":%S")
    schedule["block_by"] = message.author.id
    save_schedule()
    await message.channel.send(tr(
        lang,
        f"🔒 block 설정: **{until:%H:%M}까지** ({arg}분)\n"
        f"그동안 다른 사람의 명령에는 '{block_text(until, lang)}'가 나가고, 예약·대기열 실행도 멈춥니다. "
        "본인 명령은 그대로 됩니다.\n해제: `block off`",
        f"🔒 block set: **until {until:%H:%M}** ({arg} min)\n"
        f"Meanwhile other people's commands get '{block_text(until, lang)}', and reserved/queued runs pause. "
        "Your own commands still work.\nRelease: `block off`",
    ))


@client.event
async def on_message(message: discord.Message):
    # 서버 채널 메시지는 보지 않는다 (DM만)
    if message.guild is not None or message.author.bot:
        return
    try:
        await handle_block_dm(message)
    except Exception:
        log.exception("block DM 처리 실패")


async def schedule_loop() -> None:
    await client.wait_until_ready()
    while True:
        try:
            await schedule_tick()
        except asyncio.CancelledError:
            raise
        except Exception:
            log.exception("예약·대기열 처리 오류 (다음 주기에 재시도)")
        await asyncio.sleep(SCHEDULE_INTERVAL_SECONDS)


bat_plan_group = app_commands.Group(
    name="bat-plan",
    description=L("Run my preset (run), list reservations (list), cancel (cancel)",
                  "내 사전 설정 실행(run)·예약 목록(list)·예약 취소(cancel)"),
)


@bat_plan_group.command(name="list", description=L("Show reservations and the queue", "예약과 대기열 보기"))
@with_lang
async def schedule_list(interaction: discord.Interaction, lang: app_commands.Choice[str] | None = None):
    """예약은 시각순, 대기열은 실행 순서대로 항목마다 칸을 나눠 보여 준다."""
    lang = lang_for(interaction, lang)
    reservations = sorted(schedule["reservations"], key=res_time)
    embed = discord.Embed(title=tr(lang, "📅 예약 · 대기열", "📅 Reservations · Queue"), color=0x5865F2)
    fields = []
    for r in reservations:
        moved = tr(lang, f" · 미룸 {r['postponed']}/{MAX_POSTPONE}",
                   f" · postponed {r['postponed']}/{MAX_POSTPONE}") if r.get("postponed") else ""
        fields.append((f"⏰ {show_time(res_time(r), lang)}",
                       tr(lang, f"{req_summary(r, lang)}\n예약자 {r['user_name']}{moved}",
                          f"{req_summary(r, lang)}\nReserved by {r['user_name']}{moved}")))
    for i, q in enumerate(schedule["queue"], 1):
        fields.append((tr(lang, f"📥 대기 {i}번", f"📥 Queue #{i}"),
                       tr(lang, f"{req_summary(q, lang)}\n요청자 {q['user_name']}",
                          f"{req_summary(q, lang)}\nRequested by {q['user_name']}")))
    if not fields:
        embed.description = tr(lang, "예약과 대기열이 비어 있습니다.", "No reservations and the queue is empty.")
    for name, value in fields[:25]:  # 임베드 칸은 25개까지
        embed.add_field(name=name, value=value, inline=False)
    queued = len(schedule["queue"])
    more = len(fields) > 25
    embed.set_footer(text=tr(lang, f"예약 {len(reservations)}건 · 대기 {queued}건" + (" · 25건까지만 표시" if more else ""),
                             f"{len(reservations)} reserved · {queued} queued" + (" · showing first 25" if more else "")))
    await interaction.response.send_message(embed=embed)


def req_summary(req: dict, lang: str = DEFAULT_LANG) -> str:
    """목록 칸 안에 넣을 두 줄: 무엇을 / 어디에."""
    if req["kind"] == "plan":
        title = tr(lang, f"사전 설정 **{req['plan']}**", f"Preset **{req['plan']}**")
    else:
        title = tr(lang, "**BAT 실행**", "**BAT run**")
    product = ", ".join(req["products"]) if req["kind"] == "plan" else req["product"]
    return f"{title}\nSTAGE `{req['stage']}` · PRODUCT `{product}`"


def req_short(req: dict, lang: str = DEFAULT_LANG) -> str:
    """취소 목록용 한 줄: stset [qa - basic]"""
    product = ", ".join(req["products"]) if req["kind"] == "plan" else req["product"]
    return f"{req.get('plan') or tr(lang, 'BAT 실행', 'BAT run')} [{req['stage']} - {product}]"


async def own_schedule_autocomplete(interaction: discord.Interaction, current: str):
    lang = lang_for(interaction, getattr(interaction.namespace, "lang", None))
    items = [(f"⏰ {show_time(res_time(r), lang)} · {req_short(r, lang)}", r)
             for r in sorted(schedule["reservations"], key=res_time)]
    items += [(tr(lang, f"📥 대기 {i}번 · {req_short(q, lang)}", f"📥 Queue #{i} · {req_short(q, lang)}"), q)
              for i, q in enumerate(schedule["queue"], 1)]
    return [
        app_commands.Choice(name=label[:100], value=item["id"])
        for label, item in items
        if item["user_id"] == interaction.user.id and current.lower() in label.lower()
    ][:25]


@bat_plan_group.command(name="cancel", description=L("Cancel my reservation or queued item", "내 예약 또는 대기열 항목 취소"))
@app_commands.describe(item=L("Item to cancel", "취소할 항목"))
@app_commands.autocomplete(item=own_schedule_autocomplete)
@with_lang
async def schedule_cancel(interaction: discord.Interaction, item: str, lang: app_commands.Choice[str] | None = None):
    lang = lang_for(interaction, lang)
    for key, label in (("reservations", tr(lang, "예약", "Reservation")), ("queue", tr(lang, "대기열", "Queue item"))):
        for entry in schedule[key]:
            if entry["id"] != item:
                continue
            if entry["user_id"] != interaction.user.id:
                await interaction.response.send_message(tr(
                    lang, "본인이 등록한 항목만 취소할 수 있습니다.", "You can only cancel items you added."),
                    ephemeral=True)
                return
            schedule[key].remove(entry)
            save_schedule()
            product = ", ".join(entry["products"]) if entry["kind"] == "plan" else entry["product"]
            if entry["kind"] == "plan":
                what = tr(lang, f"사전 설정: **{entry['plan']}**", f"Preset: **{entry['plan']}**")
            else:
                what = tr(lang, "BAT 실행", "BAT run")
            lines = [tr(lang, f"🗑️ **{label} 취소**", f"🗑️ **{label} cancelled**"),
                     what,
                     f"STAGE: `{entry['stage']}`",
                     f"PRODUCT: `{product}`"]
            if key == "reservations":
                lines.append(tr(lang, f"실행 시각: {show_time(res_time(entry), lang)}",
                                f"Run at: {show_time(res_time(entry), lang)}"))
            await interaction.response.send_message("\n".join(lines))
            return
    await interaction.response.send_message(tr(
        lang, "해당 항목이 없습니다. (이미 실행됐거나 취소됨)", "No such item. (already run or cancelled)"),
        ephemeral=True)


@bat_plan_group.command(name="run", description=L(
    "Run my preset (stage and products as saved; turn on time-schedule to reserve)",
    "내 사전 설정 실행 (스테이지·제품은 사전 설정에 저장된 값, time-schedule을 켜면 예약)"))
@app_commands.describe(name=L("Preset to run", "실행할 사전 설정"),
                       reserve=L("True: open a date/time picker to reserve (blank: run now)",
                                 "True: 날짜·시간 선택창을 열어 예약 (비우면 바로 실행)"),
                       show=SHOW_DESCRIBE)
@app_commands.rename(reserve="time-schedule")  # 옵션 이름은 소문자·띄어쓰기 없이만 가능
@app_commands.choices(show=SHOW_CHOICES)
@app_commands.autocomplete(name=own_plan_autocomplete)
@with_lang
async def bat_plan(interaction: discord.Interaction, name: str, reserve: bool = False,
                   show: app_commands.Choice[str] | None = None,
                   lang: app_commands.Choice[str] | None = None):
    lang = lang_for(interaction, lang)
    plan_name = normalize_plan_name(name) or name
    message = None
    req = None
    try:
        users = await asyncio.to_thread(load_my_plans)
        plan = users.get(str(interaction.user.id), {}).get(plan_name)
        if not plan:
            message = tr(lang, f"`{plan_name}` 사전 설정이 없습니다. `/bat-preset add`로 먼저 만들어 주세요.",
                         f"No preset named `{plan_name}`. Create it first with `/bat-preset add`.")
        else:
            stage = plan.get("stage")
            products = [p for p in PLAN_PRODUCTS if p in (plan.get("products") or [])]
            if stage not in STAGES or not products:
                message = tr(lang, f"`{plan_name}` 사전 설정 내용이 올바르지 않습니다. `/bat-preset add`로 다시 저장해 주세요.",
                             f"The `{plan_name}` preset is invalid. Save it again with `/bat-preset add`.")
            else:
                req = new_request(interaction, "plan", stage, products=products, plan=plan_name,
                                  show=show.value if show else None, lang=lang)
    except ExcludeFileError as exc:
        message = plan_broken_text(exc, lang)

    if reserve:  # 선택창은 본인에게만 보이고, 예약이 끝나면 결과를 채널에 공개한다
        if req is None:
            await interaction.response.send_message(message, ephemeral=True)
            return
        view = ReserveView(interaction.user.id, req, lang)
        await interaction.response.send_message(view.summary(), view=view, ephemeral=True)
        view.message = await interaction.original_response()
        return

    await interaction.response.defer(thinking=True)  # 채널 전체에 공개
    if req is not None:
        try:
            message = await request_run(interaction, req)
        except Exception:
            log.exception("BAT 사전 설정 실행 요청 실패")
            message = request_failed_text("Jenkins 실행", lang)
    if message:
        await interaction.followup.send(message)


tree.add_command(bat_plan_group, guilds=TEST_GUILDS)


# ============================================================
# /connector on·logout·status: U+SASE·AIConnector 커넥터 제어
# 실제 조작은 connector.py가 별도 프로세스로 한다 (봇이 멈추지 않도록).
# ============================================================
CONNECTOR_SCRIPT = Path(__file__).with_name("connector.py")
CONNECTOR_ALLOWED = {
    int(x) for x in os.environ.get("CONNECTOR_ALLOWED_USER_IDS", "").split(",") if x.strip()
}
connector_lock = asyncio.Lock()


def busy_text(lang: str = DEFAULT_LANG) -> str:
    return tr(lang, "⏳ 지금은 다른 실행이 진행 중이라 실행할 수 없습니다.",
              "⏳ Another run is in progress, so this cannot run now.")


# connector.py가 돌려주는 RESULT 문구(한국어) 중 자주 나오는 고정 표현의 영어. 그 밖의 문구는 그대로 보인다.
CONNECTOR_RESULT_EN = [
    ("서버에 입력 화면이 없어(RDP 최소화·연결 끊김) 실제 마우스 조작이 필요한 단계를 진행하지 못했습니다. "
     "RDP를 연결한 상태에서 다시 시도해 주세요.",
     "The server has no input desktop (RDP minimized or disconnected), so steps that need real mouse input "
     "could not run. Please retry with RDP connected."),
    ("보안 연결 ON 완료", "secure connection ON "),
    ("이미 로그아웃 상태", "already logged out"),
    ("로그아웃 완료", "logged out"),
    ("커넥터 설정 없음: ", "no connector settings: "),
    ("예상치 못한 오류 (", "Unexpected error ("),
    (". 봇 콘솔 로그를 확인해 주세요.", ". Please check the bot console log."),
    ("마지막 로그인: ", "Last login: "),
    ("앱: ", "App: "),
    ("화면: ", "Screen: "),
    ("상태: ", "Status: "),
    ("알 수 없음", "unknown"),
    ("로그인 화면", "login screen"),
    ("로그인됨", "logged in"),
]


def connector_result_text(result: str, lang: str) -> str:
    """connector.py 결과를 보여 줄 언어로 (영어는 알려진 표현만 바꾼다)."""
    if lang == "ko":
        return result
    for ko, en in CONNECTOR_RESULT_EN:
        result = result.replace(ko, en)
    return result


async def run_connector(args: list, override: dict | None = None, lang: str = DEFAULT_LANG) -> str:
    """계정 정보는 명령줄이 아니라 표준입력으로 넘긴다 (프로세스 목록·로그에 남지 않도록)."""
    proc = await asyncio.create_subprocess_exec(
        sys.executable, str(CONNECTOR_SCRIPT), *args,
        stdin=asyncio.subprocess.PIPE,
        stdout=asyncio.subprocess.PIPE,
        stderr=asyncio.subprocess.STDOUT,
        cwd=str(CONNECTOR_SCRIPT.parent),
        env={**os.environ, "PYTHONUTF8": "1"},
    )
    payload = (json.dumps(override or {}) + "\n").encode("utf-8")
    try:
        # 앱 전환·스테이지 변경·로그인에 재시도(앱 재시작·재로그인)까지 겹치면 5분을 넘을 수 있다
        out, _ = await asyncio.wait_for(proc.communicate(payload), timeout=480)
    except asyncio.TimeoutError:
        proc.kill()
        await proc.wait()
        return tr(lang, "⏱️ 8분 안에 끝나지 않아 중단했습니다. 서버 화면을 확인해 주세요.",
                  "⏱️ Stopped because it did not finish within 8 minutes. Please check the server screen.")
    finally:
        payload = b""
    text = out.decode("utf-8", errors="replace")
    log.info("connector.py %s 출력:\n%s", args[0], text[-3000:])
    for line in reversed(text.splitlines()):
        if line.startswith("RESULT "):
            return line[len("RESULT "):]
    return tr(lang, "❌ 커넥터 스크립트가 결과 없이 끝났습니다. 봇 콘솔 로그를 확인해 주세요.",
              "❌ The connector script ended without a result. Please check the bot console log.")


async def bat_busy_reason(lang: str = DEFAULT_LANG) -> str | None:
    """테스트를 시작할 수 없는 이유. 실행 중인 테스트가 있을 때 커넥터 스테이지를 바꾸면 그 테스트가 깨진다."""
    if monitor_task is not None and not monitor_task.done():
        return monitoring_text(lang)
    active = await get_active_state(await get_job(), lang)
    if active:
        return rejected_text(active, lang)
    return None


async def prepare_and_start(send, stage: str, start, override: dict | None = None, lang: str = DEFAULT_LANG) -> str:
    """테스트 전에 커넥터를 요청한 스테이지·기본 계정으로 맞추고(이미 맞으면 바로 통과) 테스트를 시작한다.

    앱이 다르면 종료 후 전환, 다른 스테이지·계정이면 로그아웃 후 로그인, 연결이 꺼져 있으면 다시 연결까지
    connector.py on이 처리한다. 커넥터 준비에 실패하면 테스트는 실행하지 않는다.
    send: 진행 메시지를 보낼 함수 (interaction.followup.send 또는 channel.send)
    """
    if connector_lock.locked():
        return tr(lang, f"{busy_text(lang)} (테스트는 실행하지 않았습니다)", f"{busy_text(lang)} (test not started)")
    target = f"{stage} / {override['workspace']}" if override else stage
    await send(tr(lang, f"🔌 **{target}** 커넥터 확인 중... (앱 전환·로그인까지 1~3분)",
                  f"🔌 Checking the **{target}** connector... (1-3 min including app switch and login)"))
    connector_override = (
        {key: override.get(key, "") for key in ("workspace", "user", "password")} if override else None
    )
    async with connector_lock:
        result = await run_connector(["on", stage, "--stdin"], connector_override, lang)
    shown = connector_result_text(result, lang)
    if result.startswith("⏭️"):
        note = tr(lang, f"🔌 {shown} → 커넥터 확인 없이 테스트 진행", f"🔌 {shown} → running the test without a connector check")
    elif result.endswith("✅"):
        note = tr(lang, f"🔌 커넥터: {shown}", f"🔌 Connector: {shown}")
    else:
        return tr(lang, f"🔌 커넥터 준비 실패로 테스트를 실행하지 않았습니다.\n{shown}",
                  f"🔌 The connector could not be prepared, so the test was not started.\n{shown}")
    return f"{note}\n{await start()}"


async def _connector_allowed(interaction: discord.Interaction, lang: str) -> bool:
    if CONNECTOR_ALLOWED and interaction.user.id not in CONNECTOR_ALLOWED:
        await interaction.response.send_message(
            tr(lang, "이 명령어를 사용할 권한이 없습니다.", "You do not have permission to use this command."),
            ephemeral=True)
        return False
    if connector_lock.locked():
        await interaction.response.send_message(busy_text(lang), ephemeral=True)
        return False
    return True


async def _connector_run(interaction: discord.Interaction, args: list, title: str, override: dict | None = None,
                         lang: str = DEFAULT_LANG):
    await interaction.response.defer(thinking=True)
    if args[0] != "status" and await current_busy(lang):  # 진행 중인 테스트의 스테이지·로그인을 바꾸지 않도록
        await interaction.followup.send(busy_text(lang))
        return
    async with connector_lock:
        result = await run_connector(args, override, lang)
    await interaction.followup.send(f"{title}: {connector_result_text(result, lang)}")


class ConnectorLoginModal(discord.ui.Modal, title="Connector custom login"):
    # 입력칸은 언어에 맞는 이름으로 __init__에서 만든다

    def __init__(self, stage: str, lang: str = DEFAULT_LANG):
        super().__init__(title=tr(lang, "커넥터 사용자 지정 로그인", "Connector custom login"))
        self.stage = stage
        self.lang = lang
        self.workspace = discord.ui.TextInput(
            label=tr(lang, "워크스페이스", "Workspace"), required=False,
            placeholder=tr(lang, "비우면 environments.yaml 값 사용", "Blank: use the value in environments.yaml"))
        self.email = discord.ui.TextInput(label=tr(lang, "이메일", "Email"), required=True)
        self.password = discord.ui.TextInput(label=tr(lang, "비밀번호", "Password"), required=True)
        for item in (self.workspace, self.email, self.password):
            self.add_item(item)

    async def on_submit(self, interaction: discord.Interaction):
        lang = self.lang
        if connector_lock.locked():
            await interaction.response.send_message(busy_text(lang), ephemeral=True)
            return
        override = {
            "workspace": self.workspace.value.strip(),
            "user": self.email.value.strip(),
            "password": self.password.value,
        }
        try:
            await _connector_run(interaction, ["on", self.stage, "--stdin"],
                                 tr(lang, f"**{self.stage}** 커넥터 ON (사용자 지정 계정)",
                                    f"**{self.stage}** connector ON (custom account)"), override, lang)
        finally:
            override.clear()


connector_group = app_commands.Group(name="connector", description=L("Control the U+SASE·AIConnector connector",
                                                                     "U+SASE·AIConnector 커넥터 제어"))


@connector_group.command(name="on", description=L("Match app and stage, log in, then turn the secure connection ON",
                                                  "앱·스테이지를 맞추고 로그인한 뒤 보안 연결 ON"))
@app_commands.describe(stage=STAGE_DESCRIBE,
                       custom=L("True: enter workspace, email and password in a form (only you see it)",
                                "True: 워크스페이스·이메일·비밀번호를 입력창(본인만 보임)으로 직접 입력"))
@app_commands.choices(stage=stage_choices)
@with_lang
async def connector_on(interaction: discord.Interaction, stage: app_commands.Choice[str], custom: bool = False,
                       lang: app_commands.Choice[str] | None = None):
    lang = lang_for(interaction, lang)
    if not await _connector_allowed(interaction, lang):
        return
    if custom:
        await interaction.response.send_modal(ConnectorLoginModal(stage.value, lang))
        return
    await _connector_run(interaction, ["on", stage.value, "--stdin"],
                         tr(lang, f"**{stage.value}** 커넥터 ON (기본 계정)", f"**{stage.value}** connector ON (default account)"),
                         lang=lang)


@connector_group.command(name="logout", description=L("Log out of the connector", "커넥터 로그아웃"))
@with_lang
async def connector_logout(interaction: discord.Interaction, lang: app_commands.Choice[str] | None = None):
    lang = lang_for(interaction, lang)
    if not await _connector_allowed(interaction, lang):
        return
    await _connector_run(interaction, ["logout"], tr(lang, "커넥터 로그아웃", "Connector logout"), lang=lang)


@connector_group.command(name="status", description=L("Current connector app, login and connection status",
                                                      "현재 커넥터 앱·로그인·연결 상태"))
@with_lang
async def connector_status(interaction: discord.Interaction, lang: app_commands.Choice[str] | None = None):
    lang = lang_for(interaction, lang)
    if not await _connector_allowed(interaction, lang):
        return
    await _connector_run(interaction, ["status"], tr(lang, "커넥터 상태", "Connector status"), lang=lang)


tree.add_command(connector_group, guilds=TEST_GUILDS)


# ---- /help: 카테고리별 명령어 안내 + 가이드 문서 링크 버튼 ----
HELP_STAGE = f"stage: {' · '.join(STAGES)}"
# 섹션마다 ((제목 ko, en), [(사용법, (설명 ko, en), [옵션 값...]), ...]). [ ]는 비워도 되는 옵션.
HELP_SECTIONS = [
    (("▶️ 테스트 실행", "▶️ Test runs"), [
        ("/bat stage product [custom] [show]",
         ("커넥터 로그인·연결 후 BAT 테스트 실행 · custom:True면 워크스페이스·계정·토큰 직접 입력 (대기열·예약 불가) · "
          "show:all이면 결과 알림·엑셀에 SKIP·UNVERIFIED까지 전부 (기본은 PASS·FAIL·WARN·BAT_ERROR)",
          "Log in and connect the connector, then run the BAT test · custom:True to enter workspace, account and "
          "token yourself (no queue or reservation) · show:all shows everything incl. SKIP·UNVERIFIED in the "
          "result notice and Excel (default: PASS·FAIL·WARN·BAT_ERROR)"),
         [HELP_STAGE, f"product: {' · '.join(PRODUCTS)}"]),
        ("/bat-status", ("Jenkins 작업 상태", "Jenkins job status"), []),
        ("/jenkins", ("Jenkins 바로가기 링크", "Link to Jenkins"), []),
        ("/bat-result [build] [show] [file]", ("Build 결과 (비우면 마지막 Build) · show:all이면 SKIP·UNVERIFIED까지 · "
                                               "file:True면 결과 엑셀도 첨부 (lang 언어로, 테스트 재실행 없음)",
                                               "Build results (blank: last build) · show:all includes SKIP·UNVERIFIED · "
                                               "file:True also attaches the result Excel (in the lang language, no re-run)"), []),
        ("/bat-cancel", ("대기·실행 중인 Build 취소", "Cancel a queued or running build"), []),
    ]),
    (("📋 내 사전 설정", "📋 My presets"), [
        ("/bat-preset add name", ("사전 설정 만들기·수정 (스테이지·제품 선택)", "Create or edit a preset (choose stage and products)"), []),
        ("/bat-preset list", ("내 사전 설정 목록", "List my presets"), []),
        ("/bat-preset delete name", ("사전 설정 삭제", "Delete a preset"), []),
    ]),
    (("📅 사전 설정 실행·예약", "📅 Preset runs · reservations"), [
        ("/bat-plan run name [time-schedule] [show]",
         ("사전 설정 실행 · time-schedule:True면 날짜·시간 골라 예약 · show:all이면 전부 표시",
          "Run a preset · time-schedule:True to pick a date and time to reserve · show:all shows everything"), []),
        ("/bat-plan list", ("예약·대기열 보기", "Show reservations and the queue"), []),
        ("/bat-plan cancel item", ("내 예약·대기열 취소", "Cancel my reservation or queued item"), []),
    ]),
    (("🔌 커넥터", "🔌 Connector"), [
        ("/connector on stage [custom]", ("로그인 후 보안 연결 ON (1~3분) · custom:True면 계정 직접 입력",
                                          "Log in and turn the secure connection ON (1-3 min) · custom:True to enter the account yourself"),
         [HELP_STAGE]),
        ("/connector status", ("앱·로그인·연결 상태", "App, login and connection status"), []),
        ("/connector logout", ("로그아웃", "Log out"), []),
    ]),
    (("🛠️ 점검·관리", "🛠️ Checks · admin"), [
        ("/bat-doctor stage", ("인증·REST·로그 조회 점검", "Check auth, REST and log lookup"), []),
        ("/bat-sweep stage", ("이전 BAT 리소스 삭제", "Delete leftover BAT resources"), []),
        ("/bat-except add item", ("all 실행에서 뺄 제품·케이스 추가 (예: dlp, BAT-SWG-001)",
                                  "Add a product or case to skip in all runs (e.g. dlp, BAT-SWG-001)"), []),
        ("/bat-except delete item", ("제외 목록에서 빼기", "Remove from the exclusion list"), []),
        ("/bat-except list", ("제외 목록 보기", "Show the exclusion list"), []),
        ("/set lang", ("이 채널의 봇 언어 설정 (English·한국어)", "Set the bot language for this channel (English·한국어)"), []),
        ("/ping", ("봇 연결 확인", "Check the bot connection"), []),
    ]),
]


def guide_docs_url() -> str | None:
    """.env의 DISCORD_GUIDE_DOCS. 봇을 다시 켜지 않아도 바뀐 값이 반영되도록 매번 읽는다."""
    url = (dotenv_values(Path(__file__).resolve().parent / ".env").get("DISCORD_GUIDE_DOCS") or "").strip()
    return url if url.startswith(("http://", "https://")) else None


def help_page(index: int, lang: str = DEFAULT_LANG) -> discord.Embed:
    title, commands = HELP_SECTIONS[index]
    embed = discord.Embed(title=tr(lang, f"❓ 명령어 안내 · {title[0]}", f"❓ Commands · {title[1]}"), color=0x5865F2)
    # 사용법은 칸 제목(굵고 밝은 글씨), 설명과 옵션 값은 칸 내용으로 나눈다.
    for usage, text, options in commands:
        embed.add_field(name=usage, value="\n".join([tr(lang, *text)] + options), inline=False)
    embed.set_footer(text=tr(lang, f"{index + 1}/{len(HELP_SECTIONS)} · [ ]는 비워도 되는 옵션 · 모든 명령에 [lang] 옵션 (en·ko)",
                             f"{index + 1}/{len(HELP_SECTIONS)} · [ ] = optional · every command takes [lang] (en·ko)"))
    return embed


class HelpView(discord.ui.View):
    """카테고리 한 개씩 ◀ ▶로 넘긴다. 가이드 문서 링크는 버튼으로 항상 아래에 붙인다."""

    def __init__(self, interaction: discord.Interaction, lang: str = DEFAULT_LANG):
        super().__init__(timeout=300)
        self.interaction = interaction
        self.lang = lang
        self.index = 0
        self.prev.label = tr(lang, "◀ 이전", "◀ Prev")
        self.next.label = tr(lang, "다음 ▶", "Next ▶")
        url = guide_docs_url()
        if url:
            self.add_item(discord.ui.Button(label=tr(lang, "가이드 문서", "Guide"), emoji="📖", url=url))
        self.sync()

    def sync(self) -> None:
        self.prev.disabled = self.index == 0
        self.next.disabled = self.index == len(HELP_SECTIONS) - 1

    async def turn(self, interaction: discord.Interaction, step: int) -> None:
        self.index += step
        self.sync()
        await interaction.response.edit_message(embed=help_page(self.index, self.lang), view=self)

    @discord.ui.button(label="◀ Prev", style=discord.ButtonStyle.secondary)
    async def prev(self, interaction: discord.Interaction, button: discord.ui.Button):
        await self.turn(interaction, -1)

    @discord.ui.button(label="Next ▶", style=discord.ButtonStyle.secondary)
    async def next(self, interaction: discord.Interaction, button: discord.ui.Button):
        await self.turn(interaction, 1)

    async def on_timeout(self) -> None:
        self.prev.disabled = self.next.disabled = True
        try:
            await self.interaction.edit_original_response(view=self)
        except discord.HTTPException:
            pass


@tree.command(name="help", description=L("Commands and options (browse by category)",
                                         "명령어와 옵션 안내 (카테고리별로 넘겨 보기)"), guilds=TEST_GUILDS)
@with_lang
async def help_command(interaction: discord.Interaction, lang: app_commands.Choice[str] | None = None):
    lang = lang_for(interaction, lang)
    await interaction.response.send_message(embed=help_page(0, lang), view=HelpView(interaction, lang), ephemeral=True)


JENKINS_LINK_DEFAULT = "http://10.0.62.124:8080/"


def jenkins_link_url() -> str:
    """.env의 JENKINS_LINK_URL (브라우저로 여는 주소). 없으면 기본값. 매번 읽어 재시작 없이 반영."""
    url = (dotenv_values(Path(__file__).resolve().parent / ".env").get("JENKINS_LINK_URL") or "").strip()
    return url if url.startswith(("http://", "https://")) else JENKINS_LINK_DEFAULT


@tree.command(name="jenkins", description=L("Link to Jenkins", "Jenkins 바로가기 링크"), guilds=TEST_GUILDS)
@with_lang
async def jenkins_link(interaction: discord.Interaction, lang: app_commands.Choice[str] | None = None):
    lang = lang_for(interaction, lang)
    url = jenkins_link_url()
    view = discord.ui.View()
    view.add_item(discord.ui.Button(label=tr(lang, "Jenkins 열기", "Open Jenkins"), emoji="🔗", url=url))
    await interaction.response.send_message(f"🔗 Jenkins: {url}", view=view, ephemeral=True)


@tree.command(name="set", description=L("Set the bot language for this channel", "이 채널의 봇 언어 설정"),
              guilds=TEST_GUILDS)
@app_commands.describe(lang=L("Language for replies and notices in this channel", "이 채널의 응답·알림 언어"))
@app_commands.choices(lang=LANG_CHOICES)
async def set_lang(interaction: discord.Interaction, lang: app_commands.Choice[str]):
    value = norm_lang(lang) or DEFAULT_LANG
    if not interaction.channel_id:
        await interaction.response.send_message(tr(value, "채널에서만 쓸 수 있습니다.", "This only works in a channel."),
                                                ephemeral=True)
        return
    try:
        await asyncio.to_thread(set_channel_lang, interaction.channel_id, value)
    except Exception:
        log.exception("채널 언어 저장 실패")
        await interaction.response.send_message(tr(
            value, "언어 설정 저장 실패. 봇 콘솔 로그를 확인해 주세요.",
            "Failed to save the language setting. Please check the bot console log."), ephemeral=True)
        return
    await interaction.response.send_message(tr(
        value, "🌐 이 채널의 봇 언어를 **한국어**로 설정했습니다. (명령마다 `lang` 옵션으로 바꿀 수 있습니다)",
        "🌐 This channel's bot language is now **English**. (Override per command with the `lang` option)"))


@client.event
async def on_ready():
    log.info("로그인 성공: %s", client.user)


def build_ssl_context() -> ssl.SSLContext:
    """Discord 연결용 SSL 설정.

    인증서 검증과 호스트 이름 확인은 그대로 유지한다. Python 3.13부터 기본으로 켜진
    엄격 검사(VERIFY_X509_STRICT)만 끈다. 커넥터(SSE)의 SSL 검사 인증서에
    key usage 확장이 없어서 이 검사에 걸려 재연결이 실패하기 때문이다.
    """
    context = ssl.create_default_context()  # Windows 신뢰 저장소의 인증서 포함
    context.verify_flags &= ~ssl.VERIFY_X509_STRICT
    return context


async def main():
    # Discord(REST, Gateway) 연결에 위 SSL 설정 적용
    client.http.connector = aiohttp.TCPConnector(limit=0, ssl=build_ssl_context())
    try:
        async with client:
            await client.start(DISCORD_BOT_TOKEN)
    finally:
        if http is not None and not http.closed:
            await http.close()


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        log.info("봇 종료 (Ctrl+C)")
    except Exception:
        log.exception("봇이 오류로 종료됨")  # bot.log에 남긴다. start-bot.ps1이 잠시 뒤 다시 띄운다.
        sys.exit(1)