"""커넥터 자동화: 앱 전환 -> 스테이지 설정 -> 웹 로그인 -> 보안 연결 ON / 로그아웃.

앱: lguplus 스테이지는 U+SASE 앱, 그 외(qa, real 등)는 AIConnector 앱. 두 앱은 동시에 쓰지 않으며,
필요하면 실행 중인 앱을 '프로그램 종료'한 뒤 다른 앱을 실행한다.

사용법 (봇 폴더에서, .venv의 python으로):
  python connector.py on <stage>        앱 전환 + 스테이지 맞추기 + 로그인 + ON
  python connector.py on <stage> --stdin   사용자 지정 계정: 표준입력 1줄 JSON {"workspace","user","password"}
  python connector.py logout            로그아웃
  python connector.py status            현재 앱·화면·연결 상태
  python connector.py inspect <stage>   로그인 페이지의 입력칸·버튼 구조 출력 (입력값은 출력 안 함)
  python connector.py scheme            레지스트리의 커넥터 URL 프로토콜 목록

계정: <STAGE_ENV_DIR>/.env.<stage>의 BAT_CONSOLE_USER / BAT_CONSOLE_PASSWORD
워크스페이스: .env.<stage>의 CONNECTOR_WORKSPACE > environments.yaml의 스테이지 workspace(s) 값
앱 실행 파일: 봇 .env의 USASE_EXE / AICONNECTOR_EXE > 실행 중일 때 본 경로 > 기본 설치 경로

비밀번호는 화면 출력, 로그, 파일 어디에도 남기지 않는다.
마지막 줄에 'RESULT ...' 한 줄을 출력하고, 봇은 이 줄만 Discord로 보낸다.
"""
from __future__ import annotations

import getpass
import hashlib
import html
import json
import os
import re
import subprocess
import sys
import time
import traceback
from pathlib import Path
from urllib.parse import quote, urlparse

import keyring
import yaml

BASE_DIR = Path(__file__).resolve().parent
try:  # 단독 실행할 때도 .env 사용 (봇에서 실행하면 이미 로드되어 있음)
    from dotenv import load_dotenv

    load_dotenv(BASE_DIR / ".env")
except ImportError:
    pass

ENV_YAML = Path(os.environ.get("ENVIRONMENTS_YAML", r"C:\test-automation\environments.yaml"))
STAGE_ENV_DIR = Path(os.environ.get("CONNECTOR_ENV_DIR", str(ENV_YAML.parent)))  # .env.<stage> 위치
# 상태 파일 폴더 (저장소 밖, bot.py의 BOT_DATA_DIR과 같음)
DATA_DIR = Path(os.environ.get("BOT_DATA_DIR", "").strip() or r"C:\test-automation\bot-data")
DATA_DIR.mkdir(parents=True, exist_ok=True)
PROFILE_DIR = Path(os.environ.get("CONNECTOR_CHROME_PROFILE", str(DATA_DIR / "connector-chrome-profile")))
BROWSER_CHANNEL = os.environ.get("CONNECTOR_BROWSER_CHANNEL", "chrome") or None  # 비우면 Playwright Chromium
STATE_FILE = DATA_DIR / "connector-state.json"  # 마지막 로그인 정보 (계정은 해시로만 저장)
KEYRING_SERVICE = "usase-connector"
# 앱별 설정. scheme은 로그인 후 웹이 결과를 넘겨줄 프로그램 주소.
APPS = {
    "usase": {"label": "U+SASE", "scheme": "uplussase-connector", "exe_env": "USASE_EXE",
              "exe": r"C:\Program Files (x86)\U+SASE\U+SASE.exe",
              "conf_dir": r"C:\ProgramData\LGUplus\U+SASE",
              "match": r"u\+\s*sase|uplus"},
    "aiconnector": {"label": "AIConnector", "scheme": "aiconnector", "exe_env": "AICONNECTOR_EXE",
                    "exe": r"C:\Program Files (x86)\AIConnector for Client\AIConnector for Client.exe",
                    "conf_dir": r"C:\ProgramData\Monitorapp\AIConnector for Client",
                    "match": r"ai\s*connector"},
}
STAGE_APPS = {"lguplus": "usase"}
# 로그인 시작 경로. real은 /auth/login이 막혀 있어 /workspace로 시작한다. (.env.<stage>의 CONNECTOR_LOGIN_PATH로 변경 가능)
DEFAULT_LOGIN_PATH = "/auth/login"
STAGE_LOGIN_PATHS = {"real": "/workspace"}  # 나머지 스테이지는 DEFAULT_APP
DEFAULT_APP = "aiconnector"
SCHEME = APPS[DEFAULT_APP]["scheme"]  # 실행 중인 스테이지 기준으로 use_stage_scheme()이 바꾼다
OPEN_PATH = os.environ.get("CONNECTOR_OPEN_PATH", "open/")

# 숨김 메뉴 Server 목록에서 자동으로 시도할 항목 (Custom은 값이 바뀔 수 있어 제외)
ON_TEXT = "보안 연결 사용 중"
OFF_TEXT = "보안 연결 해제됨"

# 웹 로그인 화면 요소 (페이지 구조가 다르면 .env에서 바꿀 수 있음)
# ID 칸은 앞에서부터 차례로 시도한다 ('||'로 구분). 워크스페이스 칸은 ID 후보에서 제외한다.
ID_SELECTORS = os.environ.get(
    "CONNECTOR_ID_SELECTORS",
    "#id||input[type=email]||input[name*=user i]||input[name*=mail i]||input[name*=id i]"
    "||input[placeholder*=이메일]||input[placeholder*=아이디]||input[type=text]",
).split("||")
PW_SELECTOR = os.environ.get("CONNECTOR_PW_SELECTOR", "input[type=password]")
WS_SELECTOR = os.environ.get(
    "CONNECTOR_WS_SELECTOR",
    "input[name*=workspace i], input[id*=workspace i], input[placeholder*=workspace i], "
    "input[placeholder*=워크스페이스]",
)
SWITCH_WS_TEXT = os.environ.get("CONNECTOR_SWITCH_WS_TEXT", "다른 워크스페이스 로그인")
SUBMIT_SELECTOR = os.environ.get("CONNECTOR_SUBMIT_SELECTOR", "#login-submit-btn")


class ConnectorError(Exception):
    pass


NO_DESKTOP_MESSAGE = (
    "서버에 입력 화면이 없어(RDP 최소화·연결 끊김) 실제 마우스 조작이 필요한 단계를 진행하지 못했습니다. "
    "RDP를 연결한 상태에서 다시 시도해 주세요."
)


def _is_no_desktop(exc: Exception) -> bool:
    text = str(exc)
    return "no active desktop" in text.lower() or "SetCursorPos" in text


VERSION = "2026-11-12a"
STEP_WAIT = 3  # 단계 사이 대기(초). 앱이 이전 동작을 끝낼 시간을 준다.
RUN_LOG = DATA_DIR / "connector-last.log"  # 마지막 실행의 진행 기록 (토큰·비밀번호 없음, 주소 쿼리는 가림)
# 재시도 경로 시험용. 첫 시도에서 한 번만 실패를 일으킨다: double(딥링크 두 번 → 터널 충돌) / logout(로그인 직후 로그아웃)
TEST_FAIL = os.environ.get("CONNECTOR_TEST_FAIL", "").strip().lower()
_test_fail_used = False


def take_test_fail(kind: str) -> bool:
    global _test_fail_used
    if TEST_FAIL != kind or _test_fail_used:
        return False
    _test_fail_used = True
    return True


def say(message: str) -> None:
    """콘솔 출력 + 진행 기록 파일. 봇이 시간 초과로 끊어도 어디까지 갔는지 남는다."""
    print(message, flush=True)
    try:
        stamp = time.strftime("%H:%M:%S")
        with RUN_LOG.open("a", encoding="utf-8") as f:
            for line in str(message).splitlines() or [""]:
                f.write(f"{stamp} {line}\n")
    except Exception:
        pass


def wait_until(cond, timeout: float, interval: float = 1.0) -> bool:
    end = time.time() + timeout
    while time.time() < end:
        try:
            if cond():
                return True
        except Exception:
            pass
        time.sleep(interval)
    return False


# ------------------------------------------------------------
# environments.yaml
# ------------------------------------------------------------
def _dicts(obj):
    if isinstance(obj, dict):
        yield obj
        for value in obj.values():
            yield from _dicts(value)
    elif isinstance(obj, list):
        for value in obj:
            yield from _dicts(value)


def _find_key(node, name: str):
    for sub in _dicts(node):
        if sub.get(name) not in (None, ""):
            return sub[name]
    return None


def stage_node(stage: str) -> dict:
    """'qa: {console_url: ...}' 형태와 '- name: qa, console_url: ...' 형태를 모두 지원."""
    data = yaml.safe_load(ENV_YAML.read_text(encoding="utf-8")) or {}
    for d in _dicts(data):
        node = d.get(stage) if isinstance(d.get(stage), dict) else None
        if node is None and stage in (d.get("name"), d.get("stage")):
            node = d
        if node is not None and _find_key(node, "console_url"):
            return node
    raise ConnectorError(f"environments.yaml에서 '{stage}'의 console_url을 찾지 못했습니다.")


def get_console_url(stage: str) -> str:
    url = str(_find_key(stage_node(stage), "console_url")).strip().rstrip("/")
    return url if "://" in url else "https://" + url


def _workspace_value(value) -> str:
    """'qa2' / {'name': 'qa2'} / [{'name': 'qa2', 'lane': 'api'}, ...] 모두 지원 (목록이면 첫 번째)."""
    if isinstance(value, list):
        value = value[0] if value else ""
    if isinstance(value, dict):
        value = value.get("name") or value.get("workspace") or value.get("id") or ""
    return str(value).strip() if isinstance(value, (str, int)) else ""


def yaml_workspace(stage: str) -> str:
    node = stage_node(stage)
    for name in ("workspace", "workspace_name", "workspace_id", "workspaces"):
        value = _workspace_value(_find_key(node, name))
        if value:
            return value
    for sub in _dicts(node):
        for key, value in sub.items():
            if "workspace" in str(key).lower():
                value = _workspace_value(value)
                if value:
                    return value
    return ""


def host_of(url: str) -> str:
    return (urlparse(url.strip()).hostname or "").lower()


# ------------------------------------------------------------
# U+SASE 프로그램 창 (pywinauto)
# ------------------------------------------------------------
def app_for_stage(stage: str) -> str:
    app = stage_env(stage).get("CONNECTOR_APP", "").strip().lower() or STAGE_APPS.get(stage, DEFAULT_APP)
    if app not in APPS:
        raise ConnectorError(f"알 수 없는 앱 '{app}'. ({', '.join(APPS)} 중 하나)")
    return app


def use_stage_scheme(stage: str) -> str:
    global SCHEME
    value = stage_env(stage).get("CONNECTOR_SCHEME", "").strip() or APPS[app_for_stage(stage)]["scheme"]
    SCHEME = value.rstrip(":").lower()
    return SCHEME


def scheme_prefix() -> str:
    return f"{SCHEME}:"


def login_url(console_url: str, stage: str) -> str:
    path = stage_env(stage).get("CONNECTOR_LOGIN_PATH", "").strip() or STAGE_LOGIN_PATHS.get(stage, DEFAULT_LOGIN_PATH)
    path = "/" + path.strip("/")
    return f"{console_url}{path}?connector_url={quote(scheme_prefix() + OPEN_PATH, safe='')}"


# ------------------------------------------------------------
# 설정 파일(DPAPI 암호화)로 스테이지 전환 — 화면 조작 없이 server_mode만 바꾼다
# ------------------------------------------------------------
def _user_sid() -> str:
    out = subprocess.run(["whoami", "/user", "/fo", "csv", "/nh"], capture_output=True, text=True, errors="replace").stdout
    match = re.search(r"S-1-[\d-]+", out)
    return match.group(0) if match else ""


def config_dir(app: str) -> Path:
    base = Path(APPS[app]["conf_dir"])
    sid = _user_sid()
    path = base / sid / "config"
    if sid and (path / "appdata.bin").exists():
        return path
    found = sorted(base.glob("S-1-5-*/config/appdata.bin"), key=lambda p: p.stat().st_mtime, reverse=True)
    if found:
        return found[0].parent
    raise ConnectorError(f"{APPS[app]['label']} 설정 폴더를 찾지 못했습니다. ({base})")


def _dpapi_unprotect(data: bytes) -> bytes:
    import ctypes
    from ctypes import wintypes

    class BLOB(ctypes.Structure):
        _fields_ = [("cbData", wintypes.DWORD), ("pbData", ctypes.POINTER(ctypes.c_char))]

    buf = ctypes.create_string_buffer(data, len(data))
    blob_in = BLOB(len(data), ctypes.cast(buf, ctypes.POINTER(ctypes.c_char)))
    blob_out = BLOB()
    if not ctypes.windll.crypt32.CryptUnprotectData(
        ctypes.byref(blob_in), None, None, None, None, 0, ctypes.byref(blob_out)
    ):
        raise ConnectorError("설정 파일 복호화에 실패했습니다 (DPAPI). 같은 계정인지 확인해 주세요.")
    try:
        return ctypes.string_at(blob_out.pbData, blob_out.cbData)
    finally:
        ctypes.windll.kernel32.LocalFree(blob_out.pbData)


def _dpapi_protect(data: bytes) -> bytes:
    import ctypes
    from ctypes import wintypes

    class BLOB(ctypes.Structure):
        _fields_ = [("cbData", wintypes.DWORD), ("pbData", ctypes.POINTER(ctypes.c_char))]

    buf = ctypes.create_string_buffer(data, len(data))
    blob_in = BLOB(len(data), ctypes.cast(buf, ctypes.POINTER(ctypes.c_char)))
    blob_out = BLOB()
    # CRYPTPROTECT_LOCAL_MACHINE 없이 CurrentUser 범위로 다시 잠근다 (원본과 동일)
    if not ctypes.windll.crypt32.CryptProtectData(
        ctypes.byref(blob_in), None, None, None, None, 0, ctypes.byref(blob_out)
    ):
        raise ConnectorError("설정 파일 암호화에 실패했습니다 (DPAPI).")
    try:
        return ctypes.string_at(blob_out.pbData, blob_out.cbData)
    finally:
        ctypes.windll.kernel32.LocalFree(blob_out.pbData)


def _read_json_bin(path: Path) -> dict:
    return json.loads(_dpapi_unprotect(path.read_bytes()).decode("utf-8"))


def mode_for_host(app: str, target_host: str) -> int | None:
    """serverconf.bin의 server_list에서 target_host와 맞는 mode를 찾는다."""
    conf = _read_json_bin(config_dir(app) / "serverconf.bin")
    for entry in conf.get("server_list", []):
        hosts = {
            str(entry.get(key, "")).lower()
            for key in ("auth_host", "sdp_host", "api_host")
        }
        hosts.add(host_of(str(entry.get("consule_url", ""))))
        if target_host.lower() in {h for h in hosts if h}:
            return entry.get("mode")
    return None


def current_server_mode(app: str) -> int | None:
    try:
        return _read_json_bin(config_dir(app) / "appdata.bin").get("server_mode")
    except Exception:
        return None


def update_appdata(app: str, values: dict) -> bool:
    """appdata.bin의 지정한 키들만 바꿔 다시 암호화한다 (나머지는 그대로). 바뀐 게 있으면 True."""
    path = config_dir(app) / "appdata.bin"
    data = _read_json_bin(path)
    if all(data.get(k) == v for k, v in values.items()):
        return False
    data.update(values)
    encrypted = _dpapi_protect(json.dumps(data, ensure_ascii=False).encode("utf-8"))
    tmp = path.with_suffix(".bin.tmp")
    tmp.write_bytes(encrypted)
    tmp.replace(path)
    return True


def set_server_mode(app: str, mode: int) -> None:
    update_appdata(app, {"server_mode": mode})


def enable_auto_connect(app: str) -> None:
    """로그인 시 자동으로 보안 연결(ON)되도록 설정한다. 버튼 클릭 없이 ON 하기 위한 것.

    auto_connection_at_login_in_console은 True로 두면 안 된다(시험 중: 로그인 1초 뒤 앱이 스스로 로그아웃).
    """
    update_appdata(app, {
        "auto_connection_at_login": True,
        "auto_connection_at_login_in_console": False,
    })


def _process_path(pid: int) -> str:
    import ctypes
    from ctypes import wintypes

    kernel32 = ctypes.windll.kernel32
    handle = kernel32.OpenProcess(0x1000, False, pid)  # PROCESS_QUERY_LIMITED_INFORMATION
    if not handle:
        return ""
    try:
        buf = ctypes.create_unicode_buffer(1024)
        size = wintypes.DWORD(1024)
        if kernel32.QueryFullProcessImageNameW(handle, 0, buf, ctypes.byref(size)):
            return buf.value
        return ""
    finally:
        kernel32.CloseHandle(handle)


def _url_protocols():
    """레지스트리(HKCR)의 URL 프로토콜 목록: (scheme, 실행 명령)."""
    import winreg

    root = winreg.HKEY_CLASSES_ROOT
    i = 0
    while True:
        try:
            name = winreg.EnumKey(root, i)
        except OSError:
            break
        i += 1
        if name.startswith("."):
            continue
        try:
            with winreg.OpenKey(root, name) as key:
                winreg.QueryValueEx(key, "URL Protocol")
            with winreg.OpenKey(root, name + r"\shell\open\command") as key:
                command, _ = winreg.QueryValueEx(key, "")
        except OSError:
            continue
        yield name, os.path.expandvars(command or "")


def app_exe_path(win) -> str:
    try:
        return _process_path(win.element_info.process_id)
    except Exception:
        return ""


def find_window(timeout: float = 3):
    from pywinauto import Desktop

    win = Desktop(backend="uia").window(auto_id="xMainWindow", control_type="Window")
    return win if win.exists(timeout=timeout) else None


def main_window():
    win = find_window()
    if win is None:
        raise ConnectorError("커넥터 창을 찾지 못했습니다. 프로그램 창이 열려 있는지(트레이 아님) 확인해 주세요.")
    if win.is_minimized():
        win.restore()
    remember_exe(win)
    return win


def identify_app(win) -> str | None:
    text = f"{app_exe_path(win)} {win.window_text()}"
    for key in ("usase", "aiconnector"):  # U+SASE를 먼저 (설치 경로에 connector가 들어갈 수 있음)
        if re.search(APPS[key]["match"], text, re.IGNORECASE):
            return key
    return None


def remember_exe(win) -> None:
    """실행 중인 앱의 실행 파일 경로를 기억해 두고, 나중에 앱을 다시 켤 때 쓴다."""
    app, exe = identify_app(win), app_exe_path(win)
    if app and exe:
        state = load_state()
        exes = state.setdefault("exe", {})
        if exes.get(app) != exe:
            exes[app] = exe
            save_state(state)


def exe_for(app: str) -> str:
    cfg = APPS[app]
    path = (
        os.environ.get(cfg["exe_env"], "").strip().strip('"')
        or load_state().get("exe", {}).get(app, "")
        or cfg["exe"]
    )
    if not path or not Path(path).exists():
        raise ConnectorError(
            f"{cfg['label']} 실행 파일 경로를 모릅니다. 봇 .env에 {cfg['exe_env']}=실행파일경로 를 넣어 주세요."
        )
    return path


def _tasklist() -> list:
    out = subprocess.run(["tasklist", "/FO", "CSV", "/NH"], capture_output=True, text=True, errors="replace").stdout
    rows = []
    for line in out.splitlines():
        parts = [part.strip('"') for part in line.split('","')]
        if len(parts) >= 2 and parts[1].isdigit():
            rows.append((parts[0].lower(), int(parts[1])))
    return rows


def _image_of(app: str) -> str:
    try:
        return Path(exe_for(app)).name.lower()
    except ConnectorError:
        return ""


def app_running(app: str) -> bool:
    image = _image_of(app)
    return bool(image) and any(name == image for name, _ in _tasklist())


def _click_confirm(label_re: str) -> bool:
    """종료·로그아웃 확인 창이 뜨면 확인/예를 누른다."""
    from pywinauto import Desktop

    for w in Desktop(backend="uia").windows():
        try:
            if not re.search(label_re, w.window_text() or "", re.IGNORECASE):
                continue
            for button in w.descendants(control_type="Button"):
                if (button.window_text() or "").strip() in ("확인", "예", "종료", "OK", "Yes"):
                    button.invoke()
                    return True
        except Exception:
            continue
    return False


def _tray_buttons(desktop):
    taskbar = desktop.window(class_name="Shell_TrayWnd")
    if taskbar.exists(timeout=1):
        area = taskbar.child_window(class_name="TrayNotifyWnd")
        yield from (area if area.exists(timeout=0) else taskbar).descendants(control_type="Button")
    for cls in ("NotifyIconOverflowWindow", "TopLevelWindowForOverflowXamlIsland"):
        overflow = desktop.window(class_name=cls)
        if overflow.exists(timeout=0):
            yield from overflow.descendants(control_type="Button")


def _find_tray_icon(desktop, label_re: str):
    for button in _tray_buttons(desktop):
        if re.search(label_re, button.window_text() or "", re.IGNORECASE):
            return button
    taskbar = desktop.window(class_name="Shell_TrayWnd")  # 숨겨진 아이콘 펼치기
    if taskbar.exists(timeout=0):
        for button in taskbar.descendants(control_type="Button"):
            if re.search(r"숨겨진 아이콘|hidden icons|chevron", button.window_text() or "", re.IGNORECASE):
                button.click_input()
                time.sleep(1)
                break
    for button in _tray_buttons(desktop):
        if re.search(label_re, button.window_text() or "", re.IGNORECASE):
            return button
    return None


def _find_menu_item(desktop, title: str, timeout: float = 5):
    end = time.time() + timeout
    while time.time() < end:
        for w in desktop.windows():
            try:
                for item in w.descendants(control_type="MenuItem"):
                    if (item.window_text() or "").strip() == title:
                        return item
            except Exception:
                continue
        time.sleep(0.3)
    return None


def exit_via_tray(app: str) -> None:
    from pywinauto import Desktop, keyboard

    desktop = Desktop(backend="uia")
    label = APPS[app]["label"]
    icon = _find_tray_icon(desktop, APPS[app]["match"])
    if icon is None:
        raise ConnectorError(f"트레이에서 {label} 아이콘을 찾지 못했습니다.")
    try:
        icon.click_input(button="right")
    except Exception as e:
        if _is_no_desktop(e):
            raise ConnectorError(NO_DESKTOP_MESSAGE) from None
        raise
    item = _find_menu_item(desktop, "프로그램 종료")
    if item is None:
        keyboard.send_keys("{ESC}")
        raise ConnectorError(f"{label} 트레이 메뉴에서 '프로그램 종료'를 찾지 못했습니다.")
    try:
        item.invoke()
    except Exception:
        item.click_input()


def _kill_image(image: str) -> bool:
    """로그인 전(보안 연결 꺼짐) 화면 프로그램만 끈다. 백그라운드 서비스는 건드리지 않는다."""
    result = subprocess.run(["taskkill", "/IM", image, "/F"], capture_output=True, text=True, errors="replace")
    return result.returncode == 0


def show_app_window(app: str):
    """트레이에만 있는 앱의 창을 띄운다 (실행 파일을 다시 실행하면 기존 창이 열림)."""
    exe = exe_for(app)
    subprocess.Popen([exe], cwd=str(Path(exe).parent))

    def ready() -> bool:
        w = find_window(0)
        return w is not None and identify_app(w) == app

    if not wait_until(ready, timeout=30):
        return None
    return main_window()


def exit_app(app: str, win=None) -> None:
    """'프로그램 종료'로 앱을 끈다. 로그인 상태면 설정 메뉴, 아니면 트레이 메뉴를 쓴다."""
    label = APPS[app]["label"]
    image = _image_of(app)

    def gone() -> bool:
        if image:
            return not any(name == image for name, _ in _tasklist())
        return find_window(0) is None

    if win is None or _win_gone(win):
        win = show_app_window(app) if app_running(app) else None
    done = False
    if win is not None and is_home(win):
        try:
            _click_setting_menu(win, "프로그램 종료", done=gone)
            done = True
        except Exception:
            # 메뉴를 누르는 도중 앱이 꺼지면 '누르기 실패'로 끝나기도 한다. 실제로 꺼졌는지 확인한다.
            done = wait_until(gone, timeout=5, interval=0.5)
    if not done and image:
        done = _kill_image(image) or gone()  # 화면 프로그램만 종료 (서비스는 그대로)
    if not done:
        if not desktop_available():
            raise ConnectorError(f"{label}을(를) 종료하지 못했습니다. (화면 없음, 프로세스 종료 실패)")
        exit_via_tray(app)  # 마지막 수단 (실제 마우스 입력 필요)

    if not wait_until(gone, timeout=5):
        _click_confirm(APPS[app]["match"])
        if not wait_until(gone, timeout=30):
            raise ConnectorError(f"{label} 프로그램 종료를 확인하지 못했습니다.")


def start_app(app: str):
    label = APPS[app]["label"]
    exe = exe_for(app)
    say(f"{label} 실행")
    started = time.time()
    subprocess.Popen([exe], cwd=str(Path(exe).parent))

    def ready() -> bool:
        w = find_window(0)
        return w is not None and identify_app(w) == app

    if not wait_until(ready, timeout=60):
        raise ConnectorError(f"{label}을(를) 실행했지만 60초 안에 창이 뜨지 않았습니다.")
    time.sleep(5)  # 앱 초기화(서비스 연결 등) 대기
    wait_startup_login_settled(app, started)
    win = visible_window(app)  # 그 사이 남은 토큰으로 자동 로그인·연결되면 창이 트레이로 숨는다
    if win.is_minimized():
        win.restore()
    remember_exe(win)
    if dismiss_notice(win):
        say("  시작 시 자동 로그인 실패 공지 닫음 (웹 로그인으로 진행)")
    return win


def wait_startup_login_settled(app: str, started: float, timeout: float = 40) -> None:
    """앱 시작 직후 남은 토큰으로 하는 자동 로그인 재시도(최대 5회, 약 20초)가 끝날 때까지 기다린다.

    재시도 도중 웹 로그인 주소가 들어오면 두 로그인이 겹칠 수 있다. 시작 후 8초 안에 시도가 없으면 바로 진행.
    """
    def settled() -> bool:
        lines = app_log_since(app, started)
        if any("Login success" in line for line in lines):
            return True
        tried = [line for line in lines if "Auto login failed (" in line]
        if not tried:
            return time.time() - started > 8
        return "(5/5)" in tried[-1]

    if not wait_until(settled, timeout=timeout, interval=1.0):
        say("  (앱의 시작 시 자동 로그인 재시도가 끝나지 않았지만 진행)")


def ensure_app(app: str):
    """필요한 앱 창을 준비한다. (창, 전환 설명)"""
    label = APPS[app]["label"]
    other = "aiconnector" if app == "usase" else "usase"
    win = find_window()
    if win is not None:
        current = identify_app(win)
        remember_exe(win)
        if current == app:
            if win.is_minimized():
                win.restore()
            return win, ""
        if current is None:
            raise ConnectorError(f"실행 중인 커넥터 앱을 구분하지 못했습니다. (창 제목: {win.window_text()})")
        say(f"{APPS[current]['label']} 종료 후 {label}(으)로 전환")
        exit_app(current, win)
        return start_app(app), f"{APPS[current]['label']}→{label} 전환"
    if app_running(other):  # 창 없이 트레이에만 있는 경우: 창을 띄워서 상태를 보고 종료
        say(f"트레이의 {APPS[other]['label']} 종료 후 {label}(으)로 전환")
        exit_app(other, show_app_window(other))
        return start_app(app), f"{APPS[other]['label']}→{label} 전환"
    if app_running(app):  # 필요한 앱이 트레이에만 있는 경우
        win = show_app_window(app)
        if win is not None:
            return win, ""
    return start_app(app), f"{label} 실행"


def page_name(win) -> str:
    return win.child_window(auto_id="xFrame").window_text()


def is_home(win) -> bool:
    return "HomePage" in page_name(win)


def is_login(win) -> bool:
    return "LoginPage" in page_name(win)


def status_text(win) -> str:
    frame = win.child_window(auto_id="xFrame")
    for text in (ON_TEXT, OFF_TEXT):
        if frame.child_window(title=text, control_type="Text").exists(timeout=0):
            return text
    return ""


def ensure_stage(win, app: str, target_host: str):
    """appdata.bin의 server_mode를 바꿔 스테이지를 맞춘다. 화면 조작 불필요. (창, 바뀜 여부)

    로그인 전(보안 연결 꺼짐) 화면에서 앱을 끄고 server_mode만 바꾼 뒤 다시 켠다.
    """
    mode = mode_for_host(app, target_host)
    if mode is None:
        raise ConnectorError(f"{APPS[app]['label']} 서버 목록에서 {target_host}와 맞는 항목을 찾지 못했습니다.")
    if current_server_mode(app) == mode:
        return win, False
    say(f"설정 파일에서 server_mode를 {mode}로 바꿔 스테이지 변경")
    exit_app(app, win)
    # 이전 스테이지의 토큰이 남아 있으면 앱이 켜지자마자 그 토큰으로 새 서버에 자동 로그인하다 401이 나고
    # "로그인 정보 만료" 공지를 띄운다(server_mode만 바꿔 비교 확인). 토큰을 비우면 대부분 시도 자체를 안 한다.
    update_appdata(app, {"server_mode": mode, "r_token": "", "auto_login": False})
    enable_auto_connect(app)  # 앱이 꺼져 있을 때 써야 앱이 메모리 값으로 덮어쓰지 않는다
    return start_app(app), True


def desktop_available() -> bool:
    """화면이 잠기거나 연결이 끊겨 입력 화면이 없으면 False."""
    import ctypes

    user32 = ctypes.windll.user32
    handle = user32.OpenInputDesktop(0, False, 0x0100)  # DESKTOP_SWITCHDESKTOP
    if not handle:
        return False
    try:
        return bool(user32.SwitchDesktop(handle))
    finally:
        user32.CloseDesktop(handle)


def _post_click(win, element, modifiers: int = 0) -> None:
    """실제 마우스 없이 창에 클릭 메시지를 보낸다 (화면이 없어도 동작하는 경우가 많음)."""
    import win32api
    import win32con
    import win32gui

    hwnd = win.handle
    rect = element.rectangle()
    x, y = win32gui.ScreenToClient(hwnd, ((rect.left + rect.right) // 2, (rect.top + rect.bottom) // 2))
    lparam = win32api.MAKELONG(x & 0xFFFF, y & 0xFFFF)
    win32gui.PostMessage(hwnd, win32con.WM_MOUSEMOVE, modifiers, lparam)
    win32gui.PostMessage(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON | modifiers, lparam)
    time.sleep(0.05)
    win32gui.PostMessage(hwnd, win32con.WM_LBUTTONUP, modifiers, lparam)


def _setting_menu(win):
    """열린 설정 메뉴 (메인 창 안 또는 별도 팝업 창). 없으면 None."""
    from pywinauto import Desktop

    menu = win.child_window(auto_id="settingMenu", control_type="Menu")
    if menu.exists(timeout=0):
        return menu
    try:
        pid = win.process_id()
    except Exception:
        return None
    for w in Desktop(backend="uia").windows():  # 같은 프로그램의 별도 팝업 창
        try:
            if w.process_id() != pid:
                continue
            for element in w.descendants(control_type="Menu"):
                if element.element_info.automation_id == "settingMenu":
                    return element
        except Exception:
            continue
    return None


def _click_setting_menu(win, title: str, done=None) -> None:
    """설정(톱니) 버튼 -> 메뉴 항목. 로그인 상태에서만 있다.

    버튼 누르기는 화면 없이 되는 방법부터 시도한다: UIA 호출 -> 클릭 메시지 -> 실제 클릭.
    """
    try:
        win.set_focus()
    except Exception:
        pass
    button = win.child_window(auto_id="xSettingButton")
    attempts = [("invoke", lambda: button.invoke()), ("message", lambda: _post_click(win, button))]
    if desktop_available():
        attempts.append(("click", lambda: button.click_input()))
    menu, tried = None, []
    for name, press in attempts:
        try:
            press()
        except Exception as e:
            tried.append(f"{name}:{type(e).__name__}")
            continue
        tried.append(name)
        if wait_until(lambda: _setting_menu(win) is not None, timeout=3, interval=0.3):
            menu = _setting_menu(win)
            break
    if menu is None:
        hint = "" if desktop_available() else " 서버 화면이 잠긴 상태입니다(잠금 화면·화면 보호기 설정 확인)."
        raise ConnectorError(f"설정 메뉴가 열리지 않아서 '{title}'을(를) 누르지 못했습니다. (시도: {', '.join(tried)}){hint}")
    item = None
    for candidate in menu.descendants(control_type="MenuItem"):
        if (candidate.window_text() or "").strip() == title:
            item = candidate
            break
    if item is None:
        raise ConnectorError(f"설정 메뉴에서 '{title}' 항목을 찾지 못했습니다.")
    _press_menu_item(win, item, done)


def _press_menu_item(win, item, done=None) -> None:
    """메뉴 항목을 누른다. 메뉴가 닫혔거나 done()(실행 결과)이 참이면 성공.

    세션이 잠긴 상태(RDP 끊김)에서는 항목이 실행돼도 메뉴가 닫힌 것으로 보이지 않을 수 있어
    결과로도 판단한다. 결과 확인 없이 다음 방법으로 또 누르면 같은 동작이 두 번 실행될 수 있다.
    """
    def closed() -> bool:
        try:
            if done is not None and done():
                return True
        except Exception:
            pass
        return _setting_menu(win) is None

    attempts = [("invoke", lambda: item.invoke())]
    try:
        popup = item.top_level_parent()
        attempts.append(("message", lambda: _post_click(popup, item)))
    except Exception:
        pass
    attempts.append(("message-main", lambda: _post_click(win, item)))
    if desktop_available():
        attempts.append(("click", lambda: item.click_input()))
    tried = []
    for name, press in attempts:
        try:
            press()
        except Exception as e:
            tried.append(f"{name}:{type(e).__name__}")
            continue
        tried.append(name)
        # 로그아웃은 VPN 종료를 기다리느라 실행까지 몇 초 걸린다
        if wait_until(closed, timeout=3 if done is None else 8, interval=0.3):
            return
    raise ConnectorError(
        f"설정 메뉴 항목을 누르지 못했습니다. (시도: {', '.join(tried)}, "
        f"입력 화면 {'있음' if desktop_available() else '없음'})"
    )


def logout(win) -> None:
    _click_setting_menu(win, "로그아웃", done=lambda: is_login(win))
    if wait_until(lambda: is_login(win), timeout=10):
        return
    _click_confirm(APPS["usase"]["match"] + "|" + APPS["aiconnector"]["match"])  # 확인 창이 뜨는 경우
    if not wait_until(lambda: is_login(win), timeout=15):
        raise ConnectorError("로그아웃 후 로그인 화면으로 돌아오지 않았습니다.")


def _press(win, element) -> None:
    """요소를 누른다. invoke -> 클릭 메시지 -> (화면 있으면) 실제 클릭 순."""
    try:
        element.invoke()
        return
    except Exception:
        pass
    try:
        _post_click(win, element)
        return
    except Exception:
        pass
    if desktop_available():
        element.click_input()
    else:
        raise ConnectorError(NO_DESKTOP_MESSAGE)


def dismiss_notice(win) -> bool:
    """중앙 UI를 덮는 공지(로그인 만료 등) 팝업이 있으면 닫는다. 닫았으면 True."""
    frame = win.child_window(auto_id="xFrame")
    closed = False
    for _ in range(3):
        target = None
        # 같은 프로그램의 별도 팝업 창이나 xFrame 안의 확인/닫기 버튼을 찾는다
        for scope in (win, frame):
            for button in scope.descendants(control_type="Button"):
                label = (button.window_text() or "").strip()
                if label in ("확인", "닫기", "다시 로그인", "OK", "Close"):
                    target = button
                    break
            if target is not None:
                break
        if target is None:
            return closed
        try:
            _press(win, target)
            closed = True
            time.sleep(1)
        except Exception:
            return closed
    return closed


def _press_button(win, element, verify, timeout: float = 30) -> bool:
    """여러 방법으로 눌러 보고, verify()가 참이 될 때까지 기다린다. 성공하면 True.

    이 앱의 큰 버튼은 invoke/클릭 메시지를 '받았다고는 하지만' 실제로는 처리하지 않는 경우가 있어,
    누른 뒤 결과를 확인하고 안 되면 다음 방법으로 넘어간다. 좌표 클릭이 가장 확실하다.
    """
    methods = []
    if desktop_available():
        methods.append(("click_input", lambda: element.click_input()))
    methods.append(("invoke", lambda: element.invoke()))
    methods.append(("message", lambda: _post_click(win, element)))
    for name, press in methods:
        try:
            press()
        except Exception:
            continue
        if wait_until(verify, timeout=timeout, interval=1.0):
            return True
    return False


def turn_on(win) -> bool:
    """보안 연결 ON. 이미 켜져 있으면 False."""
    wait_until(lambda: status_text(win) != "", timeout=20)
    dismiss_notice(win)  # ON 버튼을 가리는 공지 팝업이 있으면 먼저 닫는다
    if status_text(win) == ON_TEXT:
        return False
    button = win.child_window(auto_id="xFrame").child_window(control_type="Button", found_index=0)
    # real 등은 터널 수립이 느려서 넉넉히(90초) 기다린다
    if _press_button(win, button, lambda: status_text(win) == ON_TEXT, timeout=90):
        return True
    dismiss_notice(win)  # 늦게 뜬 공지 팝업 대비
    if _press_button(win, button, lambda: status_text(win) == ON_TEXT, timeout=60):
        return True
    raise ConnectorError(
        f"ON 버튼을 눌렀지만 '{ON_TEXT}'로 바뀌지 않았습니다. (현재: {status_text(win) or '확인 불가'}, "
        f"입력 화면 {'있음' if desktop_available() else '없음'})"
    )


# ------------------------------------------------------------
# 웹 로그인 (Playwright, 전용 Chrome 프로필)
# ------------------------------------------------------------
def stage_env(stage: str) -> dict:
    """.env.<stage> 값을 읽기만 한다 (os.environ에 넣지 않음)."""
    path = STAGE_ENV_DIR / f".env.{stage}"
    if not path.exists():
        return {}
    from dotenv import dotenv_values

    return {k: (v or "") for k, v in dotenv_values(path).items()}


def workspace_for(stage: str) -> str:
    value = stage_env(stage).get("CONNECTOR_WORKSPACE", "").strip()
    return value or yaml_workspace(stage)


def credentials(stage: str):
    values = stage_env(stage)
    user_id = values.get("BAT_CONSOLE_USER", "").strip()
    password = values.get("BAT_CONSOLE_PASSWORD", "")
    if user_id and password:
        return user_id, password
    for key in (stage, "default"):
        user_id = keyring.get_password(KEYRING_SERVICE, f"{key}:id")
        password = keyring.get_password(KEYRING_SERVICE, f"{key}:password")
        if user_id and password:
            return user_id, password
    raise ConnectorError(
        f"'{stage}' 계정을 찾지 못했습니다. {STAGE_ENV_DIR}\\.env.{stage}에 BAT_CONSOLE_USER와 "
        "BAT_CONSOLE_PASSWORD가 있는지 확인해 주세요."
    )


def disable_password_saving() -> None:
    """전용 프로필의 Chrome이 비밀번호를 저장하지 않도록 설정.

    '항상 허용'된 프로그램 주소(aiconnector: 등) 기록도 지운다. 남아 있으면 Chrome이 직접 앱을 실행하고
    hand_to_app도 한 번 더 넘겨서 로그인이 두 번 처리되고, 자동 연결이 두 번 동시에 돌아 터널 생성이 충돌한다.
    """
    pref = PROFILE_DIR / "Default" / "Preferences"
    pref.parent.mkdir(parents=True, exist_ok=True)
    try:
        data = json.loads(pref.read_text(encoding="utf-8")) if pref.exists() else {}
    except Exception:
        data = {}
    data["credentials_enable_service"] = False
    data.setdefault("profile", {})["password_manager_enabled"] = False
    data.get("protocol_handler", {}).pop("allowed_origin_protocol_pairs", None)
    pref.write_text(json.dumps(data), encoding="utf-8")


def launch_chrome(p):
    disable_password_saving()
    options = dict(
        channel=BROWSER_CHANNEL,  # 기본: 설치된 Chrome
        headless=False,
        args=["--no-first-run", "--no-default-browser-check"],
    )
    try:  # 샌드박스를 켜서 '--no-sandbox' 경고 표시 제거
        return p.chromium.launch_persistent_context(str(PROFILE_DIR), chromium_sandbox=True, **options)
    except Exception:
        return p.chromium.launch_persistent_context(str(PROFILE_DIR), **options)


def vis(page, selector: str):
    """화면에 보이는 요소만 고른다 (숨겨진 같은 종류의 칸이 앞에 있는 페이지 대응)."""
    return page.locator(f"{selector} >> visible=true")


def _visible(locator) -> bool:
    try:
        return locator.first.is_visible()
    except Exception:
        return False


def _is_workspace_box(locator) -> bool:
    try:
        text = locator.evaluate("e => (e.name + ' ' + e.id + ' ' + e.placeholder).toLowerCase()")
    except Exception:
        return False
    return "workspace" in text or "워크스페이스" in text


def _find_id_box(page):
    for selector in ID_SELECTORS:
        for loc in page.locator(selector.strip()).all():
            if loc.is_visible() and not _is_workspace_box(loc):
                return loc
    return None


def _page_diag(ctx) -> None:
    """입력칸을 못 찾았을 때 탭별로 무엇이 보이는지 출력 (값은 출력 안 함)."""
    for i, pg in enumerate(ctx.pages):
        try:
            pw, ws = pg.locator(PW_SELECTOR), pg.locator(WS_SELECTOR)
            say(f"[진단] 탭{i} {_mask(pg.url)} 비밀번호칸 {pw.count()}개(보이는 것 {vis(pg, PW_SELECTOR).count()}개) "
                f"워크스페이스칸 {ws.count()}개(보이는 것 {vis(pg, WS_SELECTOR).count()}개) 프레임 {len(pg.frames)}개")
        except Exception as e:
            say(f"[진단] 탭{i} 확인 실패: {type(e).__name__}")


def _wait_state(ctx, holder: dict, timeout: float = 20):
    """(상태, 페이지). 새 탭으로 넘어간 경우도 있어서 모든 탭을 본다."""
    end = time.time() + timeout
    while time.time() < end:
        if holder.get("url"):  # 입력 화면 없이 바로 프로그램 주소로 넘어간 경우
            return "handoff", None
        for pg in reversed(ctx.pages):
            if pg.is_closed():
                continue
            if _visible(vis(pg, PW_SELECTOR)):
                return "password", pg
            if _visible(vis(pg, WS_SELECTOR)):
                return "workspace", pg
        time.sleep(0.3)
    _page_diag(ctx)
    raise ConnectorError(
        "로그인 페이지에서 워크스페이스 칸도 비밀번호 칸도 찾지 못했습니다. 위 [진단] 줄을 보내 주세요."
    )


def _fill_workspace(page, workspace: str, stage: str) -> None:
    if not workspace:
        raise ConnectorError(
            f"워크스페이스 입력이 필요한데 값이 없습니다. environments.yaml의 '{stage}'에 workspace를 넣거나 "
            "사용자 지정 로그인으로 입력해 주세요."
        )
    box = vis(page, WS_SELECTOR).first
    box.fill(workspace)
    button = page.get_by_role("button", name="로그인", exact=True)
    if _visible(button):
        button.first.click()
    else:
        box.press("Enter")


def login_flow(ctx, holder: dict, stage: str, user_id: str, password: str, workspace: str) -> None:
    """워크스페이스 입력 화면 / 계정 입력 화면(이전 워크스페이스가 기억된 상태)을 모두 처리.

    계정 입력 화면이 바로 뜨면 어떤 워크스페이스가 기억돼 있는지 화면에 안 나오므로,
    워크스페이스 값이 있으면 항상 '다른 워크스페이스 로그인'으로 다시 지정한다.
    """
    ws_count = 0
    switched = False
    for _ in range(8):
        state, page = _wait_state(ctx, holder)
        if state == "handoff":
            return
        if state == "workspace":
            if ws_count >= 2:
                raise ConnectorError(f"워크스페이스 '{workspace}'를 입력해도 계속 워크스페이스 화면입니다. (이름 확인)")
            _fill_workspace(page, workspace, stage)
            ws_count += 1
            # 화면 전환(페이지 이동)이 끝날 때까지 잠깐 기다린 뒤 다시 상태를 본다
            wait_until(lambda: not _visible(vis(page, WS_SELECTOR)), timeout=15, interval=0.3)
            continue

        switch = page.get_by_text(SWITCH_WS_TEXT, exact=True)
        if workspace and ws_count == 0 and not switched and _visible(switch):
            switch.first.click()
            switched = True
            wait_until(lambda: not _visible(vis(page, PW_SELECTOR)), timeout=10, interval=0.3)
            continue

        id_box = _find_id_box(page)
        if id_box is not None:
            id_box.fill(user_id)
        pw_box = vis(page, PW_SELECTOR).first
        pw_box.fill(password)
        submit = vis(page, SUBMIT_SELECTOR)
        if _visible(submit):
            submit.first.click()
        else:
            pw_box.press("Enter")
        return
    raise ConnectorError("로그인 화면 단계가 8번 넘게 반복되어 중단했습니다.")


# 자동화로 띄운 Chrome은 외부 프로그램 주소(scheme:...)를 열지 않는다.
# 그래서 로그인 뒤 넘어오는 프로그램 주소를 직접 잡아서 프로그램에 넘긴다.
CAPTURE_JS = """(() => {
  const keep = u => { try { u = String(u); if (u.toLowerCase().startsWith('__SCHEME__')) window.__connectorUrl = u; } catch (e) {} };
  const open = window.open;
  window.open = function (u, ...rest) { keep(u); return open.call(this, u, ...rest); };
  document.addEventListener('click', e => {
    const a = e.target && e.target.closest && e.target.closest('a[href]');
    if (a) keep(a.getAttribute('href'));
  }, true);
})();"""


def _mask(url: str) -> str:
    return url.split("?", 1)[0] + ("?<가림>" if "?" in url else "")


_QUERY_RE = re.compile(r"([a-z][a-z0-9+.\-]*:[^?\s'\"]*)\?[^\s'\"]*", re.IGNORECASE)


def sanitize(text: str) -> str:
    """오류 메시지·Traceback에 섞인 주소의 쿼리(토큰 포함)를 가린다."""
    return _QUERY_RE.sub(r"\1?<가림>", text)


def _protocol_command() -> str:
    """Windows에 등록된 SCHEME: 실행 명령 (없으면 빈 문자열)."""
    import winreg

    keys = (
        (winreg.HKEY_CURRENT_USER, rf"Software\Classes\{SCHEME}\shell\open\command"),
        (winreg.HKEY_LOCAL_MACHINE, rf"Software\Classes\{SCHEME}\shell\open\command"),
        (winreg.HKEY_CLASSES_ROOT, rf"{SCHEME}\shell\open\command"),
    )
    for root, path in keys:
        try:
            with winreg.OpenKey(root, path) as key:
                value, _ = winreg.QueryValueEx(key, "")
            if value:
                return os.path.expandvars(value)
        except OSError:
            continue
    return ""


def _fill_command(command: str, url: str) -> str:
    if '"%1"' in command:
        return command.replace('"%1"', f'"{url}"')
    if "%1" in command:
        return command.replace("%1", f'"{url}"')
    return f'{command} "{url}"'


def hand_to_app(win, url: str) -> str:
    """프로그램 주소를 프로그램에 넘긴다.

    주소 안의 발급 토큰은 한 번만 쓸 수 있을 수 있어서, 전달이 '실행'된 뒤에는 다른 방법으로
    다시 보내지 않고 충분히(60초) 기다린다. 실행 자체가 실패했을 때만 다음 방법을 쓴다.
    """
    command = _protocol_command()
    attempts = []
    if command:
        attempts.append(("등록된 실행 명령", lambda: subprocess.Popen(_fill_command(command, url))))
    attempts.append(("ShellExecute", lambda: os.startfile(url)))

    failures = []
    for name, run in attempts:
        try:
            proc = run()
        except OSError as e:
            failures.append(f"{name}: {type(e).__name__}({getattr(e, 'winerror', '')})")
            continue
        say(f"  -> {name} 방식으로 전달, 프로그램 로그인 확인 중...")
        if take_test_fail("double"):
            time.sleep(0.5)
            say("  [시험] 같은 주소를 한 번 더 전달 (딥링크 중복 재현)")
            run()
        if wait_until(lambda: is_home(win), timeout=60):
            return name
        launcher = ""
        if proc is not None and proc.poll() is not None:
            launcher = f", 런처 종료 코드 {proc.returncode}"
        try:
            screen = f"{page_name(win)} / {status_text(win) or '-'}"
        except Exception:
            screen = "확인 불가"
        raise ConnectorError(
            f"{SCHEME}: 주소를 {name} 방식으로 넘겼지만 60초 안에 로그인되지 않았습니다. "
            f"(현재 화면: {screen}{launcher}, 입력 화면 {'있음' if desktop_available() else '없음'})"
        )
    raise ConnectorError(f"{SCHEME}: 주소를 프로그램에 넘기지 못했습니다. (" + "; ".join(failures) + ")")


def _attach_capture(ctx, holder: dict) -> None:
    def on_request(req):
        if req.url.lower().startswith(scheme_prefix()):
            holder.setdefault("url", req.url)
        try:
            if req.is_navigation_request():
                say(f"[이동] {_mask(req.url)}")
        except Exception:
            pass

    def on_response(resp):
        try:
            location = resp.headers.get("location", "")
        except Exception:
            location = ""
        if location.lower().startswith(scheme_prefix()):
            holder.setdefault("url", location)
        holder.setdefault("trail", []).append(f"{resp.status} {_mask(resp.url)}")
        del holder["trail"][:-8]

    ctx.on("request", on_request)
    ctx.on("response", on_response)
    ctx.add_init_script(CAPTURE_JS.replace("__SCHEME__", scheme_prefix()))


def _attach_cdp_capture(ctx, page, holder: dict) -> None:
    """외부 프로그램 주소로 가려는 이동은 네트워크 요청이 안 생겨서, Chrome 내부 이벤트로 잡는다."""
    def keep(event):
        url = (event or {}).get("url", "")
        if url.lower().startswith(scheme_prefix()):
            holder.setdefault("url", url)

    try:
        cdp = ctx.new_cdp_session(page)
        cdp.send("Page.enable")
        cdp.on("Page.frameRequestedNavigation", keep)
        cdp.on("Page.frameScheduledNavigation", keep)
    except Exception as e:
        say(f"(CDP 감시 비활성: {type(e).__name__})")


def _scan_pages(ctx) -> str:
    for page in ctx.pages:
        try:
            if page.url.lower().startswith(scheme_prefix()):
                return page.url
            value = page.evaluate("window.__connectorUrl || ''")
            if value:
                return value
            match = re.search(re.escape(scheme_prefix()) + r"[^\s\"'<>]+", page.content(), re.IGNORECASE)
            if match:
                return html.unescape(match.group(0))
        except Exception:
            continue
    return ""


def wait_handoff(win, ctx, holder: dict, timeout: float = 60) -> str:
    """프로그램이 로그인되거나, 프로그램 주소를 잡아서 직접 넘길 때까지 기다린다."""
    end = time.time() + timeout
    while time.time() < end:
        if is_home(win):
            return "브라우저가 직접 전달"
        url = holder.get("url") or _scan_pages(ctx)
        if url:
            # Chrome이 직접 넘겼을 수 있으니 잠깐 기다려 보고, 이미 로그인됐으면 다시 보내지 않는다 (중복 로그인 방지)
            if wait_until(lambda: is_home(win), timeout=5, interval=0.5):
                return "브라우저가 직접 전달"
            say(f"{SCHEME}: 주소를 잡아서 프로그램에 직접 전달: {_mask(url)}")
            return hand_to_app(win, url)
        time.sleep(1)
        remaining = int(end - time.time())
        if remaining % 15 == 0:
            say(f"  ...대기 중 (남은 {remaining}초)")
    trail = " | ".join(holder.get("trail", [])[-4:]) or "기록 없음"
    raise ConnectorError(
        f"로그인 후 {int(timeout)}초 동안 {SCHEME}: 주소를 찾지 못했습니다. (계정·워크스페이스 확인) 마지막 응답: {trail}"
    )


class BrowserDied(Exception):
    """웹 로그인 중 Chrome(또는 Playwright 드라이버)이 꺼짐. 서버 메모리가 모자랄 때 생긴다."""


def free_commit_mb() -> int | None:
    """새 프로그램이 더 쓸 수 있는 메모리(커밋 여유, MB). 이게 바닥나면 Chrome이 뜨다가 죽는다."""
    import ctypes

    class MEMORYSTATUSEX(ctypes.Structure):
        _fields_ = [("dwLength", ctypes.c_ulong), ("dwMemoryLoad", ctypes.c_ulong),
                    ("ullTotalPhys", ctypes.c_ulonglong), ("ullAvailPhys", ctypes.c_ulonglong),
                    ("ullTotalPageFile", ctypes.c_ulonglong), ("ullAvailPageFile", ctypes.c_ulonglong),
                    ("ullTotalVirtual", ctypes.c_ulonglong), ("ullAvailVirtual", ctypes.c_ulonglong),
                    ("ullAvailExtendedVirtual", ctypes.c_ulonglong)]

    status = MEMORYSTATUSEX(dwLength=ctypes.sizeof(MEMORYSTATUSEX))
    if not ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(status)):
        return None
    return int(status.ullAvailPageFile // (1024 * 1024))


def _ctx_alive(ctx) -> bool:
    try:
        return bool(ctx.pages) and ctx.pages[0].evaluate("1") == 1
    except Exception:
        return False


def _web_login_once(win, stage: str, login_url: str, user_id: str, password: str, workspace: str) -> str:
    from playwright.sync_api import sync_playwright

    holder: dict = {}
    with sync_playwright() as p:
        ctx = launch_chrome(p)
        try:
            # 이전 웹 로그인 세션이 남아 있으면 입력 없이 그 계정으로 자동 로그인되므로 매번 지운다
            ctx.clear_cookies()
            _attach_capture(ctx, holder)
            page = ctx.pages[0] if ctx.pages else ctx.new_page()
            _attach_cdp_capture(ctx, page, holder)
            page.goto(login_url, wait_until="domcontentloaded", timeout=30000)
            login_flow(ctx, holder, stage, user_id, password, workspace)
            say(f"로그인 정보 제출 완료, {SCHEME}: 주소 기다리는 중...")
            return wait_handoff(win, ctx, holder)
        except Exception as e:
            # 브라우저가 죽으면 '입력칸을 못 찾음' 같은 엉뚱한 오류로 보이므로 원인을 바로잡는다
            if not _ctx_alive(ctx):
                raise BrowserDied(type(e).__name__) from e
            raise
        finally:
            try:
                ctx.close()
            except Exception:
                pass
            password = None  # noqa: F841


def web_login(win, stage: str, login_url: str, user_id: str, password: str, workspace: str) -> str:
    """웹 로그인. 중간에 브라우저가 꺼지면(메모리 부족 등) 한 번 더 시도한다."""
    for attempt in (1, 2):
        try:
            return _web_login_once(win, stage, login_url, user_id, password, workspace)
        except BrowserDied as e:
            free = free_commit_mb()
            free_text = f"남은 메모리 {free}MB" if free is not None else "남은 메모리 확인 불가"
            if attempt == 2:
                password = None  # noqa: F841
                raise ConnectorError(
                    f"웹 로그인 중 브라우저가 두 번 모두 꺼졌습니다 ({e}, {free_text}). 서버 메모리가 부족한 것 같습니다. "
                    "VS Code처럼 안 쓰는 프로그램을 닫고 다시 시도해 주세요."
                ) from e
            say(f"  브라우저가 중간에 꺼짐 ({e}, {free_text}) → 10초 뒤 웹 로그인 다시 시도")
            time.sleep(10)
            try:
                if is_home(win):  # 꺼지기 전에 이미 로그인이 넘어간 경우
                    return "브라우저가 직접 전달"
            except Exception:
                pass
    raise AssertionError("unreachable")


INSPECT_JS = """() => {
  const vis = e => !!(e.offsetWidth || e.offsetHeight || e.getClientRects().length);
  const out = [];
  for (const e of document.querySelectorAll('input, textarea, select'))
    if (vis(e) && e.type !== 'hidden')
      out.push(`INPUT  type=${e.type} name=${e.name} id=${e.id} placeholder=${e.placeholder || ''}`);
  for (const e of document.querySelectorAll('button, a, [role=button]'))
    if (vis(e) && (e.innerText || '').trim())
      out.push(`BUTTON <${e.tagName.toLowerCase()}> "${e.innerText.trim().replace(/\\s+/g, ' ').slice(0, 50)}"`);
  for (const e of document.querySelectorAll('h1, h2, h3, label'))
    if (vis(e) && (e.innerText || '').trim())
      out.push(`TEXT   <${e.tagName.toLowerCase()}> "${e.innerText.trim().replace(/\\s+/g, ' ').slice(0, 50)}"`);
  return out;
}"""


def cmd_scheme() -> None:
    win = main_window()
    say(f"U+SASE 실행 파일: {app_exe_path(win) or '확인 불가'}")
    say("레지스트리의 URL 프로토콜 중 sase / uplus / connector가 들어간 것:")
    for name, command in _url_protocols():
        if re.search(r"sase|uplus|connector", name + command, re.IGNORECASE):
            say(f"  {name}:  ->  {command}")
    say("RESULT scheme 확인 종료")


def cmd_inspect(stage: str) -> None:
    """입력칸의 종류·이름만 출력하고, 입력된 값은 출력하지 않는다."""
    from playwright.sync_api import sync_playwright

    use_stage_scheme(stage)
    with sync_playwright() as p:
        ctx = launch_chrome(p)
        page = ctx.pages[0] if ctx.pages else ctx.new_page()
        page.goto(login_url(get_console_url(stage), stage))
        while True:
            answer = input("\n화면을 확인할 상태로 두고 Enter (끝내려면 q + Enter): ").strip().lower()
            if answer == "q":
                break
            say(f"--- {urlparse(page.url).path} ---")
            for line in page.evaluate(INSPECT_JS):
                say(line)
        ctx.close()
    say("RESULT inspect 종료")


# ------------------------------------------------------------
# 명령
# ------------------------------------------------------------
def _signature(stage: str, workspace: str, user_id: str) -> dict:
    return {
        "stage": stage,
        "workspace": workspace,
        "user": hashlib.sha256(user_id.lower().encode()).hexdigest()[:16],
    }


def load_state() -> dict:
    """{"logins": {앱: {...그 앱의 마지막 로그인...}}, "exe": {앱: 실행 파일 경로}}"""
    try:
        data = json.loads(STATE_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {}
    if "stage" in data:  # 이전 형식
        data = {"login": data}
    data.pop("login", None)  # 앱 구분 없던 이전 형식은 버린다 (다른 앱 로그인으로 덮여 오판함)
    return data


def login_state(app: str) -> dict | None:
    return load_state().get("logins", {}).get(app)


def set_login_state(app: str, value: dict | None) -> None:
    """앱별로 기록한다. 두 앱은 각자 토큰을 갖고 있어서, 다시 켜면 그 앱의 마지막 계정으로 자동 로그인될 수 있다."""
    state = load_state()
    logins = state.setdefault("logins", {})
    if value:
        logins[app] = value
    else:
        logins.pop(app, None)
    save_state(state)


def save_state(data: dict) -> None:
    STATE_FILE.write_text(json.dumps(data), encoding="utf-8")


# 일반 브라우저(Chrome/Edge)에 남는 커넥터 로그인 탭. 커넥터 앱이 로그인 화면이 될 때 기본 브라우저로 연다.
# 주소에 connector_url= 이 붙은 콘솔 로그인 탭만 닫는다 (사람이 직접 연 콘솔 탭은 건드리지 않음).
# 예: https://sse-qastage.aioncloud.com:50443/auth/login?connector_url=aiconnector%3Aopen%2F
BROWSER_IMAGES = ("chrome.exe", "msedge.exe")
LOGIN_TAB_MARK = "connector_url="


def _browser_address(window) -> str:
    """브라우저 창의 주소창 값 (현재 탭). 웹페이지 안의 입력칸과 구분하려고 주소창 이름으로 고른다."""
    fallback = ""
    for edit in window.descendants(control_type="Edit"):
        name = edit.window_text() or ""
        value = _edit_value(edit)
        if re.search(r"주소|address", name, re.IGNORECASE):
            return value
        if not fallback and value and " " not in value and "." in value:
            fallback = value
    return fallback


def _close_current_tab(window) -> bool:
    """키 입력 없이 현재 탭을 닫는다. 탭이 하나뿐이면 창을 닫는다."""
    tabs = window.descendants(control_type="TabItem")
    if len(tabs) <= 1:
        window.close()
        return True
    for tab in tabs:
        try:
            if not tab.iface_selection_item.CurrentIsSelected:
                continue
        except Exception:
            continue
        for button in tab.descendants(control_type="Button"):
            button.invoke()  # 탭의 닫기(X) 버튼
            return True
    return False


def _is_browser_window(window) -> bool:
    """Chrome/Edge 창인지. VS Code·Discord 같은 Electron 앱도 창 클래스가 같아서 프로세스 이름으로 거른다."""
    try:
        return Path(_process_path(window.process_id())).name.lower() in BROWSER_IMAGES
    except Exception:
        return False


def _tab_selected(tab) -> bool:
    try:
        return bool(tab.iface_selection_item.CurrentIsSelected)
    except Exception:
        return False


def _is_login_tab_url(url: str, host: str) -> bool:
    url = url.lower()
    return bool(host) and host in url and LOGIN_TAB_MARK in url


def _close_login_tabs_in(window, host: str, label: str) -> int:
    """창 하나의 모든 탭을 훑어 커넥터 로그인 탭을 닫는다.

    탭 주소는 그 탭을 선택해야 주소창에서 읽을 수 있어서 하나씩 선택한다. 끝나면 원래 보던 탭으로 돌아간다.
    """
    tabs = window.descendants(control_type="TabItem")
    original = next((t.window_text() for t in tabs if _tab_selected(t)), None)
    closed = 0
    unreadable = 0
    for _ in range(30):  # 닫을 때마다 탭 목록이 바뀌므로 처음부터 다시 훑는다 (무한 반복 방지)
        tabs = window.descendants(control_type="TabItem")
        found = False
        unreadable = 0
        for tab in tabs or [None]:  # 탭 목록이 안 보이면 현재 탭만 확인
            if tab is not None and not _tab_selected(tab):
                tab.iface_selection_item.Select()
                time.sleep(0.3)
            url = _browser_address(window)
            if not url:
                unreadable += 1
            elif _is_login_tab_url(url, host):
                found = True  # 이 탭이 지금 선택돼 있음
                break
        if not found:
            break
        last_tab = len(tabs) <= 1
        if not _close_current_tab(window):
            say(f"  로그인 탭 정리: {label} 탭 닫기 버튼을 찾지 못함")
            break
        closed += 1
        time.sleep(0.7)
        if last_tab:
            return closed  # 탭이 하나뿐이라 창을 닫음
    if unreadable:
        say(f"  로그인 탭 정리: {label} 주소를 읽지 못한 탭 {unreadable}개 (화면 잠김·최소화 가능성)")
    if original:
        for tab in window.descendants(control_type="TabItem"):
            if tab.window_text() == original and not _tab_selected(tab):
                tab.iface_selection_item.Select()
                break
    return closed


def close_browser_login_tabs(host: str) -> int:
    """일반 Chrome/Edge 창의 모든 탭에서 커넥터 로그인 탭(connector_url=)을 닫는다. 닫은 개수를 돌려준다.

    결과와 못 닫은 이유는 진행 기록(connector-last.log)에 남긴다. 주소 자체는 남기지 않는다.
    """
    from pywinauto import Desktop

    host = (host or "").lower()
    try:
        windows = [w for w in Desktop(backend="uia").windows(class_name="Chrome_WidgetWin_1") if _is_browser_window(w)]
    except Exception as e:
        say(f"  로그인 탭 정리: 브라우저 창 목록을 못 읽음 ({type(e).__name__})")
        return 0
    closed = 0
    for index, window in enumerate(windows, 1):
        label = f"창 {index}"
        try:
            closed += _close_login_tabs_in(window, host, label)
        except Exception as e:
            say(f"  로그인 탭 정리: {label} 처리 중 중단 ({type(e).__name__})")
    say(f"  로그인 탭 정리: 브라우저 창 {len(windows)}개, 닫은 탭 {closed}개")
    return closed


def _close_login_tabs_after(stage: str) -> int:
    """connector on이 끝난 뒤(성공·이미 ON·실패 모두) 남은 로그인 탭을 닫는다. 여기서 난 오류로 결과를 바꾸지 않는다."""
    try:
        time.sleep(2)  # 앱이 브라우저를 조금 늦게 여는 경우
        return close_browser_login_tabs(host_of(get_console_url(stage)))
    except Exception as e:
        say(f"  로그인 탭 정리 실패 ({type(e).__name__})")
        return 0


def _today_log(app: str) -> Path | None:
    import datetime

    base = Path(APPS[app]["conf_dir"]).parent  # ...\Monitorapp 또는 ...\LGUplus 상위
    name = f"connector-{datetime.date.today():%Y%m%d}.log"
    found = sorted(Path(APPS[app]["conf_dir"]).glob(f"S-1-5-*/log/{name}"),
                   key=lambda p: p.stat().st_mtime, reverse=True)
    return found[0] if found else None


def app_started_at(app: str) -> float:
    """앱 로그에서 마지막으로 앱이 시작된 시각(epoch). 모르면 0."""
    import datetime

    for line in reversed(app_log_since(app, 0)):
        if "############### Start" in line:
            stamp = re.match(r"﻿?(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})", line).group(1)
            return datetime.datetime.strptime(stamp, "%Y-%m-%d %H:%M:%S").timestamp()
    return 0


def visible_window(app: str, win=None):
    """조작할 수 있는 앱 창. 연결되면 앱이 창을 트레이로 숨기므로, 없으면 실행 파일을 다시 실행해 띄운다."""
    if win is not None and not _win_gone(win):
        return win
    for _ in range(3):  # 창이 숨는 순간과 겹치면 찾은 창이 바로 사라질 수 있어 몇 번 다시 본다
        try:
            w = find_window(1)
            if w is not None and identify_app(w) == app:
                return w
            w = show_app_window(app)
            if w is not None and not _win_gone(w):
                return w
        except Exception:
            pass
        time.sleep(2)
    raise ConnectorError(f"{APPS[app]['label']} 창을 다시 띄우지 못했습니다.")


def app_log_since(app: str, since: float) -> list:
    """앱 로그에서 since(epoch) 이후의 줄들."""
    path = _today_log(app)
    if path is None:
        return []
    try:
        lines = path.read_text(encoding="utf-8", errors="replace").splitlines()
    except Exception:
        return []
    import datetime

    out = []
    for line in lines[-2000:]:
        m = re.match(r"﻿?(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})", line)
        if not m:
            continue
        try:
            ts = datetime.datetime.strptime(m.group(1), "%Y-%m-%d %H:%M:%S").timestamp()
        except ValueError:
            continue
        if ts >= since - 2:
            out.append(line)
    return out


# 앱 로그로 본 연결 결과
CONNECTED = "connected"
VPN_FAILED = "vpn_failed"  # 로그인은 됐지만 터널 수립 실패
LOGGED_OUT = "logged_out"  # 로그인 후 세션을 잃음 (로그인 만료 공지 등)


def connect_state(app: str, since: float) -> str:
    """since 이후 앱 로그의 마지막 결과. 아직 진행 중이면 빈 문자열."""
    state, logged_in = "", False
    for line in app_log_since(app, since):
        if "Login success" in line:
            state, logged_in = "", True
        elif "VPN Connected" in line:
            state = CONNECTED
        elif "VPN Stopped" in line and state == CONNECTED:
            state = ""
        elif "Tunnel Connection Error" in line or "Tunnel start failed (attempt 3/3)" in line:
            state = VPN_FAILED
        elif logged_in and ("Login Status Change Event Occurred - False" in line or "Refresh token expired" in line):
            state, logged_in = LOGGED_OUT, False
    return state


def wait_connected(win, app: str, since: float, timeout: float = 90) -> str:
    """보안 연결 결과를 기다린다: CONNECTED / VPN_FAILED / LOGGED_OUT / 'timeout'.

    화면(상태 텍스트)과 앱 로그를 함께 본다. 실패가 로그에 찍히면 시간 초과까지 기다리지 않는다.
    """
    result = {"state": ""}

    def done() -> bool:
        w = win if (win is not None and not _win_gone(win)) else find_window(0)
        if w is not None:
            try:
                if status_text(w) == ON_TEXT:
                    result["state"] = CONNECTED
                    return True
            except Exception:
                pass
        state = connect_state(app, since)
        if state == VPN_FAILED:  # 앱이 이어서 다시 연결하는 경우가 있어 잠깐 더 본다
            time.sleep(5)
            state = connect_state(app, since)
        result["state"] = state
        return bool(state)

    wait_until(done, timeout=timeout, interval=2.0)
    return result["state"] or "timeout"


def _win_gone(win) -> bool:
    try:
        return not win.exists(timeout=0)
    except Exception:
        return True


class NotConfigured(ConnectorError):
    """이 스테이지는 커넥터 로그인 설정이 없다 (console_url 미설정 등)."""


def connector_on(stage: str, override: dict | None = None) -> tuple[str, str]:
    """앱 전환 + 스테이지 + 로그인 + ON. 끝나면 성공·이미 ON·실패와 상관없이 브라우저에 남은 로그인 탭을 닫는다."""
    try:
        workspace, detail = _connector_on(stage, override)
    except NotConfigured:
        raise  # 콘솔 주소가 없는 스테이지: 열린 탭도 없다
    except Exception:
        _close_login_tabs_after(stage)
        raise
    closed = _close_login_tabs_after(stage)
    if closed:
        detail += f" → 로그인 탭 {closed}개 닫음"
    return workspace, detail


def _connector_on(stage: str, override: dict | None = None) -> tuple[str, str]:
    override = override or {}
    try:
        console_url = get_console_url(stage)
    except ConnectorError as e:
        raise NotConfigured(str(e)) from None
    target = host_of(console_url)
    if "todo" in console_url.lower() or not re.fullmatch(r"[a-z0-9.-]+\.[a-z]{2,}", target or ""):
        raise NotConfigured(f"environments.yaml의 '{stage}' console_url이 아직 설정되지 않았습니다.")
    if override.get("user") and override.get("password"):
        user_id, password = override["user"].strip(), override["password"]
    else:
        user_id, password = credentials(stage)
    workspace = (override.get("workspace") or "").strip() or workspace_for(stage)
    wanted = _signature(stage, workspace, user_id)

    app = app_for_stage(stage)
    use_stage_scheme(stage)
    enable_auto_connect(app)  # 로그인 시 앱이 알아서 ON 하도록 (버튼 클릭 불필요, 화면 불필요)
    say(f"[1] 앱 준비 ({APPS[app]['label']})")
    win, note = ensure_app(app)
    steps = [note] if note else []
    win = visible_window(app, win)
    say(f"[2] 현재 화면: {page_name(win)}")

    already = is_home(win) and login_state(app) == wanted
    if already and (status_text(win) == ON_TEXT or connect_state(app, app_started_at(app)) == CONNECTED):
        # 앱을 켜면서 남은 토큰으로 같은 계정이 자동 로그인·연결된 경우 (연결 후 창은 트레이로 숨는다)
        password = None  # noqa: F841
        steps.append("같은 계정으로 이미 로그인·ON")
        return workspace, " → ".join(steps)
    if is_home(win) and not already:
        say("[3] 로그아웃")
        win = visible_window(app, win)
        logout(win)
        set_login_state(app, None)
        steps.append("로그아웃")
        time.sleep(3)  # 너무 빨리 재로그인하면 "로그인 정보 만료" 공지가 뜨는 경우가 있어 잠시 대기

    since = time.time()
    if not already:
        if not is_home(win):
            say(f"[4] 스테이지 확인 ({target})")
            win, changed = ensure_stage(win, app, target)
            if changed:
                steps.append("스테이지 변경")
            if not wait_until(lambda: is_login(win) or is_home(win), timeout=30):
                raise ConnectorError(f"스테이지 변경 후 앱 화면을 확인하지 못했습니다. (현재: {page_name(win)})")
            if is_home(win):  # 남은 토큰으로 자동 로그인된 경우
                win = visible_window(app, win)  # 자동 연결까지 되면 창이 트레이로 숨는다
                logout(win)
                set_login_state(app, None)
                time.sleep(3)
                if not wait_until(lambda: is_login(win), timeout=15):
                    raise ConnectorError("자동 로그인된 상태에서 로그아웃했지만 로그인 화면으로 돌아오지 않았습니다.")
            time.sleep(STEP_WAIT)
            say("[5] 웹 로그인")
            how = web_login(win, stage, login_url(console_url, stage), user_id, password, workspace)
            set_login_state(app, wanted)
            if take_test_fail("logout"):
                say("  [시험] 로그인 직후 로그아웃 (세션 상실 재현)")
                win = visible_window(app, win)
                logout(win)
            steps.append(f"로그인({workspace or '워크스페이스 입력 없음'}, {how})")
    else:
        steps.append("같은 계정으로 이미 로그인됨")

    # 이미 켜져 있으면 끝. 아니면 자동 연결이 켜지길 기다린다.
    w = win if not _win_gone(win) else find_window(0)
    if w is not None and not _win_gone(w):
        try:
            if status_text(w) == ON_TEXT:
                password = None  # noqa: F841
                steps.append("이미 ON")
                return workspace, " → ".join(steps)
        except Exception:
            pass
    time.sleep(STEP_WAIT)
    try:
        current = win if not _win_gone(win) else find_window(0)
        if current is not None and dismiss_notice(current):
            say("  공지 팝업 닫음")
    except Exception:
        pass
    started = app_started_at(app)
    if already and time.time() - started > 60:
        # 원래 로그인돼 있던 앱은 로그인 이벤트가 새로 없어서 자동 연결이 걸리지 않는다. 기다리지 않고 재로그인한다.
        state = "off"
    else:
        say("[6] 자동 연결 확인 (최대 90초)")
        state = wait_connected(win, app, min(since, started) if already else since, timeout=90)
    if state == CONNECTED:
        steps.append("ON(자동 연결)")
    else:
        # VPN 실패든 로그인 상실이든, '프로그램 종료'는 로그아웃까지 하므로 재시작만으로는 다시 연결되지 않는다.
        # 앱을 다시 켜고 웹 로그인부터 한 번 더 해서 자동 연결을 다시 건다.
        reason = {VPN_FAILED: "VPN 터널 연결 실패", LOGGED_OUT: "로그인 후 세션 상실(로그인 만료)",
                  "off": "로그인돼 있지만 보안 연결 꺼짐"}.get(state, "90초 안에 연결 확인 안 됨")
        say(f"[7] 앱 로그 판정: {reason} → 앱 재시작 후 웹 로그인 재시도")
        exit_app(app, win if not _win_gone(win) else None)
        enable_auto_connect(app)  # 앱이 꺼져 있을 때 다시 확실히 켜 둔다
        since = time.time()
        win = start_app(app)
        if is_home(win):  # 남은 토큰으로 자동 로그인된 경우: 이번 로그인으로 자동 연결이 걸리므로 그대로 기다린다
            how = "앱 자동 로그인"
        else:
            how = web_login(win, stage, login_url(console_url, stage), user_id, password, workspace)
        set_login_state(app, wanted)
        dismiss_notice(win)
        state = wait_connected(win, app, since, timeout=90)
        if state == CONNECTED:
            steps.append(f"{reason} → 재로그인({how}) → ON(자동 연결)")
        else:
            # 그래도 안 되면 화면이 있을 때 한해 버튼으로 시도
            w = visible_window(app, win) if desktop_available() else None
            if w is not None and is_home(w) and turn_on(w):
                steps.append(f"{reason} → 재로그인 → ON(버튼)")
            else:
                password = None  # noqa: F841
                again = {VPN_FAILED: "VPN 터널 연결 실패", LOGGED_OUT: "세션 상실"}.get(state, "시간 초과")
                raise ConnectorError(
                    f"자동 연결이 두 번 모두 실패했습니다. (1차: {reason}, 2차: {again}) 앱 로그를 확인해 주세요. "
                    f"(입력 화면 {'있음' if desktop_available() else '없음'})"
                )
    password = None  # noqa: F841
    return workspace, " → ".join(steps)


def running_app_window():
    """(앱, 창). 연결 후 트레이로 숨은 창도 다시 띄워서 돌려준다."""
    win = find_window(1)
    if win is not None and identify_app(win):
        return identify_app(win), main_window()
    for app in APPS:
        if app_running(app):
            return app, visible_window(app)
    raise ConnectorError("실행 중인 커넥터 앱(창 또는 트레이)을 찾지 못했습니다.")


def connector_logout() -> str:
    app, win = running_app_window()
    if not is_home(win):
        set_login_state(app, None)
        return "이미 로그아웃 상태"
    logout(win)
    set_login_state(app, None)
    return "로그아웃 완료"


def connector_status() -> str:
    app, win = running_app_window()
    login = login_state(app) or {}
    parts = [
        f"앱: {APPS[app]['label'] if app else '알 수 없음'}",
        f"화면: {'로그인됨' if is_home(win) else '로그인 화면' if is_login(win) else page_name(win)}",
        f"상태: {status_text(win) or '-'}",
    ]
    if is_home(win) and login:
        parts.append(f"마지막 로그인: {login.get('stage')} / {login.get('workspace')}")
    return " / ".join(parts)


def cmd_set_credential(key: str) -> None:
    user_id = input(f"[{key}] 로그인 ID: ").strip()
    password = getpass.getpass("비밀번호 (입력해도 화면에 표시되지 않음): ")
    if not user_id or not password:
        raise ConnectorError("ID와 비밀번호가 모두 필요합니다.")
    keyring.set_password(KEYRING_SERVICE, f"{key}:id", user_id)
    keyring.set_password(KEYRING_SERVICE, f"{key}:password", password)
    say(f"RESULT ✅ '{key}' 계정을 Windows 자격 증명 관리자에 저장했습니다.")


def cmd_setup(stage: str) -> None:
    from playwright.sync_api import sync_playwright

    use_stage_scheme(stage)
    url = login_url(get_console_url(stage), stage)
    with sync_playwright() as p:
        ctx = launch_chrome(p)
        page = ctx.pages[0] if ctx.pages else ctx.new_page()
        page.goto(url)
        say("열린 Chrome에서 로그인 페이지가 정상으로 뜨는지 확인해 주세요.")
        input("프로그램이 로그인 상태로 바뀌었으면 Enter를 눌러 주세요...")
        ctx.close()
    say("RESULT ✅ 전용 Chrome 설정 완료")


def main(argv) -> int:
    commands = ("set-credential", "setup", "inspect", "on", "logout", "status", "scheme")
    if len(argv) < 2 or argv[1] not in commands:
        say(__doc__)
        return 2
    cmd = argv[1]
    try:
        RUN_LOG.write_text("", encoding="utf-8")
    except Exception:
        pass
    say(f"connector.py {VERSION} / {cmd} {' '.join(argv[2:3])} / 입력 화면 {'있음' if desktop_available() else '없음'}")
    try:
        if cmd not in ("status", "scheme", "logout") and len(argv) < 3:
            raise ConnectorError(f"사용법: python connector.py {cmd} <stage>")
        if cmd == "set-credential":
            cmd_set_credential(argv[2])
        elif cmd == "setup":
            cmd_setup(argv[2])
        elif cmd == "inspect":
            cmd_inspect(argv[2])
        elif cmd == "scheme":
            cmd_scheme()
        elif cmd == "on":
            override = {}
            if "--stdin" in argv:
                line = sys.stdin.readline()
                override = json.loads(line) if line.strip() else {}
            workspace, detail = connector_on(argv[2], override)
            say(f"진행: {detail}")  # 단계 상세는 기록·봇 콘솔에만 남기고 Discord에는 결과만 보낸다
            say(f"RESULT [{workspace or argv[2]}] 보안 연결 ON 완료✅")
        elif cmd == "logout":
            say(f"RESULT ✅ {connector_logout()}")
        else:
            say(f"RESULT {connector_status()}")
        return 0
    except NotConfigured as e:
        say(f"RESULT ⏭️ 커넥터 설정 없음: {sanitize(str(e))}")
        return 3
    except ConnectorError as e:
        say(f"RESULT ❌ {sanitize(str(e))}")
        return 1
    except Exception as e:
        say(sanitize(traceback.format_exc()))
        if _is_no_desktop(e):
            say(f"RESULT ❌ {NO_DESKTOP_MESSAGE}")
            return 1
        say(f"RESULT ❌ 예상치 못한 오류 ({type(e).__name__}). 봇 콘솔 로그를 확인해 주세요.")
        return 1


if __name__ == "__main__":
    sys.exit(main(sys.argv))