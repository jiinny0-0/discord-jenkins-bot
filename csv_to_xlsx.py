"""BAT 결과 CSV를 현황 엑셀(reports/bat-test-status.xlsx '현황' 시트)과 같은 서식의 xlsx로 변환한다.

- 원본 CSV 파일은 건드리지 않고, 메모리에서 xlsx를 만들어 bytes로 돌려준다.
- 값은 CSV에 적힌 글자 그대로 넣는다 (숫자/날짜 자동 변환 없음, '='로 시작해도 수식으로 처리하지 않음).

서식 (현황 엑셀 기준)
1. 상태 칸 색: Status 열과 망별 Status 열(qa Status, lg Status 등)의 칸에만 값에 따라 배경/글자색
2. 머리글: 기본 열은 4472C4, 망별 열(qa Status, qa Last Run 등)은 2F5597, 흰색 굵은 글씨
3. 표 스타일: 시트 전체를 엑셀 표(TableStyleMedium2, 줄무늬)로 만들고 상태 칸 색을 그 위에 덧칠
"""

import csv
import io
import re
import unicodedata

from openpyxl import Workbook
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

# 상태별 색: (배경, 글자) — status_book.py와 동일
STATUS_STYLES = {
    "PASS": ("C6EFCE", "006100"),
    "FAIL": ("FFC7CE", "9C0006"),
    "WARN": ("FCE4D6", "C65911"),
    "UNVERIFIED": ("DDEBF7", "1F4E78"),
    "BAT_ERROR": ("F4B084", "833C0B"),
    "ERROR": ("F4B084", "833C0B"),  # 옛 BAT의 BAT_ERROR
    "SKIP": ("FFEB9C", "9C6500"),
    "미실행": ("D9D9D9", "404040"),
}

# 상태 열로 볼 머리글 (소문자 비교). 'qa Status'처럼 ' status'로 끝나는 열도 포함
STATUS_HEADER_NAMES = {"status", "result", "verdict", "상태", "결과", "판정"}

# 망별 열: 머리글이 망 이름으로 시작하는 열 (예: qa Status, lg Last Run)
NETWORK_HEADER_PATTERN = re.compile(r"^(qa|real|test|lguplus|lg-preprod|lg)\s+\S", re.IGNORECASE)

HEADER_FILL_BASE = PatternFill("solid", fgColor="4472C4")
HEADER_FILL_NETWORK = PatternFill("solid", fgColor="2F5597")
HEADER_FONT = Font(bold=True, color="FFFFFF")

EXCEL_CELL_LIMIT = 32767
MIN_WIDTH, MAX_WIDTH = 8, 60


def decode_csv(data: bytes) -> str:
    """UTF-8(BOM 포함)을 먼저 시도하고, 실패하면 CP949(한글 Windows)로 읽는다."""
    for encoding in ("utf-8-sig", "cp949"):
        try:
            return data.decode(encoding)
        except UnicodeDecodeError:
            continue
    raise ValueError("CSV 인코딩을 확인할 수 없습니다 (UTF-8, CP949 모두 실패)")


def normalize_status(value: str) -> str | None:
    upper = (value or "").strip().upper()
    for status in STATUS_STYLES:
        if upper == status.upper() or upper.startswith(status.upper()):  # PASSED, FAILED 등 포함
            return status
    return None


def is_status_header(name: str) -> bool:
    lowered = name.strip().lower()
    return lowered in STATUS_HEADER_NAMES or lowered.endswith(" status")


def find_status_columns(header: list[str], rows: list[list[str]]) -> list[int]:
    by_name = [index for index, name in enumerate(header) if is_status_header(name)]
    if by_name:
        return by_name
    # 이름으로 못 찾으면, 값의 80% 이상이 상태값인 열을 찾는다
    best, best_ratio = None, 0.0
    for index in range(len(header)):
        values = [row[index] for row in rows if index < len(row) and row[index].strip()]
        if not values:
            continue
        ratio = sum(normalize_status(v) is not None for v in values) / len(values)
        if ratio > best_ratio:
            best, best_ratio = index, ratio
    return [best] if best is not None and best_ratio >= 0.8 else []


def unique_headers(header: list[str]) -> list[str]:
    """엑셀 표는 머리글이 비어 있거나 중복되면 파일이 손상되므로 정리한다."""
    result: list[str] = []
    seen: set[str] = set()
    for index, name in enumerate(header, start=1):
        base = " ".join(ILLEGAL_CHARACTERS_RE.sub("", name).split()) or f"열{index}"
        candidate, n = base, 2
        while candidate.lower() in seen:
            candidate = f"{base} ({n})"
            n += 1
        seen.add(candidate.lower())
        result.append(candidate)
    return result


def display_width(text: str) -> int:
    """한글 등 전각 문자는 2칸으로 계산한다. 여러 줄이면 가장 긴 줄 기준."""
    longest = 0
    for line in str(text).splitlines() or [""]:
        width = sum(2 if unicodedata.east_asian_width(ch) in ("W", "F") else 1 for ch in line)
        longest = max(longest, width)
    return longest


def clean_value(value: str) -> tuple[str, bool]:
    """엑셀에 넣을 수 없는 제어 문자를 지우고 셀 길이 제한을 맞춘다."""
    text = ILLEGAL_CHARACTERS_RE.sub("", value)
    if len(text) > EXCEL_CELL_LIMIT:
        return text[: EXCEL_CELL_LIMIT - 20] + "\n...(이하 생략)", True
    return text, False


def csv_to_xlsx(data: bytes, sheet_title: str = "BAT 결과") -> tuple[bytes, list[str]]:
    """CSV bytes → (xlsx bytes, 주의사항 목록)."""
    notes: list[str] = []
    rows = list(csv.reader(io.StringIO(decode_csv(data), newline="")))
    rows = [row for row in rows if any(cell.strip() for cell in row)]
    if not rows:
        raise ValueError("CSV에 내용이 없습니다")

    column_count = max(len(row) for row in rows)
    rows = [row + [""] * (column_count - len(row)) for row in rows]
    header = unique_headers(rows[0])
    body = rows[1:]
    status_cols = set(find_status_columns(header, body))

    wb = Workbook()
    ws = wb.active
    ws.title = sheet_title

    # 머리글
    widths = []
    for c, name in enumerate(header, start=1):
        cell = ws.cell(row=1, column=c, value=name)
        cell.data_type = "s"
        cell.fill = HEADER_FILL_NETWORK if NETWORK_HEADER_PATTERN.match(name) else HEADER_FILL_BASE
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        widths.append(display_width(name))

    # 본문 (표 줄무늬가 보이도록 상태 칸 외에는 배경을 칠하지 않음)
    truncated = 0
    for r, row in enumerate(body, start=2):
        for c, raw in enumerate(row, start=1):
            value, was_cut = clean_value(raw)
            truncated += was_cut
            cell = ws.cell(row=r, column=c)
            cell.value = value
            cell.data_type = "s"  # 항상 텍스트로 (수식/숫자 변환 방지)
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            if c - 1 in status_cols:
                status = normalize_status(raw)
                if status:
                    fill, font = STATUS_STYLES[status]
                    cell.fill = PatternFill("solid", fgColor=fill)
                    cell.font = Font(bold=True, color=font)
                    cell.alignment = Alignment(horizontal="center", vertical="top", wrap_text=True)
            widths[c - 1] = max(widths[c - 1], display_width(value))

    for index, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(index)].width = min(max(width + 2, MIN_WIDTH), MAX_WIDTH)
    ws.freeze_panes = "A2"  # 머리글 고정

    # 엑셀 표 (줄무늬, 정렬·필터 드롭다운). 데이터 행이 없으면 표를 만들 수 없음
    if body:
        table = Table(displayName="BATResult", ref=f"A1:{get_column_letter(column_count)}{len(body) + 1}")
        table.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium2", showRowStripes=True, showColumnStripes=False
        )
        ws.add_table(table)
    else:
        notes.append("결과 행이 없어 표 서식은 적용하지 않음")

    if not status_cols:
        notes.append("상태 열을 찾지 못해 상태 색은 적용하지 않음")
    if truncated:
        notes.append(f"엑셀 셀 길이 제한으로 {truncated}개 칸 일부 생략")

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue(), notes