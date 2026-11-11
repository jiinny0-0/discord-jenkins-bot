"""BAT 결과 CSV를 다른 언어(ko/en)로 다시 쓴다. 테스트를 다시 돌리지 않는다.

결과 CSV에서 실행마다 달라지는 값은 Case ID별 Status뿐이고, 나머지 열은 BAT의
tests/case_descriptions.yaml에 있는 고정 문장이다. 그래서 원본 CSV에서 상태만 읽고,
BAT 자신의 결과 파일 작성 코드(bat.report.result_book)로 설명을 원하는 언어로 채운다.

BAT의 venv python으로 실행한다 (봇이 subprocess로 부름):
    <BAT>\\.venv\\Scripts\\python.exe relang_csv.py --bat-dir <BAT> --src in.csv --lang en --out out.csv
"""
from __future__ import annotations

import argparse
import csv
import io
import os
import sys
from pathlib import Path


# 판정 체계 개편(FAIL = 치명) 전의 CSV에는 Status Meaning 열이 없고, 그때의 FAIL은 치명 기준이 아니었다.
OLD_FAIL_TEXT = {"ko": "기대 결과와 다르게 동작함", "en": "Behaved differently from the expected result"}


def read_statuses(data: bytes) -> tuple[dict[str, str], dict[str, dict[str, str]], bool]:
    """원본 CSV → ({Case ID: Status}, {Case ID: 원본 행}, 새 형식 여부). 옛 BAT의 ERROR는 BAT_ERROR로 읽는다."""
    for encoding in ("utf-8-sig", "cp949"):
        try:
            text = data.decode(encoding)
            break
        except UnicodeDecodeError:
            continue
    else:
        raise ValueError("CSV encoding not recognised (UTF-8, CP949)")
    rows = list(csv.DictReader(io.StringIO(text, newline="")))
    if not rows:
        raise ValueError("empty CSV")
    headers = {name.strip().lower(): name for name in rows[0] if name}
    case_col = headers.get("case id") or headers.get("case_id")
    status_col = headers.get("status")
    if not case_col or not status_col:
        raise ValueError("no Case ID / Status column")
    statuses: dict[str, str] = {}
    originals: dict[str, dict[str, str]] = {}
    for row in rows:
        case_id = (row.get(case_col) or "").strip()
        status = (row.get(status_col) or "").strip()
        if not case_id:
            continue
        statuses[case_id] = "BAT_ERROR" if status.upper() == "ERROR" else status
        originals[case_id] = {
            "product": (row.get(headers.get("product", ""), "") or "").strip(),
            "scenario": (row.get(headers.get("scenario", ""), "") or "").strip(),
        }
    return statuses, originals, "status meaning" in headers


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--bat-dir", required=True)
    parser.add_argument("--src", required=True)
    parser.add_argument("--lang", choices=["ko", "en"], required=True)
    parser.add_argument("--out", required=True)
    args = parser.parse_args()

    bat_dir = Path(args.bat_dir).resolve()
    sys.path.insert(0, str(bat_dir))
    os.chdir(bat_dir)
    from bat.report import result_book, status_book

    statuses, originals, new_format = read_statuses(Path(args.src).read_bytes())
    # 원본에 있던 케이스만 (그 뒤에 추가된 케이스를 미실행으로 끼워 넣지 않는다)
    cases = [c for c in status_book.catalog() if c["case_id"] in statuses]
    descriptions = result_book.load_descriptions(result_book.DESCRIPTIONS, args.lang)
    rows = result_book.build_rows(statuses, cases, descriptions, args.lang)

    # 폐기돼 지금 BAT에 없는 케이스: 설명을 만들 수 없으니 원본 값으로 남긴다
    known = {c["case_id"] for c in cases}
    retired = [cid for cid in statuses if cid not in known]
    for case_id in retired:
        rows.append({"case_id": case_id, "status": statuses[case_id], **originals[case_id]})
    if retired:
        print(f"not in the current BAT (kept as in the source): {', '.join(sorted(retired))}", file=sys.stderr)

    if not new_format:
        # 옛 형식: 원본에 없던 Status Meaning은 비우고, FAIL을 '치명(롤백)'으로 쓰지 않는다
        fail_default = result_book.DEFAULT_ACTUAL[args.lang]["FAIL"]
        for row in rows:
            row["status_meaning"] = ""
            if row.get("status") == "FAIL" and row.get("actual_result") == fail_default:
                row["actual_result"] = OLD_FAIL_TEXT[args.lang]

    rows.sort(key=lambda r: r["case_id"])
    out = result_book.write(rows, Path(args.out))
    print(out)
    return 0


if __name__ == "__main__":
    sys.exit(main())
