from __future__ import annotations

import argparse
import csv
import io
import json
import pathlib
import re
import shutil
from dataclasses import dataclass
from typing import Iterable

import requests
from openpyxl import Workbook
from PIL import Image


PHOTO_URL_TEMPLATE = "https://zhjw.fjcuc.cn/photo/{student_id}.jpg"
STUDENT_ID_RE = re.compile(r"^(124\d{9})$")


@dataclass
class StudentRow:
    student_id: str
    name: str


def normalize_text(value: object) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    return "" if text.lower() == "nan" else text


def read_utf8_text(path: pathlib.Path) -> str:
    return path.read_text(encoding="utf-8-sig")


def parse_text_line(line: str) -> StudentRow | None:
    text = normalize_text(line)
    if not text or text.startswith("#"):
        return None

    parts = [part.strip() for part in re.split(r"[\s,\t]+", text, maxsplit=1) if part.strip()]
    if len(parts) != 2:
        return None

    student_id, name = parts
    if not STUDENT_ID_RE.match(student_id):
        return None

    return StudentRow(student_id=student_id, name=name)


def load_students_from_text(path: pathlib.Path) -> list[StudentRow]:
    rows: list[StudentRow] = []
    for line in read_utf8_text(path).splitlines():
        row = parse_text_line(line)
        if row:
            rows.append(row)
    return rows


def load_students_from_csv(path: pathlib.Path) -> list[StudentRow]:
    rows: list[StudentRow] = []
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        for record in reader:
            student_id = normalize_text(record.get("student_id") or record.get("学号"))
            name = normalize_text(record.get("name") or record.get("姓名"))
            if STUDENT_ID_RE.match(student_id) and name:
                rows.append(StudentRow(student_id=student_id, name=name))
    return rows


def load_students_from_json(path: pathlib.Path) -> list[StudentRow]:
    payload = json.loads(read_utf8_text(path))
    rows: list[StudentRow] = []
    if not isinstance(payload, list):
        return rows
    for item in payload:
        if not isinstance(item, dict):
            continue
        student_id = normalize_text(item.get("student_id") or item.get("学号"))
        name = normalize_text(item.get("name") or item.get("姓名"))
        if STUDENT_ID_RE.match(student_id) and name:
            rows.append(StudentRow(student_id=student_id, name=name))
    return rows


def load_students(student_list_path: pathlib.Path) -> list[StudentRow]:
    suffix = student_list_path.suffix.lower()
    if suffix == ".csv":
        rows = load_students_from_csv(student_list_path)
    elif suffix == ".json":
        rows = load_students_from_json(student_list_path)
    else:
        rows = load_students_from_text(student_list_path)

    unique_rows: list[StudentRow] = []
    seen_ids: set[str] = set()
    for row in rows:
        if row.student_id in seen_ids:
            continue
        seen_ids.add(row.student_id)
        unique_rows.append(row)
    return unique_rows


def safe_filename(student: StudentRow, ext: str = ".jpg") -> str:
    return f"{student.student_id}{student.name}{ext}"


def csv_report_path(out_dir: pathlib.Path) -> pathlib.Path:
    return out_dir / "_photo_download_report.csv"


def folder_named_report_path(out_dir: pathlib.Path) -> pathlib.Path:
    folder_name = out_dir.resolve().name or "report"
    return out_dir / f"{folder_name}.xlsx"


def write_csv_report(
    report_path: pathlib.Path,
    fieldnames: list[str],
    rows: list[dict[str, object]],
) -> pathlib.Path:
    with report_path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)
    return report_path


def write_xlsx_report(
    report_path: pathlib.Path,
    columns: list[tuple[str, str]],
    rows: list[dict[str, object]],
) -> pathlib.Path:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "report"
    worksheet.append([header for _, header in columns])

    for row in rows:
        worksheet.append([row.get(field) for field, _ in columns])

    workbook.save(report_path)
    return report_path


def download_photos(students: Iterable[StudentRow], out_dir: pathlib.Path) -> pathlib.Path:
    out_dir.mkdir(parents=True, exist_ok=True)
    csv_path = csv_report_path(out_dir)
    xlsx_path = folder_named_report_path(out_dir)
    csv_columns = [
        ("seq", "序号"),
        ("student_id", "学号"),
        ("name", "姓名"),
        ("url", "照片地址"),
        ("status", "状态"),
        ("bytes", "字节数"),
    ]
    xlsx_columns = csv_columns[:3]
    rows = []

    session = requests.Session()
    session.verify = False

    for idx, student in enumerate(students, start=1):
        url = PHOTO_URL_TEMPLATE.format(student_id=student.student_id)
        dst = out_dir / safe_filename(student)
        status = "downloaded"
        byte_count = 0

        try:
            response = session.get(url, timeout=20)
            response.raise_for_status()
            byte_count = len(response.content)
            image = Image.open(io.BytesIO(response.content))
            image.load()
            dst.write_bytes(response.content)
        except Exception:
            status = "failed"

        if dst.exists() and byte_count > 0 and byte_count < 12000:
            status = "bad_content_9664" if byte_count == 9664 else "bad_content"

        rows.append(
            {
                "seq": idx,
                "student_id": student.student_id,
                "name": student.name,
                "url": url,
                "status": status if status != "downloaded" or byte_count > 0 else "failed",
                "bytes": byte_count,
            }
        )

    write_csv_report(csv_path, [field for field, _ in csv_columns], rows)
    return write_xlsx_report(xlsx_path, xlsx_columns, rows)


def copy_wechat_cache(cache_dir: pathlib.Path, out_dir: pathlib.Path) -> pathlib.Path:
    out_dir.mkdir(parents=True, exist_ok=True)
    copied = []

    for path in sorted(cache_dir.glob("f_*")):
        try:
            with Image.open(path) as img:
                if img.format != "JPEG" or img.size != (480, 640):
                    continue
        except Exception:
            continue

        dst = out_dir / f"{path.name}.jpg"
        shutil.copy2(path, dst)
        copied.append(dst.name)

    manifest = out_dir / "_cache_raw_files.txt"
    manifest.write_text("\n".join(copied), encoding="utf-8")
    return manifest


def delete_cache_copies(out_dir: pathlib.Path) -> pathlib.Path:
    out_dir.mkdir(parents=True, exist_ok=True)
    report_path = out_dir / "_cache_cleanup_report.xlsx"
    columns = [
        ("file", "文件名"),
        ("status", "状态"),
    ]
    rows = []

    for path in sorted(out_dir.glob("f_*.jpg")):
        try:
            path.unlink()
            rows.append({"file": path.name, "status": "deleted"})
        except Exception:
            rows.append({"file": path.name, "status": "failed"})

    return write_xlsx_report(report_path, columns, rows)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Batch photo helper for student photos.")
    parser.add_argument(
        "--list",
        type=pathlib.Path,
        help="Student list file from mini program, supports .txt .csv .json.",
    )
    parser.add_argument("--out", type=pathlib.Path, required=True, help="Output folder.")
    parser.add_argument("--download", action="store_true", help="Download photos by student id.")
    parser.add_argument("--copy-cache", type=pathlib.Path, help="Copy WeChat cache JPEGs into the output folder.")
    parser.add_argument(
        "--delete-cache-copies",
        action="store_true",
        help="Delete cached copy files named like f_000xxx.jpg from the output folder.",
    )
    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()

    if args.download:
        if not args.list:
            parser.error("--download requires --list")
        students = load_students(args.list)
        report = download_photos(students, args.out)
        print(report)

    if args.copy_cache:
        manifest = copy_wechat_cache(args.copy_cache, args.out)
        print(manifest)

    if args.delete_cache_copies:
        report = delete_cache_copies(args.out)
        print(report)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
