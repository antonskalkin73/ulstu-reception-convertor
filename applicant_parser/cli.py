from __future__ import annotations

import argparse
import re
import sys
from datetime import datetime
from datetime import date
from pathlib import Path

from .config import INPUT_TYPE_DEANS_LIST
from .pdf_parser import parse_deans_list_pdf, split_result_by_faculty
from .xlsx_writer import write_result_xlsx


def main(argv: list[str] | None = None) -> None:
    _configure_console_output()
    args = _build_parser().parse_args(argv)

    if args.input:
        input_path = _strip_outer_quotes(args.input)
        output_dir = Path(args.output_dir) if args.output_dir else _default_output_dir(INPUT_TYPE_DEANS_LIST)
        _process_deans_list(input_path, output_dir)
        return

    _run_menu()


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Консольный парсер PDF-файлов абитуриентов в XLSX."
    )
    parser.add_argument(
        "-i",
        "--input",
        help="Путь к PDF-файлу. Кавычки по краям допустимы.",
    )
    parser.add_argument(
        "-o",
        "--output-dir",
        help="Каталог для результата. По умолчанию: Результаты/<дата>_СписокДеканам.",
    )
    return parser


def _run_menu() -> None:
    print("Парсер абитуриентов из PDF")
    print("1. Список деканам")
    print("0. Выход")

    choice = input("Выберите вид входного файла: ").strip()
    if choice == "0":
        print("Выход.")
        return
    if choice != "1":
        print("Пока поддерживается только пункт 1: Список деканам.")
        return

    input_path = _strip_outer_quotes(input("Введите путь к PDF-файлу: "))
    output_dir = _default_output_dir(INPUT_TYPE_DEANS_LIST)
    _process_deans_list(input_path, output_dir)


def _process_deans_list(input_path: str, output_dir: Path) -> None:
    print("Читаю PDF...")
    parse_result = parse_deans_list_pdf(input_path)
    faculty_results = split_result_by_faculty(parse_result)

    source_stem = _safe_filename(Path(parse_result.source_path).stem)

    print("Формирую XLSX...")
    result_paths = []
    for faculty_result in faculty_results:
        faculty_name = _safe_filename(faculty_result.faculty)
        output_path = _available_output_path(output_dir / f"{faculty_name}_полная_выгрузка_{source_stem}.xlsx")
        result_paths.append(write_result_xlsx(faculty_result, output_path))

    print("Готово.")
    print(f"Факультетов: {len(faculty_results)}")
    print(f"Всего конкурсных групп: {len(parse_result.groups)}")
    print(f"Всего строк заявлений: {len(parse_result.records)}")
    for faculty_result, result_path in zip(faculty_results, result_paths):
        print(
            f"{faculty_result.faculty}: групп {len(faculty_result.groups)}, "
            f"заявлений {len(faculty_result.records)}, "
            f"уникальных абитуриентов {len(faculty_result.applicants)}"
        )
        print(f"Файл: {result_path}")


def _default_output_dir(input_type: str) -> Path:
    folder_name = f"{date.today().isoformat()}_{_safe_filename(input_type)}"
    return Path("Результаты") / folder_name


def _strip_outer_quotes(value: str) -> str:
    text = value.strip()
    if len(text) >= 2 and text[0] == text[-1] and text[0] in {"'", '"'}:
        return text[1:-1].strip()
    return text


def _safe_filename(value: str) -> str:
    text = re.sub(r'[<>:"/\\|?*]+', "_", value).strip().replace(" ", "")
    return text or "result"


def _available_output_path(path: Path) -> Path:
    if not path.exists():
        return path

    timestamp = datetime.now().strftime("%H%M%S")
    candidate = path.with_name(f"{path.stem}_{timestamp}{path.suffix}")
    counter = 2
    while candidate.exists():
        candidate = path.with_name(f"{path.stem}_{timestamp}_{counter}{path.suffix}")
        counter += 1
    return candidate


def _configure_console_output() -> None:
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8")
