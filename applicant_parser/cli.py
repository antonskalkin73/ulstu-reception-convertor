from __future__ import annotations

import argparse
import re
import sys
from datetime import date, datetime
from pathlib import Path
from tempfile import NamedTemporaryFile

from .config import ensure_short_names_file
from .models import ParseResult
from .pdf_parser import parse_deans_list_pdf, split_result_by_faculty
from .xlsx_writer import (
    ExistingWorkbookData,
    PriorityCheckResult,
    compare_priorities,
    read_existing_workbook,
    write_result_xlsx,
)


def main(argv: list[str] | None = None) -> None:
    _configure_console_output()
    short_names_path = ensure_short_names_file()
    args = _build_parser().parse_args(argv)

    if args.input:
        input_path = _strip_outer_quotes(args.input)
        output_dir = Path(args.output_dir) if args.output_dir else _default_output_dir()
        existing_path = _strip_outer_quotes(args.existing) if args.existing else ""
        _process_deans_list(input_path, output_dir, args.faculty, existing_path)
        return

    try:
        _run_menu(short_names_path)
    except KeyboardInterrupt:
        print("\nРабота отменена пользователем.")
    except Exception as exc:
        print(f"\nОшибка: {exc}")
    finally:
        _wait_for_keypress()


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
        help="Каталог для результата. По умолчанию: Результаты.",
    )
    parser.add_argument(
        "-f",
        "--faculty",
        help="Код факультета для выгрузки. Пусто или 0 - создать файлы для всех факультетов.",
    )
    parser.add_argument(
        "-e",
        "--existing",
        help="Путь к существующему XLSX-файлу для обновления с сохранением примечаний.",
    )
    return parser


def _run_menu(short_names_path: Path) -> None:
    print("Парсер абитуриентов из PDF")
    print(f"Файл сокращений направлений: {short_names_path}")
    print("Введите код факультета, например ФИСТ.")
    print("Оставьте пустым или введите 0, чтобы создать файлы для всех факультетов.")

    faculty = input("Факультет: ").strip()
    input_path = _strip_outer_quotes(input("Введите путь к PDF-файлу: "))
    existing_path = _strip_outer_quotes(
        input("Путь к существующему XLSX (Enter - создать новый): ")
    )
    output_dir = _default_output_dir()
    _process_deans_list(input_path, output_dir, faculty, existing_path)


def _process_deans_list(
    input_path: str,
    output_dir: Path,
    faculty: str | None = None,
    existing_path: str | None = None,
) -> None:
    existing_data = read_existing_workbook(existing_path) if existing_path else None
    faculty = _faculty_for_update(faculty, existing_data)

    print("Читаю PDF...")
    parse_result = parse_deans_list_pdf(input_path)
    faculty_results = split_result_by_faculty(parse_result)
    faculty_results = _filter_faculty_results(faculty_results, faculty)
    if not faculty_results:
        available = ", ".join(sorted({result.faculty for result in split_result_by_faculty(parse_result)}))
        print(f"Факультет не найден. Доступные факультеты: {available}")
        return

    if existing_data is not None and len(faculty_results) != 1:
        raise ValueError(
            "Для обновления существующего XLSX должен быть выбран один факультет."
        )

    print("Формирую XLSX...")
    result_paths = []
    priority_checks: list[PriorityCheckResult | None] = []
    result_date = date.today().isoformat()
    for faculty_result in faculty_results:
        if existing_data is not None:
            result_path, priority_check = _update_existing_workbook(
                faculty_result,
                existing_data,
            )
        else:
            faculty_name = _safe_filename(faculty_result.faculty)
            output_path = _available_output_path(
                output_dir / f"Абитуриенты_{faculty_name}_{result_date}.xlsx"
            )
            result_path = write_result_xlsx(faculty_result, output_path)
            priority_check = None
        result_paths.append(result_path)
        priority_checks.append(priority_check)

    print("Готово.")
    print(f"Факультетов: {len(faculty_results)}")
    print(f"Всего конкурсных групп: {len(parse_result.groups)}")
    print(f"Всего строк заявлений: {len(parse_result.records)}")
    for faculty_result, result_path, priority_check in zip(
        faculty_results,
        result_paths,
        priority_checks,
    ):
        print(
            f"{faculty_result.faculty}: групп {len(faculty_result.groups)}, "
            f"заявлений {len(faculty_result.records)}, "
            f"уникальных абитуриентов {len(faculty_result.applicants)}"
        )
        print(f"Файл: {result_path}")
        if priority_check is not None:
            print(
                "Проверка приоритетов: "
                f"изменились у {priority_check.changed}, "
                f"без изменений у {priority_check.unchanged}, "
                f"новых абитуриентов {priority_check.added}, "
                f"нет в новом PDF {priority_check.missing}."
            )


def _faculty_for_update(
    requested_faculty: str | None,
    existing_data: ExistingWorkbookData | None,
) -> str | None:
    if existing_data is None or not existing_data.faculty:
        return requested_faculty

    requested = (requested_faculty or "").strip()
    if not requested or requested == "0":
        return existing_data.faculty
    if requested.casefold() != existing_data.faculty.casefold():
        raise ValueError(
            f"Выбран факультет '{requested}', а существующий файл относится "
            f"к факультету '{existing_data.faculty}'."
        )
    return requested


def _update_existing_workbook(
    faculty_result: ParseResult,
    existing_data: ExistingWorkbookData,
) -> tuple[Path, PriorityCheckResult]:
    existing_path = existing_data.path
    existing_path.parent.mkdir(parents=True, exist_ok=True)

    with NamedTemporaryFile(
        prefix=f".{existing_path.stem}_",
        suffix=".xlsx",
        dir=existing_path.parent,
        delete=False,
    ) as temporary_file:
        temporary_path = Path(temporary_file.name)

    try:
        write_result_xlsx(
            faculty_result,
            temporary_path,
            notes_by_applicant=existing_data.notes_by_applicant,
        )
        current_data = read_existing_workbook(temporary_path)
        priority_check = compare_priorities(existing_data, current_data)
        temporary_path.replace(existing_path)
    finally:
        temporary_path.unlink(missing_ok=True)

    return existing_path, priority_check


def _filter_faculty_results(faculty_results: list, faculty: str | None) -> list:
    faculty_code = (faculty or "").strip()
    if not faculty_code or faculty_code == "0":
        return faculty_results
    return [result for result in faculty_results if result.faculty.casefold() == faculty_code.casefold()]


def _default_output_dir() -> Path:
    return Path("Результаты")


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


def _wait_for_keypress() -> None:
    print("\nНажмите любую клавишу для выхода...")
    try:
        import msvcrt

        msvcrt.getwch()
    except (ImportError, OSError):
        input()
