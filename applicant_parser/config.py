from __future__ import annotations

import configparser
from dataclasses import dataclass
from pathlib import Path
import sys


INPUT_TYPE_DEANS_LIST = "Список деканам"
SHORT_NAMES_FILENAME = "direction_short_names.ini"

DIRECTIONS = [
    "Информатика и вычислительная техника",
    "Программная инженерия",
    "Прикладная информатика",
    "Информационные системы и технологии",
    "Приборостроение",
    "Прикладная математика",
    "Математическое обеспечение и администрирование информационных систем",
]

DIRECTION_SHORT_NAMES = {
    "Информатика и вычислительная техника": "ИВТ",
    "Программная инженерия": "ПИ",
    "Прикладная информатика": "ИСЭ",
    "Информационные системы и технологии": "ИСТ",
    "Приборостроение": "ПС",
    "Прикладная математика": "ПМ",
    "Математическое обеспечение и администрирование информационных систем": "МО",
}

SHORT_NAMES_TEMPLATE = """# Сокращения направлений для XLSX-выгрузки.
# Файл можно редактировать в Блокноте.
# Формат:
# [КОД_ФАКУЛЬТЕТА]
# Полное название направления = Короткое название
#
# Если для факультета или направления нет строки, приложение создаст сокращение автоматически.

[ФИСТ]
Информатика и вычислительная техника = ИВТ
Программная инженерия = ПИ
Прикладная информатика = ИСЭ
Информационные системы и технологии = ИСТ
Приборостроение = ПС
Прикладная математика = ПМ
Математическое обеспечение и администрирование информационных систем = МО
"""


def get_app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(sys.argv[0]).resolve().parent


def get_short_names_path() -> Path:
    return get_app_dir() / SHORT_NAMES_FILENAME


def ensure_short_names_file() -> Path:
    path = get_short_names_path()
    if not path.exists():
        path.write_text(SHORT_NAMES_TEMPLATE, encoding="utf-8")
    return path


def load_direction_short_names() -> dict[str, dict[str, str]]:
    path = ensure_short_names_file()
    parser = configparser.ConfigParser()
    parser.optionxform = str
    parser.read(path, encoding="utf-8")

    short_names: dict[str, dict[str, str]] = {}
    for section in parser.sections():
        short_names[section.strip()] = {
            direction.strip(): short_name.strip()
            for direction, short_name in parser.items(section)
            if direction.strip() and short_name.strip()
        }
    return short_names

QUOTAS = [
    "Бюджет",
    "ОтдельнаяКвота",
    "ОсобаяКвота",
    "Платно",
    "Целевая",
    "Другое",
]

PRIORITY_SUFFIX = {
    "Бюджет": "",
    "ОтдельнаяКвота": "ОК",
    "ОсобаяКвота": "ОсК",
    "Платно": "П",
    "Целевая": "Ц",
    "Другое": "?",
}


@dataclass(frozen=True)
class GroupInfo:
    faculty: str
    raw_name: str
    direction: str
    quota: str
    quota_raw: str
    study_form: str
    level: str
