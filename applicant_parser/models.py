from __future__ import annotations

from dataclasses import dataclass, field


@dataclass
class ApplicantRecord:
    faculty: str
    fio: str
    email: str
    phone: str
    total_score: int
    subject_score: int
    achievement_score: int
    consent: bool
    direction: str
    quota: str
    quota_raw: str
    priority: str
    group_name: str
    study_form: str
    level: str
    page_number: int


@dataclass
class ApplicantSummary:
    fio: str
    email: str = ""
    phone: str = ""
    total_score: int = 0
    subject_score: int = 0
    achievement_score: int = 0
    consent: bool = False
    priorities: dict[tuple[str, str], str] = field(default_factory=dict)
    source_groups: set[str] = field(default_factory=set)


@dataclass
class ParseResult:
    source_path: str
    faculty: str
    records: list[ApplicantRecord]
    applicants: list[ApplicantSummary]
    groups: set[str]
    pages_count: int
