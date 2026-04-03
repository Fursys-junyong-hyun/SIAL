from __future__ import annotations

from pathlib import Path


def project_root() -> Path:
    return Path(__file__).resolve().parents[2]


def default_template_path() -> Path:
    root = project_root()
    candidates = [
        root / "resources" / "template.xlsx",
        root / "SAMPLE" / "(양식)자재확인서_업체명_261Q.xlsx",
    ]
    for candidate in candidates:
        if candidate.exists():
            return candidate
    raise FileNotFoundError("양식 파일을 찾을 수 없습니다.")
