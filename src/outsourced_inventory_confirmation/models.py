from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from pathlib import Path


@dataclass(slots=True)
class SourceFile:
    path: Path
    site_name: str


@dataclass(slots=True)
class InventoryRow:
    site_name: str
    vendor_name: str
    material_code: str
    color_code: str
    material_name: str


@dataclass(slots=True)
class VendorPreview:
    vendor_name: str
    rows: list[InventoryRow]


@dataclass(slots=True)
class RenderedDocument:
    vendor_name: str
    report_date: date
    output_dir: Path
    xlsx_path: Path
    pdf_path: Path
