from dataclasses import dataclass
from pathlib import Path


@dataclass
class FreezeSheetSettings:
    name: str
    preserve_sheet: bool
    cells_with_preserved_formulae: list[str]


@dataclass
class FreezeTableConfiguration:
    source: Path
    destination: Path
    sheet_settings: list[FreezeSheetSettings]
