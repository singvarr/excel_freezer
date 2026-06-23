from dataclasses import dataclass
from pathlib import Path


@dataclass
class TableProcessingPaths:
    table: Path
    settings: Path
