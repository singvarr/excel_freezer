from dataclasses import dataclass
from pathlib import Path


@dataclass
class TableConfigurationProcessingPaths:
    table: Path
    settings: Path
