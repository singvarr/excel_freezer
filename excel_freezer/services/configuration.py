import os
from typing import Optional
from pathlib import Path
from json import load

from dotenv import load_dotenv

from excel_freezer.models.table_configuration_processing_paths import (
    TableConfigurationProcessingPaths,
)
from excel_freezer.models.freeze_table_configuration import (
    FreezeSheetSettings,
    FreezeTableConfiguration,
)


class ConfigurationService:
    _PATHS_SEPARATOR = ";"
    _REQUIRED_ENV_CONFIGURATION = [
        "FREEZE_TABLE_CONFIGURATIONS_DIR",
        "TABLE_CONFIGURATION_PROCESSING_PATHS",
    ]

    @property
    def table_configuration_processing_paths(self) -> list[TableConfigurationProcessingPaths]:
        configuration_str = os.environ["TABLE_CONFIGURATION_PROCESSING_PATHS"]

        if not configuration_str:
            return []

        configurations = configuration_str.split(self._PATHS_SEPARATOR)

        if len(configurations) % 2:
            raise Exception("Each table should have own configuration")

        result = []

        for i in range(len(configurations)):
            if i % 2:
                continue

            configuration = TableConfigurationProcessingPaths(
                table=Path(configurations[i]),
                settings=Path(configurations[i + 1]),
            )

            result.append(configuration)

        return result

    def _check_required_env_variables(self):
        for variable in self._REQUIRED_ENV_CONFIGURATION:
            if variable not in os.environ:
                raise Exception(f"Variable {variable} is missing in .env")

    def get_freeze_table_configuration(
        self,
        source: Path,
        destination: Path,
    ) -> Optional[FreezeTableConfiguration]:
        configuration = next(
            (
                config
                for config in self.table_configuration_processing_paths
                if config.table == source
            ),
            None,
        )

        if configuration:
            with open(configuration.settings, "r", encoding="utf-8") as settings:
                freeze_table_config = load(settings)

            return FreezeTableConfiguration(
                source=source,
                destination=destination,
                sheet_settings=[FreezeSheetSettings(**config) for config in freeze_table_config],
            )

        return FreezeTableConfiguration(source=source, destination=destination, sheet_settings=[])

    def bootstrap(self) -> list[TableConfigurationProcessingPaths]:
        load_dotenv()
        self._check_required_env_variables()

        if not os.path.isdir(os.environ["FREEZE_TABLE_CONFIGURATIONS_DIR"]):
            raise Exception("No directory with configurations exists")

        return self.table_configuration_processing_paths
