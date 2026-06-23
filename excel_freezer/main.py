from pathlib import Path

from excel_freezer.services.configuration import ConfigurationService
from excel_freezer.services.dialog_manager import DialogManager
from excel_freezer.services.excel_freezer import ExcelFreezer

if __name__ == "__main__":
    try:
        configuration_service = ConfigurationService()
        dialog_manager = DialogManager()

        table_processing_paths = configuration_service.bootstrap()
        input_path = dialog_manager.request_table_to_freeze()

        if not input_path:
            raise Exception("No table selected")

        table_path = Path(input_path)
        default_output_file_name = ExcelFreezer.get_default_output_file_name(source=table_path)
        output_path = dialog_manager.request_save_path(default_file_name=default_output_file_name)

        if not output_path:
            raise Exception("No output path is selected")

        freeze_table_configuration = configuration_service.get_freeze_table_configuration(
            source=table_path,
            destination=Path(output_path),
        )

        excel_freezer = ExcelFreezer(configuration=freeze_table_configuration)
        excel_freezer.run()

        dialog_manager.show_success_message(message=f"Готово!\nФайл збережено як:\n{output_path}")
    except Exception as error:
        if dialog_manager:
            dialog_manager.show_error_message(error=error)
