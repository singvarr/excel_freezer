from pathlib import Path

from excel_freezer.services.dialog_manager import DialogManager
from excel_freezer.services.excel_freezer import ExcelFreezer


def freeze_excel_values():
    dialog_manager = DialogManager()

    input_path = dialog_manager.request_table_to_freeze()

    if not input_path:
        return

    excel_freezer = ExcelFreezer(source=Path(input_path))

    output_path = dialog_manager.request_path_for_save(
        default_file_name=excel_freezer.default_output_file_name,
    )

    if not output_path:
        return

    try:
        excel_freezer.run()
        dialog_manager.show_success_message(message=f"Готово!\nФайл збережено як:\n{output_path}")
    except Exception as error:
        dialog_manager.show_error_message(error=error)


if __name__ == "__main__":
    freeze_excel_values()