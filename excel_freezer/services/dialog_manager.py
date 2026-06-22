import tkinter as tk
from tkinter import filedialog, messagebox


class DialogManager:
    def __init__(self):
        root = tk.Tk()
        root.withdraw()

    def request_table_to_freeze(self):
        return filedialog.askopenfilename(
            title="Оберіть оригінальний Excel файл",
            filetypes=[("Excel files", "*.xlsx")]
        )

    def request_save_path(self, default_file_name: str):
        return filedialog.asksaveasfilename(
            title="Оберіть куди зберегти копію",
            initialfile=default_file_name,
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )

    def show_error_message(self, error: Exception):
        messagebox.showerror("Помилка", f"Сталася помилка при обробці: {error}")

    def show_success_message(self, message: str):
        messagebox.showinfo("Успіх", message=message)


