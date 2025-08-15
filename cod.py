"""GUI generator for creating DOCX chapter files.

The application registers the bundled 'Cattedrale' font from the local fonts
directory so the font file must be available on disk.
"""

import base64
import ctypes
import os
import sys
import time
import tkinter as tk
from tkinter import filedialog, simpledialog, ttk
from tkinter import font as tkfont

import customtkinter as ctk
from docx import Document

# Path to store window geometry
CONFIG_PATH = os.path.join(os.path.dirname(__file__), "window.cfg")

ICON_PNG_B64 = (
    "iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAAr0lEQVR4nO3QLW5CQRAA4C9F9AYVCBxp0jOAqCAQwgEqMfQInILgcEU0XABFguIW3ACFwlY0NSNe1rS8"
    "lCfIfGp2fnYnS0oppTuxxkOdwVpDhUd84bvJBXqV+BWHiPu3XqCLTwwquSH2EY+wib5/t8AST0V+XZzbWEXvr675gQ5OuFRyLzgWfefIta+4+8/G2GEa5zmeK/U3bKPv"
    "Zlp4j/ijqM2i3phJk4+llFK6Tz/HBhQcv0+QOQAAAABJRU5ErkJggg=="
)


def ensure_icon(path: str) -> None:
    """Create the icon file from embedded data if it is missing."""
    if not os.path.exists(path):
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, "wb") as f:
            f.write(base64.b64decode(ICON_PNG_B64))

class Application(tk.Tk):
    def __init__(self):
        super().__init__()
        icon_path = os.path.join(os.path.dirname(__file__), "icons", "code.png")
        ensure_icon(icon_path)
        self.iconphoto(True, tk.PhotoImage(file=icon_path))

        # Путь к вашему шрифту
        font_path = os.path.join(
            os.path.dirname(__file__),
            "fonts",
            "Cattedrale[RUSbypenka220]-Regular.ttf",
        )

        # Регистрация и настройка шрифта
        if sys.platform.startswith("win"):
            ctypes.windll.gdi32.AddFontResourceExW(font_path, 0x10, 0)
            font_family = "Cattedrale [RUS by penka220]"
        else:
            # На системах Unix пытаемся зарегистрировать шрифт через tkfont.
            # Если это невозможно, используем системный шрифт по умолчанию.
            try:
                self._loaded_font = tkfont.Font(file=font_path)
                font_family = self._loaded_font.actual("family")
            except tk.TclError:
                font_family = tkfont.nametofont("TkDefaultFont").actual("family")
        default_font = tkfont.nametofont("TkDefaultFont")
        custom_font = ctk.CTkFont(
            family=font_family,
            size=default_font.cget("size") + 4,
        )
        self.custom_font = custom_font
        self.option_add("*Font", custom_font)

        self.style = ttk.Style(self)
        self.style.theme_use("clam")

        self.style.configure("Custom.TFrame", background="#2f2f2f")
        self.style.configure("Custom.TLabel", background="#2f2f2f", foreground="#eeeeee")

        # Создаем окно перед настройкой шрифта
        self.title("")
        # Set window geometry
        saved_geom = ""
        if os.path.exists(CONFIG_PATH):
            with open(CONFIG_PATH) as f:
                saved_geom = f.read().strip()
        if saved_geom:
            self.geometry(saved_geom)
        else:
            self.geometry("500x400")  # Размер окна
        self.configure(bg="#2f2f2f")  # Темно-серый фон
        self.resizable(True, True)
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

        # Создаем рамку для текста
        self.frame = ttk.Frame(self, style="Custom.TFrame")
        self.frame.pack(padx=20, pady=20, expand=True, fill="both")

        # Создаем метку
        header_font = ctk.CTkFont(family=custom_font.actual("family"), size=18, weight="bold")
        self.label = ttk.Label(self.frame, text="Генератор Глав", font=header_font, style="Custom.TLabel")
        self.label.pack(pady=20)

        # Кнопка для начала генерации
        self.ask_button = ctk.CTkButton(
            self.frame,
            text="Начать генерацию",
            command=self.ask_questions,
            corner_radius=12,
            border_width=0,
            fg_color="#313131",
            hover_color="#3e3e3e",
            font=self.custom_font,
        )
        self.ask_button.pack(pady=10)

        # Поле для ввода пути с кнопкой
        self.path_label = ttk.Label(
            self.frame, text="Выберите путь для сохранения:", style="Custom.TLabel"
        )
        self.path_label.pack(pady=10)

        self.path_entry = ctk.CTkEntry(
            self.frame,
            corner_radius=12,
            fg_color="#ffffff",
            text_color="#303030",
            border_color="#2f2f2f",
            border_width=0,
            bg_color="#2f2f2f",
            font=self.custom_font,
        )
        self.path_entry.pack(fill=tk.X, padx=10, pady=5)

        # Кнопка для выбора папки
        self.browse_button = ctk.CTkButton(
            self.frame,
            text="Выбрать папку",
            command=self.browse_folder,
            corner_radius=12,
            border_width=0,
            fg_color="#313131",
            hover_color="#3e3e3e",
            font=self.custom_font,
        )
        self.browse_button.pack(pady=10)

    def browse_folder(self):
        folder_selected = filedialog.askdirectory(title="Выберите папку для сохранения")
        if folder_selected:
            self.path_entry.delete(0, tk.END)
            self.path_entry.insert(0, folder_selected)

    def ask_questions(self):
        total_chapters = simpledialog.askinteger("Сколько ебануть?", "Введите количество глав:", parent=self, minvalue=1, maxvalue=100)
        parts_per_chapter = simpledialog.askinteger("На сколько делим епт?", "Введите количество частей в главе:", parent=self, minvalue=1, maxvalue=10)

        save_location = self.path_entry.get()

        if save_location:
            self.generate_files(save_location, total_chapters, parts_per_chapter)
        else:
            self.show_error("Не выбрана папка для сохранения.")

    def generate_files(self, save_location, total_chapters, parts_per_chapter):
        if not os.path.exists(save_location):
            os.makedirs(save_location)

        # Создание подпапки для файлов
        timestamp = time.strftime("%Y-%m-%d_%H-%M-%S")
        folder_for_chapters = os.path.join(save_location, f"Генерация_{timestamp}")
        os.makedirs(folder_for_chapters)

        for chapter in range(1, total_chapters + 1):
            for part in range(1, parts_per_chapter + 1):
                file_name = f"Глава {chapter}.{part}.docx"
                file_path = os.path.join(folder_for_chapters, file_name)

                doc = Document()
                # Пустой документ с минимальным содержимым (например, один пробел)
                doc.add_paragraph(" ")  # Добавляем один пробел
                doc.save(file_path)

        self.show_message(f"Создано {total_chapters * parts_per_chapter} файлов в папке {folder_for_chapters}")

    def show_message(self, message):
        self.show_popup(message)

    def show_error(self, message):
        self.show_popup(message, "red")

    def on_closing(self):
        with open(CONFIG_PATH, "w") as f:
            f.write(self.geometry())
        self.destroy()

    def show_popup(self, message, color="green"):
        popup = tk.Toplevel(self)
        popup.geometry("300x100")
        popup.title("Результат")
        popup.configure(bg="#2f2f2f")

        style = ttk.Style(popup)
        style.configure("Popup.TLabel", background="#2f2f2f", foreground=color)

        label = ttk.Label(popup, text=message, style="Popup.TLabel")
        label.pack(pady=20)

        close_frame = ctk.CTkFrame(
            popup,
            fg_color="#FFFFFF",
            width=200,
            height=50,
            corner_radius=12,
        )
        close_frame.pack(pady=5)
        close_frame.pack_propagate(False)
        close_glow = ctk.CTkLabel(
            close_frame,
            text="",
            fg_color="#E0E0E0",
            corner_radius=12,
        )
        close_glow.place(relx=0.5, rely=0.5, anchor="center", relwidth=1, relheight=1)
        close_button = ctk.CTkButton(
            close_frame,
            text="Закрыть",
            command=popup.destroy,
            corner_radius=12,
            fg_color="transparent",
            text_color="#313131",
            hover_color="#FFFFFF",
            border_color="white",
            border_width=2,
            font=self.custom_font,
        )
        close_button.place(relx=0.5, rely=0.5, anchor="center")

if __name__ == "__main__":
    app = Application()
    app.mainloop()
