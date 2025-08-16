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
from tkinter import filedialog, ttk
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
        self.icon_path = icon_path
        self.iconphoto(True, tk.PhotoImage(file=icon_path))

        self.config_data = self.load_config()

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
        base_size = default_font.cget("size") + 4
        font_size = int(self.config_data.get("font_size", base_size))
        custom_font = ctk.CTkFont(family=font_family, size=font_size)
        self.custom_font = custom_font
        self.option_add("*Font", custom_font)

        self.style = ttk.Style(self)
        self.style.theme_use("clam")

        self.style.configure("Custom.TFrame", background="#2f2f2f")
        self.style.configure("Custom.TLabel", background="#2f2f2f", foreground="#eeeeee")

        # Создаем окно перед настройкой шрифта
        self.title("")
        # Set window geometry
        saved_geom = self.config_data.get("geometry", "")
        if saved_geom:
            self.geometry(saved_geom)
        else:
            self.geometry("500x400")  # Размер окна
        self.configure(bg="#2f2f2f")  # Темно-серый фон
        self.resizable(True, True)
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

        # Создаем рамку для текста
        self.frame = ctk.CTkFrame(self, fg_color="#2f2f2f")
        self.frame.pack(padx=20, pady=20, expand=True, fill="both")

        # Создаем метку
        header_font = ctk.CTkFont(
            family=custom_font.actual("family"), size=20, weight="bold"
        )
        self.label = ttk.Label(self.frame, text="Генератор Глав", font=header_font, style="Custom.TLabel")
        self.label.pack(pady=20)

        # Кнопка для начала генерации
        self.ask_button = ctk.CTkButton(
            self.frame,
            text="Начать генерацию",
            command=self.ask_questions,
            corner_radius=12,
            fg_color="#313131",
            hover_color="#3e3e3e",
            bg_color="#2f2f2f",
            text_color="#eeeeee",
            border_width=0,
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
            fg_color="#313131",
            hover_color="#3e3e3e",
            bg_color="#2f2f2f",
            text_color="#eeeeee",
            border_width=0,
            font=self.custom_font,
        )
        self.browse_button.pack(pady=10)

        # Button to adjust font size
        self.font_button = ctk.CTkButton(
            self,
            text="A",
            width=24,
            height=24,
            command=self.show_font_settings,
            corner_radius=12,
            fg_color="#313131",
            hover_color="#3e3e3e",
            bg_color="#2f2f2f",
            text_color="#eeeeee",
            border_width=0,
            font=self.custom_font,
        )
        self.font_button.place(relx=1.0, rely=0.0, x=-10, y=10, anchor="ne")

    def browse_folder(self):
        folder_selected = filedialog.askdirectory(title="Выберите папку для сохранения")
        if folder_selected:
            self.path_entry.delete(0, tk.END)
            self.path_entry.insert(0, folder_selected)

    def ask_questions(self):
        def style_dialog(d):
            if hasattr(d, "_ok_button"):
                for btn in (d._ok_button, d._cancel_button):
                    btn.configure(
                        bg_color="#2f2f2f",
                        fg_color="#313131",
                        hover_color="#3e3e3e",
                        corner_radius=12,
                        text_color="#eeeeee",
                        border_width=0,
                    )
                if hasattr(d, "_entry"):
                    d._entry.configure(
                        bg_color="#2f2f2f",
                        fg_color="#ffffff",
                        border_color="#2f2f2f",
                        corner_radius=12,
                        text_color="#303030",
                        border_width=0,
                    )
            else:
                d.after(20, lambda: style_dialog(d))

        total_dialog = ctk.CTkInputDialog(
            title="Сколько ебануть?",
            text="Введите количество глав:",
            fg_color="#2f2f2f",
            text_color="#eeeeee",
            button_fg_color="#313131",
            button_hover_color="#3e3e3e",
            entry_fg_color="#ffffff",
            entry_border_color="#2f2f2f",
            entry_text_color="#303030",
            font=self.custom_font,
        )
        total_dialog._window.wm_iconphoto(True, tk.PhotoImage(file=self.icon_path))
        style_dialog(total_dialog)

        total_chapters = total_dialog.get_input()
        if total_chapters is None:
            return
        total_chapters = int(total_chapters)

        parts_dialog = ctk.CTkInputDialog(
            title="На сколько делим епт?",
            text="Введите количество частей в главе:",
            fg_color="#2f2f2f",
            text_color="#eeeeee",
            button_fg_color="#313131",
            button_hover_color="#3e3e3e",
            entry_fg_color="#ffffff",
            entry_border_color="#2f2f2f",
            entry_text_color="#303030",
            font=self.custom_font,
        )
        parts_dialog._window.wm_iconphoto(True, tk.PhotoImage(file=self.icon_path))
        style_dialog(parts_dialog)

        parts_per_chapter = parts_dialog.get_input()
        if parts_per_chapter is None:
            return
        parts_per_chapter = int(parts_per_chapter)

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
        self.config_data["geometry"] = self.geometry()
        self.config_data.setdefault("font_size", str(self.custom_font.cget("size")))
        self.save_config()
        self.destroy()

    def show_popup(self, message, color="green"):
        popup = ctk.CTkToplevel(self, fg_color="#2f2f2f")
        popup.geometry("300x100")
        popup.title("Результат")

        frame = ctk.CTkFrame(popup, corner_radius=12, fg_color="#2f2f2f")
        frame.pack(fill="both", expand=True)

        label = ctk.CTkLabel(
            frame, text=message, font=self.custom_font, text_color=color
        )
        label.pack(pady=20)

        close_button = ctk.CTkButton(
            frame,
            text="Закрыть",
            command=popup.destroy,
            corner_radius=12,
            bg_color="#2f2f2f",
            fg_color="#313131",
            hover_color="#3e3e3e",
            text_color="#eeeeee",
            border_width=0,
            font=self.custom_font,
        )
        close_button.pack(pady=5)

    def show_font_settings(self):
        top = ctk.CTkToplevel(self, fg_color="#2f2f2f")
        top.title("Font size")

        frame = ctk.CTkFrame(top, corner_radius=12, fg_color="#2f2f2f")
        frame.pack(fill="both", expand=True)

        slider = ctk.CTkSlider(frame, from_=8, to=32)
        slider.set(self.custom_font.cget("size"))
        slider.pack(padx=20, pady=20)

        def apply():
            size = int(slider.get())
            self.custom_font.configure(size=size)
            self.option_add("*Font", self.custom_font)
            self.config_data["font_size"] = str(size)
            self.save_config()
            top.destroy()

        apply_button = ctk.CTkButton(
            frame,
            text="Apply",
            command=apply,
            corner_radius=12,
            fg_color="#313131",
            hover_color="#3e3e3e",
            bg_color="#2f2f2f",
            text_color="#eeeeee",
            border_width=0,
        )
        apply_button.pack(pady=10)

    def load_config(self):
        config = {}
        if os.path.exists(CONFIG_PATH):
            with open(CONFIG_PATH) as f:
                for line in f:
                    line = line.strip()
                    if not line:
                        continue
                    if "=" in line:
                        key, value = line.split("=", 1)
                        config[key] = value
                    else:
                        config["geometry"] = line
        return config

    def save_config(self):
        with open(CONFIG_PATH, "w") as f:
            for key, value in self.config_data.items():
                f.write(f"{key}={value}\n")

if __name__ == "__main__":
    app = Application()
    app.mainloop()
