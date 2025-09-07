"""GUI generator for creating DOCX chapter files.

The application registers the bundled 'Cattedrale' font from the local fonts
directory so the font file must be available on disk.
"""

import ctypes
import os
import sys
import time
import re
import tkinter as tk
from tkinter import filedialog, ttk
from tkinter import font as tkfont

import copy
import customtkinter as ctk
from docx import Document
from typing import Dict, List, Tuple


def split_document(file_path: str) -> List[str]:
    """Split a DOCX document into chapters based on heading pattern.

    Parameters
    ----------
    file_path: str
        Path to the source DOCX file.

    Returns
    -------
    List[str]
        List of paths to the created chapter files.
    """

    heading_pattern = re.compile(r"^Глава\s+\d+(?:\.\d+)?")
    output_dir = filedialog.askdirectory(
        title="Выберите выходной каталог",
        initialdir=os.path.dirname(file_path),
    )
    if not output_dir:
        output_dir = os.path.dirname(file_path)
    document = Document(file_path)
    current_doc = None
    current_title = ""
    created_files: List[str] = []

    def _unique_path(directory: str, filename: str) -> str:
        base, ext = os.path.splitext(filename)
        candidate = os.path.join(directory, filename)
        counter = 2
        while os.path.exists(candidate):
            candidate = os.path.join(directory, f"{base} ({counter}){ext}")
            counter += 1
        return candidate

    for para in document.paragraphs:
        text = para.text.strip()
        if heading_pattern.match(text):
            if current_doc is not None:
                sanitized = re.sub(r"[^\w\s.-]", "", current_title).strip()
                if not sanitized:
                    sanitized = "section"
                file_name = f"{sanitized}.docx"
                out_path = _unique_path(output_dir, file_name)
                current_doc.save(out_path)
                created_files.append(out_path)
            current_doc = Document()
            current_title = text
        else:
            if current_doc is not None:
                new_element = copy.deepcopy(para._element)
                current_doc._element.body.append(new_element)

    if current_doc is not None:
        sanitized = re.sub(r"[^\w\s.-]", "", current_title).strip()
        if not sanitized:
            sanitized = "section"
        file_name = f"{sanitized}.docx"
        out_path = _unique_path(output_dir, file_name)
        current_doc.save(out_path)
        created_files.append(out_path)

    return created_files


def check_english_words(file_path: str) -> Dict[str, List[Tuple[int, int]]]:
    """Return mapping of English words to their paragraph and character positions."""

    document = Document(file_path)
    results: Dict[str, List[Tuple[int, int]]] = {}
    for p_idx, para in enumerate(document.paragraphs, start=1):
        for match in re.finditer(r"[A-Za-z]+", para.text):
            word = match.group()
            char_pos = match.start() + 1  # 1-indexed character position
            results.setdefault(word, []).append((p_idx, char_pos))
    return {word: positions for word, positions in sorted(results.items())}

# Path to store window geometry
CONFIG_PATH = os.path.join(os.path.dirname(__file__), "window.cfg")

class CustomInputDialog(ctk.CTkToplevel):
    """Simple dialog asking the user for a single line of text."""

    def __init__(self, master, question: str, font: ctk.CTkFont, icon_path: str):
        super().__init__(master)
        self.icon_path = icon_path
        self.iconbitmap(self.icon_path)
        self.title("")
        self.resizable(False, False)
        self.configure(fg_color="#2f2f2f")
        self.result = None

        self._label = ctk.CTkLabel(
            self, text=question, text_color="#eeeeee", font=font
        )
        self._label.pack(padx=20, pady=(20, 10))

        self._entry = ctk.CTkEntry(
            self,
            fg_color="#ffffff",
            border_color="#2f2f2f",
            text_color="#303030",
            corner_radius=12,
            border_width=0,
            font=font,
        )
        self._entry.pack(padx=20, pady=(0, 20))

        button_frame = ctk.CTkFrame(self, fg_color="#2f2f2f")
        button_frame.pack(padx=20, pady=(0, 20))

        self._ok_button = ctk.CTkButton(
            button_frame,
            text="OK",
            command=self._ok,
            fg_color="#313131",
            hover_color="#3e3e3e",
            text_color="#eeeeee",
            corner_radius=12,
            border_width=0,
            font=font,
        )
        self._ok_button.pack(side="left", padx=(0, 10))

        self._cancel_button = ctk.CTkButton(
            button_frame,
            text="Cancel",
            command=self._cancel,
            fg_color="#313131",
            hover_color="#3e3e3e",
            text_color="#eeeeee",
            corner_radius=12,
            border_width=0,
            font=font,
        )
        self._cancel_button.pack(side="left")

        self._entry.bind("<Return>", lambda event: self._ok())
        self.protocol("WM_DELETE_WINDOW", self._cancel)

        self.update_idletasks()
        master.update_idletasks()
        x = master.winfo_rootx() + (master.winfo_width() // 2) - (
            self.winfo_width() // 2
        )
        y = master.winfo_rooty() + (master.winfo_height() // 2) - (
            self.winfo_height() // 2
        )
        self.geometry(f"+{x}+{y}")

    def _ok(self) -> None:
        self.result = self._entry.get()
        self.destroy()

    def _cancel(self) -> None:
        self.result = None
        self.destroy()

    def get_input(self):
        self._entry.focus()
        self.grab_set()
        self.wait_window()
        return self.result

class Application(tk.Tk):
    def __init__(self):
        super().__init__()
        self.icon_path = os.path.join(
            os.path.dirname(__file__),
            "ChatGPT Image 15 авг. 2025 г., 20_42_09.ico",
        )
        self.iconbitmap(self.icon_path)

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
        self.config_data["font_size"] = str(font_size)
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
            family=custom_font.actual("family"), size=25, weight="bold"
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
        self.path_entry = ctk.CTkEntry(
            self.frame,
            placeholder_text="Куда сейвим?",
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

        # Button to split document into chapters
        self.split_button = ctk.CTkButton(
            self.frame,
            text="Разбить!",
            command=self.split_document,
            corner_radius=12,
            fg_color="#313131",
            hover_color="#3e3e3e",
            bg_color="#2f2f2f",
            text_color="#eeeeee",
            border_width=0,
            font=self.custom_font,
        )
        self.split_button.pack(pady=10)

        # Button to search for English words
        self.artifacts_button = ctk.CTkButton(
            self.frame,
            text="Проверка на артефакты",
            command=self.check_english_words,
            corner_radius=12,
            fg_color="#313131",
            hover_color="#3e3e3e",
            bg_color="#2f2f2f",
            text_color="#eeeeee",
            border_width=0,
            font=self.custom_font,
        )
        self.artifacts_button.pack(pady=10)


    def browse_folder(self):
        folder_selected = filedialog.askdirectory(title="Выберите папку для сохранения")
        if folder_selected:
            self.path_entry.delete(0, tk.END)
            self.path_entry.insert(0, folder_selected)

    def split_document(self):
        file_path = filedialog.askopenfilename(
            title="Выберите документ",
            filetypes=[("Word Documents", "*.docx")],
        )
        if not file_path:
            return

        created = split_document(file_path)
        if created:
            self.show_message(f"Создано {len(created)} файлов")

    def check_english_words(self):
        file_path = filedialog.askopenfilename(
            title="Выберите документ",
            filetypes=[("Word Documents", "*.docx")],
        )
        if not file_path:
            return

        words_with_pos = check_english_words(file_path)

        if not words_with_pos:
            self.show_popup("Английские слова не найдены.")
            return

        popup = ctk.CTkToplevel(self, fg_color="#2f2f2f")
        popup.iconbitmap(self.icon_path)
        popup.title("")
        popup.geometry("400x400")

        tree = ttk.Treeview(
            popup, columns=("word", "paragraph"), show="headings"
        )
        tree.heading("word", text="Слово")
        tree.heading("paragraph", text="№ параграфа")
        tree.column("word", anchor="w")
        tree.column("paragraph", anchor="center")

        for word, positions in words_with_pos.items():
            paragraphs = ", ".join(str(p) for p, _ in positions)
            tree.insert("", "end", values=(word, paragraphs))

        tree.pack(expand=True, fill="both", padx=10, pady=10)

        button_frame = ctk.CTkFrame(popup, fg_color="#2f2f2f")
        button_frame.pack(pady=(0, 10))

        if len(words_with_pos) > 50:
            save_button = ctk.CTkButton(
                button_frame,
                text="Сохранить",
                command=lambda: self.save_words_to_file(words_with_pos),
                corner_radius=12,
                bg_color="#2f2f2f",
                fg_color="#313131",
                hover_color="#3e3e3e",
                text_color="#eeeeee",
                border_width=0,
                font=self.custom_font,
            )
            save_button.pack(side="left", padx=(0, 10))

        close_button = ctk.CTkButton(
            button_frame,
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
        close_button.pack(side="left")

    def save_words_to_file(self, words):
        folder = filedialog.askdirectory(title="Выберите папку для сохранения списка")
        if not folder:
            return
        file_path = os.path.join(folder, "english_words.txt")
        lines = [
            f"{word} - {', '.join(str(p) for p, _ in positions)}"
            for word, positions in words.items()
        ]
        with open(file_path, "w", encoding="utf-8") as f:
            f.write("\n".join(lines))
        self.show_popup(f"Список сохранен в {file_path}")

    def ask_questions(self):
        total_dialog = CustomInputDialog(
            self, "Сколько ебануть?", self.custom_font, self.icon_path
        )

        total_chapters = total_dialog.get_input()
        if total_chapters is None:
            return
        total_chapters = int(total_chapters)

        parts_dialog = CustomInputDialog(
            self,
            "На сколько частей делим?",
            self.custom_font,
            self.icon_path,
        )

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
        self.show_popup(message, color="#ff0000")

    def on_closing(self):
        self.config_data["geometry"] = self.geometry()
        self.config_data["font_size"] = str(self.custom_font.cget("size"))
        self.save_config()
        self.destroy()

    def show_popup(self, message, color="#00ff00"):
        popup = ctk.CTkToplevel(self, fg_color="#2f2f2f")
        popup.iconbitmap(self.icon_path)
        popup.title("")
        popup.geometry("300x100")

        frame = ctk.CTkFrame(popup, corner_radius=12, fg_color="#2f2f2f")
        frame.pack(fill="both", expand=True)

        label = ctk.CTkLabel(
            frame,
            text=message,
            text_color=color,
            font=ctk.CTkFont(
                family=self.custom_font.actual("family"),
                size=self.custom_font.cget("size"),
                weight="bold",
            ),
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
