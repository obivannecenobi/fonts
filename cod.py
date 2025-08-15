import tkinter as tk
from tkinter import simpledialog, filedialog, ttk
from tkinter import font as tkfont
import ctypes
import os
import time
from docx import Document

# Path to store window size
CONFIG_PATH = os.path.join(os.path.dirname(__file__), "window_size.txt")

class Application(tk.Tk):
    def __init__(self):
        super().__init__()

        # Путь к вашему шрифту
        font_path = os.path.join(
            os.path.dirname(__file__),
            "fonts",
            "Cattedrale[RUSbypenka220]-Regular.ttf",
        )

        # Регистрация и настройка шрифта
        ctypes.windll.gdi32.AddFontResourceExW(font_path, 0x10, 0)
        default_font = tkfont.nametofont("TkDefaultFont")
        custom_font = tkfont.Font(
            family="Cattedrale [RUS by penka220]",
            size=default_font.cget("size") + 2,
        )
        self.custom_font = custom_font
        self.option_add("*Font", custom_font)

        self.style = ttk.Style(self)
        self.style.theme_use("clam")

        self.style.configure("Custom.TFrame", background="#2f2f2f")
        self.style.configure("Custom.TLabel", background="#2f2f2f", foreground="#eeeeee")
        self.style.configure(
            "TButton",
            background="#eeeeee",
            foreground="#313131",
            borderwidth=2,
            padding=6,
            relief="flat",
            font=self.custom_font,
        )
        self.style.map("TButton", background=[("active", "#ffffff")])
        self.style.configure(
            "Custom.TEntry",
            foreground="#303030",
            fieldbackground="#ffffff",
            borderwidth=2,
        )

        # Создаем окно перед настройкой шрифта
        self.title("Генератор Глав")
        # Set window geometry
        if os.path.exists(CONFIG_PATH):
            with open(CONFIG_PATH) as f:
                size = f.read().strip()
            if size:
                self.geometry(size)
            else:
                self.geometry("500x400")  # Размер окна
        else:
            self.geometry("500x400")  # Размер окна
        self.configure(bg="#2f2f2f")  # Темно-серый фон
        self.resizable(True, True)
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

        # Создаем рамку для текста
        self.frame = ttk.Frame(self, style="Custom.TFrame")
        self.frame.pack(padx=20, pady=20, expand=True, fill="both")

        # Создаем метку
        header_font = tkfont.Font(family=custom_font.actual("family"), size=16, weight="bold")
        self.label = ttk.Label(self.frame, text="Генератор Глав", font=header_font, style="Custom.TLabel")
        self.label.pack(pady=20)

        # Кнопки для взаимодействия
        self.ask_button = ttk.Button(
            self.frame,
            text="Начать генерацию",
            command=self.ask_questions,
            style="TButton",
        )
        self.ask_button.pack(pady=10)

        # Поле для ввода пути с кнопкой
        self.path_label = ttk.Label(
            self.frame, text="Выберите путь для сохранения:", style="Custom.TLabel"
        )
        self.path_label.pack(pady=10)

        self.path_entry = ttk.Entry(self.frame, style="Custom.TEntry")
        self.path_entry.pack(fill=tk.X, padx=10, pady=5)

        # Кнопка для выбора папки
        self.browse_button = ttk.Button(
            self.frame,
            text="Выбрать папку",
            command=self.browse_folder,
            style="TButton",
        )
        self.browse_button.pack(pady=5)

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
        style.configure("TButton", font=self.custom_font)
        style.configure("Popup.TLabel", background="#2f2f2f", foreground=color)

        label = ttk.Label(popup, text=message, style="Popup.TLabel")
        label.pack(pady=20)

        close_button = ttk.Button(
            popup,
            text="Закрыть",
            command=popup.destroy,
            style="TButton",
        )
        close_button.pack()

if __name__ == "__main__":
    app = Application()
    app.mainloop()
