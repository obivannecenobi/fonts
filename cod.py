import tkinter as tk
from tkinter import simpledialog, filedialog
from tkinter import ttk, font as tkfont
import os
import time
from docx import Document

class Application(tk.Tk):
    def __init__(self):
        super().__init__()

        # Создаем окно перед настройкой шрифта
        self.title("Генератор Глав")
        self.geometry("500x400")  # Размер окна
        self.configure(bg="#2f2f2f")  # Темно-серый фон
        self.resizable(False, False)  # Запрещаем изменение размера окна

        # Путь к вашему шрифту
        font_path = os.path.join(
            os.path.dirname(__file__),
            "fonts",
            "Cattedrale[RUSbypenka220]-Regular.ttf",
        )

        # Настройка шрифта
        default_size = tkfont.nametofont("TkDefaultFont").cget("size")
        custom_font = tkfont.Font(file=font_path, family="Cattedrale", size=default_size)
        self.option_add("*Font", custom_font)

        # Стиль для виджетов
        self.style = ttk.Style(self)
        self.style.configure("TButton",
                             background="#eeeeee",  # Светло-серые кнопки
                             foreground="#313131",  # Темно-серый текст в кнопках
                             relief="flat",  # Без границ
                             padding=12)

        self.style.configure("TEntry",
                             foreground="#303030",  # Темно-серый текст в поле ввода
                             background="#ffffff",  # Белое поле ввода
                             font=("Arial", default_size),  # Шрифт по умолчанию для поля ввода
                             fieldbackground="#ffffff")

        self.style.map("TButton", background=[("active", "#2AD1A3")])  # Цвет кнопки при наведении (неон бирюзово-зеленый)

        # Скругление кнопок и поля ввода
        self.style.configure("TButton",
                             borderwidth=5,
                             relief="solid",
                             width=20,
                             padding=(10, 5))  # Сильно скругленные углы кнопок

        self.style.configure("TEntry",
                             relief="solid",
                             borderwidth=3,
                             width=30,
                             padding=(10, 5))  # Поле ввода с четким контуром и скругленными углами

        # Создаем рамку для текста
        self.frame = tk.Frame(self, bg="#2f2f2f")
        self.frame.pack(padx=20, pady=20, expand=True, fill="both")

        # Создаем метку
        self.label = tk.Label(self.frame, text="Генератор Глав", fg="#eeeeee", bg="#2f2f2f", font=("Cattedrale", 16, "bold"))
        self.label.pack(pady=20)

        # Кнопки для взаимодействия
        self.ask_button = ttk.Button(self.frame, text="Начать генерацию", command=self.ask_questions, style="TButton")
        self.ask_button.pack(pady=10)

        # Поле для ввода пути с кнопкой
        self.path_label = tk.Label(self.frame, text="Выберите путь для сохранения:", fg="#eeeeee", bg="#2f2f2f", font=("Arial", 12))
        self.path_label.pack(pady=10)

        self.path_entry = ttk.Entry(self.frame)
        self.path_entry.pack(fill=tk.X, padx=10, pady=5)

        # Кнопка для выбора папки
        self.browse_button = ttk.Button(self.frame, text="Выбрать папку", command=self.browse_folder, style="TButton")
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

    def show_popup(self, message, color="green"):
        popup = tk.Toplevel(self)
        popup.geometry("300x100")
        popup.title("Результат")
        popup.configure(bg="#2f2f2f")

        label = tk.Label(popup, text=message, fg=color, bg="#2f2f2f", font=("Arial", 12))
        label.pack(pady=20)

        close_button = ttk.Button(popup, text="Закрыть", command=popup.destroy, style="TButton")
        close_button.pack()

if __name__ == "__main__":
    app = Application()
    app.mainloop()
