import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext, messagebox
import threading
from docx_image_parser import process_document, setup_tesseract, SUPPORTED_LANGS
import os
import webbrowser
from pathlib import Path

class DocumentProcessorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Document OCR")
        self.root.geometry("800x600")
        self.root.resizable(False, False)
        self.selected_lang = tk.StringVar(value='Русский')
        self.file_path = None
        self.results = None
        self.doc_text = ''
        self.output_folder = None
        
        # Создаем папку для временных файлов в AppData
        app_data = os.getenv('APPDATA')
        self.app_folder = os.path.join(app_data, 'DocumentOCR')
        self.output_folder = os.path.join(self.app_folder, 'temp')
        os.makedirs(self.output_folder, exist_ok=True)

        self.setup_styles()

        main_frame = ttk.Frame(root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        header_frame = ttk.Frame(main_frame)
        header_frame.pack(fill=tk.X, pady=(0, 20))
        ttk.Label(
            header_frame,
            text="Document OCR",
            font=("Segoe UI", 24, "bold"),
            style="Header.TLabel"
        ).pack(side=tk.LEFT)
        ttk.Button(
            header_frame,
            text="?",
            width=3,
            command=self.show_help,
            style="Accent.TButton"
        ).pack(side=tk.RIGHT)

        file_frame = ttk.Frame(main_frame)
        file_frame.pack(fill=tk.X, pady=10)
        ttk.Button(
            file_frame,
            text="Открыть файл",
            command=self.select_file,
            style="Accent.TButton",
            width=15
        ).pack(side=tk.LEFT)
        self.file_label = ttk.Label(
            file_frame,
            text="Файл не выбран",
            style="Info.TLabel"
        )
        self.file_label.pack(side=tk.LEFT, padx=15)

        output_frame = ttk.Frame(main_frame)
        output_frame.pack(fill=tk.X, pady=10)
        ttk.Button(
            output_frame,
            text="Выбрать папку",
            command=self.select_output_folder,
            style="Accent.TButton",
            width=15
        ).pack(side=tk.LEFT)
        self.output_label = ttk.Label(
            output_frame,
            text="Папка не выбрана",
            style="Info.TLabel"
        )
        self.output_label.pack(side=tk.LEFT, padx=15)

        lang_frame = ttk.Frame(main_frame)
        lang_frame.pack(fill=tk.X, pady=10)
        ttk.Label(
            lang_frame,
            text="Язык OCR:",
            style="Info.TLabel",
            width=10
        ).pack(side=tk.LEFT)
        self.lang_combo = ttk.Combobox(
            lang_frame,
            values=list(SUPPORTED_LANGS.keys()),
            state="readonly",
            width=20,
            textvariable=self.selected_lang,
            style="Accent.TCombobox"
        )
        self.lang_combo.pack(side=tk.LEFT)

        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=10)
        self.process_btn = ttk.Button(
            btn_frame,
            text="Распознать",
            command=self.process_document,
            state=tk.DISABLED,
            style="Accent.TButton",
            width=15
        )
        self.process_btn.pack(side=tk.LEFT)
        self.save_btn = ttk.Button(
            btn_frame,
            text="Сохранить TXT",
            command=self.save_results,
            state=tk.DISABLED,
            style="Accent.TButton",
            width=15
        )
        self.save_btn.pack(side=tk.LEFT, padx=10)

        self.progress = ttk.Progressbar(
            main_frame,
            orient=tk.HORIZONTAL,
            length=760,
            mode='determinate',
            style="Accent.Horizontal.TProgressbar"
        )
        self.progress.pack(fill=tk.X, pady=10)

        text_frame = ttk.Frame(main_frame)
        text_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        self.output_text = scrolledtext.ScrolledText(
            text_frame,
            wrap=tk.WORD,
            width=90,
            height=20,
            font=("Segoe UI", 10),
            bg="#2b2b2b",
            fg="#ffffff",
            insertbackground="#ffffff"
        )
        self.output_text.pack(fill=tk.BOTH, expand=True)

        self.status_label = ttk.Label(
            main_frame,
            text="Готов к работе",
            style="Info.TLabel"
        )
        self.status_label.pack(pady=10)

        self.check_tesseract()

    def setup_styles(self):
        style = ttk.Style()
        style.theme_use('clam')

        bg_color = "#1e1e1e"
        accent_color = "#0078d4"
        text_color = "#ffffff"
        secondary_bg = "#2b2b2b"
        border_color = "#3c3c3c"

        style.configure(".",
            background=bg_color,
            foreground=text_color,
            font=("Segoe UI", 10)
        )

        style.configure("TButton",
            padding=6,
            relief="flat",
            background=secondary_bg,
            foreground=text_color,
            font=("Segoe UI", 10, "bold")
        )
        style.configure("Accent.TButton",
            background=accent_color,
            foreground="white",
            padding=6
        )
        style.map("Accent.TButton",
            background=[("active", "#106ebe"), ("disabled", "#404040")]
        )

        style.configure("TLabel",
            background=bg_color,
            foreground=text_color,
            padding=2,
            font=("Segoe UI", 10)
        )
        style.configure("Info.TLabel",
            font=("Segoe UI", 9),
            foreground="#888888"
        )
        style.configure("Header.TLabel",
            font=("Segoe UI", 24, "bold"),
            foreground=text_color
        )

        style.configure("TCombobox",
            padding=6,
            selectbackground=accent_color,
            fieldbackground=secondary_bg,
            background=secondary_bg,
            foreground=text_color,
            arrowcolor=text_color,
            font=("Segoe UI", 10)
        )
        style.configure("Accent.TCombobox",
            fieldbackground=secondary_bg,
            background=secondary_bg,
            foreground=text_color,
            arrowcolor=text_color
        )
        style.map("TCombobox",
            fieldbackground=[("readonly", secondary_bg)],
            selectbackground=[("readonly", accent_color)],
            selectforeground=[("readonly", "white")]
        )

        style.configure("Accent.Horizontal.TProgressbar",
            troughcolor=secondary_bg,
            background=accent_color,
            thickness=8,
            borderwidth=0
        )

        style.configure("TFrame",
            background=bg_color
        )

        self.root.configure(bg=bg_color)
        self.root.option_add("*TCombobox*Listbox.background", secondary_bg)
        self.root.option_add("*TCombobox*Listbox.foreground", text_color)
        self.root.option_add("*TCombobox*Listbox.selectBackground", accent_color)
        self.root.option_add("*TCombobox*Listbox.selectForeground", "white")

    def show_help(self):
        messagebox.showinfo(
            "Справка",
            "1. Выберите DOCX или PDF файл.\n"
            "2. Выберите язык OCR.\n"
            "3. Нажмите 'Распознать'.\n"
            "4. Просмотрите результаты.\n"
            "5. Сохраните результат в TXT.\n\n"
            "Текст извлекается как из документа, так и из изображений."
        )

    def check_tesseract(self):
        try:
            setup_tesseract()
            self.update_status("Tesseract готов к работе")
        except Exception as e:
            self.update_status("Tesseract не найден")
            if messagebox.askyesno(
                "Tesseract не установлен",
                "Для работы приложения требуется Tesseract OCR.\n\n"
                "Хотите открыть страницу загрузки Tesseract?"
            ):
                webbrowser.open("https://github.com/UB-Mannheim/tesseract/wiki")
            self.process_btn.config(state=tk.DISABLED)

    def select_output_folder(self):
        folder = filedialog.askdirectory(title="Выберите папку для сохранения результатов")
        if folder:
            self.output_folder = folder
            self.output_label.config(text=os.path.basename(folder))
            if self.file_path:
                self.process_btn.config(state=tk.NORMAL)

    def select_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[
                ("Документы", "*.docx;*.pdf"),
                ("Word documents", "*.docx"),
                ("PDF files", "*.pdf")
            ]
        )
        if file_path:
            self.file_path = file_path
            self.file_label.config(text=os.path.basename(file_path))
            if self.output_folder:
                self.process_btn.config(state=tk.NORMAL)
            self.save_btn.config(state=tk.DISABLED)
            self.output_text.delete(1.0, tk.END)
            self.status_label.config(text="Файл выбран. Готов к распознаванию.")

    def process_document(self):
        if not self.file_path or not self.output_folder:
            return
        self.process_btn.config(state=tk.DISABLED)
        self.save_btn.config(state=tk.DISABLED)
        self.progress['value'] = 0
        self.output_text.delete(1.0, tk.END)
        self.update_status("Обработка документа...")
        lang = SUPPORTED_LANGS[self.selected_lang.get()]
        
        def process():
            try:
                doc_text, results, results_file = process_document(self.file_path, self.output_folder, lang=lang)
                self.doc_text = doc_text
                self.results = results
                self.show_results(doc_text, results)
                self.save_btn.config(state=tk.NORMAL)
                self.update_status("Готово!")
            except Exception as e:
                self.update_status(f"Ошибка: {str(e)}")
            finally:
                self.process_btn.config(state=tk.NORMAL)
                self.progress['value'] = 100
        
        threading.Thread(target=process, daemon=True).start()
        self.progress['value'] = 30

    def show_results(self, doc_text, results):
        self.output_text.delete(1.0, tk.END)
        self.output_text.insert(tk.END, "Текст из документа:\n\n")
        self.output_text.insert(tk.END, doc_text + '\n\n')
        self.output_text.insert(tk.END, "Текст из изображений:\n\n")
        for i, img in enumerate(results, 1):
            self.output_text.insert(tk.END, f"Изображение {i}: {img['filename']}\n")
            self.output_text.insert(tk.END, f"Уверенность: {img['confidence']:.2f}%\n")
            self.output_text.insert(tk.END, "Текст:\n")
            self.output_text.insert(tk.END, img['text'] + '\n')
            self.output_text.insert(tk.END, "-" * 40 + "\n")

    def save_results(self):
        if not self.output_folder:
            return
        save_path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt")]
        )
        if save_path:
            try:
                with open(os.path.join(self.output_folder, 'extracted_text.txt'), 'r', encoding='utf-8') as src:
                    with open(save_path, 'w', encoding='utf-8') as dst:
                        dst.write(src.read())
                messagebox.showinfo("Сохранено", f"Результат сохранён в {save_path}")
            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось сохранить файл: {str(e)}")

    def update_status(self, message):
        self.status_label.config(text=message)
        self.root.update_idletasks()

def main():
    root = tk.Tk()
    app = DocumentProcessorGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main() 