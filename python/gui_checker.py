"""
Графический интерфейс для детектора плагиата Word файлов
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from pathlib import Path
import threading
from word_plagiarism_checker import WordPlagiarismChecker


class PlagiarismCheckerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Детектор плагиата для Word файлов")
        self.root.geometry("900x700")
        self.root.resizable(True, True)
        
        # Переменные
        self.docs_directory = tk.StringVar(value=str(Path(__file__).parent.parent / "docs"))
        self.target_file = tk.StringVar()
        self.checker = None
        self.results = []
        
        self.create_widgets()
        
    def create_widgets(self):
        # Главный контейнер
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Настройка весов для растягивания
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Заголовок
        title_label = ttk.Label(
            main_frame, 
            text="Детектор плагиата для Word файлов", 
            font=("Arial", 16, "bold")
        )
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # Секция выбора директории
        dir_frame = ttk.LabelFrame(main_frame, text="Директория с документами", padding="10")
        dir_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        dir_frame.columnconfigure(1, weight=1)
        
        ttk.Label(dir_frame, text="Путь к директории:").grid(row=0, column=0, sticky=tk.W, padx=(0, 5))
        dir_entry = ttk.Entry(dir_frame, textvariable=self.docs_directory, width=50)
        dir_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 5))
        
        ttk.Button(
            dir_frame, 
            text="Выбрать...", 
            command=self.select_directory
        ).grid(row=0, column=2)
        
        # Секция выбора целевого файла
        file_frame = ttk.LabelFrame(main_frame, text="Целевой файл для сравнения", padding="10")
        file_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        file_frame.columnconfigure(1, weight=1)
        
        ttk.Label(file_frame, text="Имя файла:").grid(row=0, column=0, sticky=tk.W, padx=(0, 5))
        
        file_combo = ttk.Combobox(file_frame, textvariable=self.target_file, width=47, state="readonly")
        file_combo.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 5))
        self.file_combo = file_combo
        
        ttk.Button(
            file_frame, 
            text="Обновить список", 
            command=self.refresh_file_list
        ).grid(row=0, column=2)
        
        ttk.Label(
            file_frame, 
            text="(Оставьте пустым для использования первого файла)", 
            font=("Arial", 8)
        ).grid(row=1, column=0, columnspan=3, pady=(5, 0))
        
        # Кнопки управления
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=3, column=0, columnspan=3, pady=(0, 10))
        
        self.start_button = ttk.Button(
            button_frame, 
            text="Начать проверку", 
            command=self.start_check,
            style="Accent.TButton"
        )
        self.start_button.pack(side=tk.LEFT, padx=(0, 5))
        
        ttk.Button(
            button_frame, 
            text="Очистить", 
            command=self.clear_results
        ).pack(side=tk.LEFT, padx=5)
        
        # Прогресс-бар
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Статус-бар
        self.status_label = ttk.Label(main_frame, text="Готов к работе", relief=tk.SUNKEN)
        self.status_label.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Результаты
        results_frame = ttk.LabelFrame(main_frame, text="Результаты", padding="10")
        results_frame.grid(row=6, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        results_frame.columnconfigure(0, weight=1)
        results_frame.rowconfigure(0, weight=1)
        main_frame.rowconfigure(6, weight=1)
        
        # Таблица результатов
        columns = ('Файл', 'Процент схожести', 'Уровень')
        self.results_tree = ttk.Treeview(results_frame, columns=columns, show='headings', height=15)
        
        for col in columns:
            self.results_tree.heading(col, text=col)
            self.results_tree.column(col, width=200)
        
        self.results_tree.column('Файл', width=400)
        self.results_tree.column('Процент схожести', width=150)
        self.results_tree.column('Уровень', width=150)
        
        # Скроллбар для таблицы
        scrollbar = ttk.Scrollbar(results_frame, orient=tk.VERTICAL, command=self.results_tree.yview)
        self.results_tree.configure(yscrollcommand=scrollbar.set)
        
        self.results_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # Детальный вывод
        detail_frame = ttk.LabelFrame(main_frame, text="Детальный вывод", padding="10")
        detail_frame.grid(row=7, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S))
        detail_frame.columnconfigure(0, weight=1)
        detail_frame.rowconfigure(0, weight=1)
        main_frame.rowconfigure(7, weight=1)
        
        self.detail_text = scrolledtext.ScrolledText(detail_frame, height=10, wrap=tk.WORD)
        self.detail_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Инициализация списка файлов
        self.refresh_file_list()
        
    def select_directory(self):
        """Выбор директории с документами"""
        directory = filedialog.askdirectory(initialdir=self.docs_directory.get())
        if directory:
            self.docs_directory.set(directory)
            self.refresh_file_list()
    
    def refresh_file_list(self):
        """Обновление списка доступных файлов"""
        try:
            docs_dir = Path(self.docs_directory.get())
            if not docs_dir.exists():
                self.file_combo['values'] = []
                self.status_label.config(text="Директория не найдена")
                return
            
            docx_files = list(docs_dir.glob("*.docx"))
            file_names = [f.name for f in sorted(docx_files)]
            
            self.file_combo['values'] = file_names
            
            if file_names and not self.target_file.get():
                self.target_file.set(file_names[0])
            
            self.status_label.config(text=f"Найдено {len(file_names)} Word файлов")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось обновить список файлов:\n{e}")
    
    def get_plagiarism_level(self, rate):
        """Определение уровня плагиата"""
        if rate >= 70:
            return "🔴 ВЫСОКИЙ"
        elif rate >= 40:
            return "🟡 СРЕДНИЙ"
        elif rate >= 10:
            return "🟢 НИЗКИЙ"
        else:
            return "⚪ МИНИМАЛЬНЫЙ"
    
    def start_check(self):
        """Запуск проверки в отдельном потоке"""
        if not Path(self.docs_directory.get()).exists():
            messagebox.showerror("Ошибка", "Директория не найдена!")
            return
        
        # Отключаем кнопку и запускаем прогресс
        self.start_button.config(state='disabled')
        self.progress.start()
        self.detail_text.delete(1.0, tk.END)
        self.results_tree.delete(*self.results_tree.get_children())
        self.status_label.config(text="Обработка файлов...")
        
        # Запускаем в отдельном потоке
        thread = threading.Thread(target=self.run_check, daemon=True)
        thread.start()
    
    def run_check(self):
        """Выполнение проверки плагиата"""
        try:
            docs_dir = Path(self.docs_directory.get())
            target_file = self.target_file.get() if self.target_file.get() else None
            
            # Создаем детектор
            self.checker = WordPlagiarismChecker(docs_dir, k_gram=5)
            
            # Обрабатываем файлы
            docx_files = list(docs_dir.glob("*.docx"))
            
            self.update_detail(f"Найдено {len(docx_files)} Word файлов\n")
            self.update_detail("="*60 + "\n")
            
            for filepath in docx_files:
                filename = filepath.name
                self.update_detail(f"Обработка файла: {filename}\n")
                content = self.checker.read_word_file(filepath)
                self.checker.calculate_hash(content, filename)
                hash_count = len(self.checker.hash_tables[filename])
                self.update_detail(f"  ✓ Создано {hash_count} хешей\n")
            
            # Определяем целевой файл
            if target_file:
                if target_file not in [f.name for f in docx_files]:
                    self.update_detail(f"Ошибка: файл {target_file} не найден\n")
                    self.finish_check()
                    return
                target_name = target_file
            else:
                target_name = docx_files[0].name if docx_files else None
                if not target_name:
                    self.update_detail("Ошибка: файлы не найдены\n")
                    self.finish_check()
                    return
            
            self.update_detail(f"\n{'='*60}\n")
            self.update_detail(f"Сравнение файла '{target_name}' с остальными:\n")
            self.update_detail(f"{'='*60}\n\n")
            
            # Сравниваем файлы
            results = []
            for filepath in docx_files:
                compare_name = filepath.name
                if compare_name == target_name:
                    continue
                
                rate = self.checker.compare_files(target_name, compare_name)
                results.append({
                    'file': compare_name,
                    'rate': rate
                })
                
                self.update_detail(f"{target_name} vs {compare_name}: {rate:.2f}%\n")
            
            # Сортируем результаты
            results.sort(key=lambda x: x['rate'], reverse=True)
            self.results = results
            
            # Обновляем таблицу
            self.update_results_table(results)
            
            # Выводим статистику
            if results:
                avg_rate = sum(r['rate'] for r in results) / len(results)
                max_rate = max(r['rate'] for r in results)
                min_rate = min(r['rate'] for r in results)
                
                self.update_detail(f"\n{'='*60}\n")
                self.update_detail("Статистика:\n")
                self.update_detail(f"  - Всего файлов обработано: {len(docx_files)}\n")
                self.update_detail(f"  - Файлов сравнено: {len(results)}\n")
                self.update_detail(f"  - Средний процент схожести: {avg_rate:.2f}%\n")
                self.update_detail(f"  - Максимальная схожесть: {max_rate:.2f}%\n")
                self.update_detail(f"  - Минимальная схожесть: {min_rate:.2f}%\n")
            
            self.status_label.config(text=f"Проверка завершена. Сравнено {len(results)} файлов")
            
        except Exception as e:
            self.update_detail(f"\nОШИБКА: {str(e)}\n")
            messagebox.showerror("Ошибка", f"Произошла ошибка при проверке:\n{e}")
            self.status_label.config(text="Ошибка при выполнении проверки")
        finally:
            self.finish_check()
    
    def update_results_table(self, results):
        """Обновление таблицы результатов"""
        self.root.after(0, self._update_table, results)
    
    def _update_table(self, results):
        """Обновление таблицы (выполняется в главном потоке)"""
        self.results_tree.delete(*self.results_tree.get_children())
        for result in results:
            level = self.get_plagiarism_level(result['rate'])
            self.results_tree.insert(
                '', 
                'end', 
                values=(result['file'], f"{result['rate']:.2f}%", level)
            )
    
    def update_detail(self, text):
        """Обновление детального вывода (потокобезопасно)"""
        self.root.after(0, self._update_detail_text, text)
    
    def _update_detail_text(self, text):
        """Обновление текста (выполняется в главном потоке)"""
        self.detail_text.insert(tk.END, text)
        self.detail_text.see(tk.END)
    
    def finish_check(self):
        """Завершение проверки"""
        self.root.after(0, self._finish_check)
    
    def _finish_check(self):
        """Завершение проверки (выполняется в главном потоке)"""
        self.progress.stop()
        self.start_button.config(state='normal')
    
    def clear_results(self):
        """Очистка результатов"""
        self.results_tree.delete(*self.results_tree.get_children())
        self.detail_text.delete(1.0, tk.END)
        self.results = []
        self.status_label.config(text="Готов к работе")


def main():
    root = tk.Tk()
    app = PlagiarismCheckerGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()

