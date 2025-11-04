"""
Детектор плагиата для Word файлов на русском языке
Сравнивает заданный файл со всеми остальными файлами в директории
"""
import os
import sys
from pathlib import Path
from docx import Document
import rabin_karp
import numpy as np
import re

# Русские стоп-слова
RUSSIAN_STOPWORDS = {
    'и', 'в', 'во', 'не', 'что', 'он', 'на', 'я', 'с', 'со', 'как', 'а', 'то', 'все', 
    'она', 'так', 'его', 'но', 'да', 'ты', 'к', 'у', 'же', 'вы', 'за', 'бы', 'по', 
    'только', 'ее', 'мне', 'было', 'вот', 'от', 'меня', 'еще', 'нет', 'о', 'из', 
    'ему', 'теперь', 'когда', 'даже', 'ну', 'вдруг', 'ли', 'если', 'уже', 'или', 
    'ни', 'быть', 'был', 'него', 'до', 'вас', 'нибудь', 'опять', 'уж', 'вам', 'ведь', 
    'там', 'потом', 'себя', 'ничего', 'ей', 'может', 'они', 'тут', 'где', 'есть', 
    'надо', 'ней', 'для', 'мы', 'тебя', 'их', 'чем', 'была', 'сам', 'чтоб', 'без', 
    'будто', 'чего', 'раз', 'тоже', 'себе', 'под', 'будет', 'ж', 'тогда', 'кто', 
    'этот', 'того', 'потому', 'этого', 'какой', 'совсем', 'ним', 'здесь', 'этом', 
    'один', 'почти', 'мой', 'тем', 'чтобы', 'нее', 'сейчас', 'были', 'куда', 'зачем', 
    'всех', 'никогда', 'можно', 'при', 'наконец', 'два', 'об', 'другой', 'хоть', 
    'после', 'над', 'больше', 'тот', 'через', 'эти', 'нас', 'про', 'всего', 'них', 
    'какая', 'много', 'разве', 'три', 'эту', 'моя', 'впрочем', 'хорошо', 'свою', 
    'этой', 'перед', 'иногда', 'лучше', 'чуть', 'том', 'нельзя', 'такой', 'им', 
    'более', 'всегда', 'конечно', 'всю', 'между'
}


class WordPlagiarismChecker:
    """Класс для проверки плагиата в Word документах на русском языке"""
    
    def __init__(self, docs_directory, k_gram=5):
        """
        Инициализация детектора плагиата
        
        Args:
            docs_directory: путь к директории с Word файлами
            k_gram: размер окна для k-грамм (по умолчанию 5)
        """
        self.docs_directory = Path(docs_directory)
        self.k_gram = k_gram
        self.hash_tables = {}  # Словарь: имя_файла -> список хешей
        
    def read_word_file(self, filepath):
        """
        Читает текст из Word файла (.docx)
        
        Args:
            filepath: путь к Word файлу
            
        Returns:
            str: извлеченный текст
        """
        try:
            doc = Document(filepath)
            full_text = []
            for paragraph in doc.paragraphs:
                full_text.append(paragraph.text)
            return '\n'.join(full_text)
        except Exception as e:
            print(f"Ошибка при чтении файла {filepath}: {e}")
            return ""
    
    def prepare_content(self, content):
        """
        Подготавливает текст: удаляет стоп-слова, нормализует и токенизирует
        
        Args:
            content: исходный текст
            
        Returns:
            list: список обработанных слов
        """
        # Удаляем пунктуацию и приводим к нижнему регистру
        content = re.sub(r'[^\w\s]', ' ', content.lower())
        
        # Токенизация (разбиваем на слова)
        words = content.split()
        
        # Удаляем стоп-слова и короткие слова (меньше 2 символов)
        filtered_content = []
        for word in words:
            word = word.strip()
            if word and len(word) >= 2 and word not in RUSSIAN_STOPWORDS:
                # Простая нормализация: приводим к базовой форме
                # В реальном проекте можно использовать pymorphy2 для морфологического анализа
                filtered_content.append(word)
        
        return filtered_content
    
    def calculate_hash(self, content, filename):
        """
        Вычисляет хеши для документа и сохраняет их
        
        Args:
            content: текст документа
            filename: имя файла (для идентификации)
        """
        # Подготавливаем текст
        text_words = self.prepare_content(content)
        if not text_words:
            print(f"Предупреждение: файл {filename} не содержит текста после обработки")
            self.hash_tables[filename] = []
            return
        
        # Объединяем слова в строку (без пробелов для k-грамм)
        text = "".join(text_words)
        
        if len(text) < self.k_gram:
            print(f"Предупреждение: файл {filename} слишком короткий для анализа")
            self.hash_tables[filename] = []
            return
        
        # Создаем rolling hash
        text_hash = rabin_karp.rolling_hash(text, self.k_gram)
        
        # Вычисляем хеши для всех окон
        hash_list = []
        for _ in range(len(text) - self.k_gram + 1):
            hash_list.append(text_hash.hash)
            if text_hash.next_window() == False:
                break
        
        self.hash_tables[filename] = hash_list
    
    def compare_files(self, file_a, file_b):
        """
        Сравнивает два файла и возвращает процент плагиата
        
        Args:
            file_a: имя первого файла
            file_b: имя второго файла
            
        Returns:
            float: процент плагиата (0-100)
        """
        if file_a not in self.hash_tables or file_b not in self.hash_tables:
            return 0.0
        
        hash_a = np.array(self.hash_tables[file_a])
        hash_b = np.array(self.hash_tables[file_b])
        
        if len(hash_a) == 0 or len(hash_b) == 0:
            return 0.0
        
        # Находим пересечение хешей
        sh = len(np.intersect1d(hash_a, hash_b))
        th_a = len(hash_a)
        th_b = len(hash_b)
        
        # Формула для расчета процента плагиата
        # P = (2 * SH / (THA + THB)) * 100%
        if th_a + th_b == 0:
            return 0.0
        
        plagiarism_rate = (float(2 * sh) / (th_a + th_b)) * 100
        return plagiarism_rate
    
    def process_directory(self, target_file=None):
        """
        Обрабатывает все Word файлы в директории
        
        Args:
            target_file: имя файла для сравнения (если None, сравнивает все со всеми)
        """
        # Находим все .docx файлы в директории
        docx_files = list(self.docs_directory.glob("*.docx"))
        
        if not docx_files:
            print(f"В директории {self.docs_directory} не найдено Word файлов (.docx)")
            return
        
        print(f"Найдено {len(docx_files)} Word файлов")
        
        # Обрабатываем все файлы
        for filepath in docx_files:
            filename = filepath.name
            print(f"Обработка файла: {filename}")
            content = self.read_word_file(filepath)
            self.calculate_hash(content, filename)
        
        # Определяем целевой файл для сравнения
        if target_file:
            if target_file not in [f.name for f in docx_files]:
                print(f"Ошибка: файл {target_file} не найден в директории")
                return
            target_name = target_file
        else:
            # Если файл не указан, используем первый
            if docx_files:
                target_name = docx_files[0].name
                print(f"Целевой файл не указан, используем: {target_name}")
            else:
                return
        
        # Сравниваем целевой файл со всеми остальными
        print(f"\n{'='*60}")
        print(f"Сравнение файла '{target_name}' с остальными:")
        print(f"{'='*60}")
        
        results = []
        for filepath in docx_files:
            compare_name = filepath.name
            if compare_name == target_name:
                continue
            
            rate = self.compare_files(target_name, compare_name)
            results.append({
                'file': compare_name,
                'rate': rate
            })
            print(f"{target_name} vs {compare_name}: {rate:.2f}%")
        
        # Сортируем результаты по проценту плагиата
        results.sort(key=lambda x: x['rate'], reverse=True)
        
        print(f"\n{'='*60}")
        print("Результаты (отсортированы по проценту плагиата):")
        print(f"{'='*60}")
        for result in results:
            print(f"{result['file']}: {result['rate']:.2f}%")
        
        return results


def main():
    """Основная функция для запуска детектора плагиата"""
    # Получаем путь к директории docs
    current_dir = Path(__file__).parent
    docs_dir = current_dir.parent / "docs"
    
    # Проверяем наличие директории
    if not docs_dir.exists():
        print(f"Ошибка: директория {docs_dir} не найдена")
        return
    
    # Создаем экземпляр детектора
    checker = WordPlagiarismChecker(docs_dir, k_gram=5)
    
    # Получаем целевой файл из аргументов командной строки (если указан)
    target_file = None
    if len(sys.argv) > 1:
        target_file = sys.argv[1]
    
    # Обрабатываем директорию
    checker.process_directory(target_file=target_file)


if __name__ == "__main__":
    main()

