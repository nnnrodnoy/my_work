#!/usr/bin/env python3
"""
СОЗДАТЕЛЬ DOCX ДОКУМЕНТОВ
"""

import os
import re
import sys
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

# =============================================================================
# ФУНКЦИИ ДЛЯ ВЫВОДА ASCII АРТА
# =============================================================================

def print_art():
    """Выводит встроенный ASCII арт"""
    art = """
╔══════════════════════════════════════╗
║       ДОКУМЕНТОГЕНЕРАТОР 3000       ║
║           версия 1.0.0              ║
╚══════════════════════════════════════╝
    """
    print(art)
    return True

def clear_screen():
    """Очистка экрана"""
    os.system('cls' if os.name == 'nt' else 'clear')

# =============================================================================
# КЛАСС ДЛЯ СОЗДАНИЯ DOCX ДОКУМЕНТОВ
# =============================================================================

class DocxCreator:
    def __init__(self):
        self.document = Document()
        self.setup_fixed_styles()
        
    def setup_fixed_styles(self):
        """Настройка фиксированных стилей документа"""
        # Константы
        self.DEFAULT_FONT = 'Times New Roman'
        self.DEFAULT_COLOR = RGBColor(0, 0, 0)  # Черный
        self.LINE_SPACING = 1.15
        
        # Настройка страницы
        section = self.document.sections[0]
        section.page_height = Cm(29.7)  # A4
        section.page_width = Cm(21.0)
        section.left_margin = Cm(3.0)    # Левый отступ 3 см
        section.right_margin = Cm(2.0)   # Правый отступ 2 см
        section.top_margin = Cm(2.0)     # Верхний отступ 2 см
        section.bottom_margin = Cm(2.0)  # Нижний отступ 2 см
        
        # Настройка стиля Normal
        style = self.document.styles['Normal']
        style.font.name = self.DEFAULT_FONT
        style.font.size = Pt(12)
        style.font.color.rgb = self.DEFAULT_COLOR
        style.paragraph_format.line_spacing = self.LINE_SPACING
        style.paragraph_format.space_before = Cm(0)
        style.paragraph_format.space_after = Cm(0.5)
        style.paragraph_format.first_line_indent = Cm(1.25)  # Красная строка
    
    def create_document(self, text):
        """Создать документ из текста"""
        # Разделяем на строки
        lines = text.strip().split('\n')
        
        for i, line in enumerate(lines):
            line = line.strip()
            if not line:
                continue
            
            # Обрабатываем разрыв страницы
            if line == '[PAGE_BREAK]':
                self.document.add_page_break()
                continue
            
            # Обрабатываем заголовок документа
            if line.startswith('<title>') and line.endswith('</title>'):
                self.add_title(line)
                continue
            
            # Обрабатываем строку с тегами
            self.process_line(line)
    
    def process_line(self, line):
        """Обработать одну строку с тегами"""
        line = line.strip()
        
        # Проверяем, есть ли заголовок h1-h4
        heading_match = re.search(r'<(h[1-4])>(.*?)</\1>', line)
        if heading_match:
            heading_type = heading_match.group(1)  # h1, h2, h3, h4
            heading_content = heading_match.group(2).strip()
            level = int(heading_type[1])  # Извлекаем цифру
            
            # Извлекаем выравнивание
            alignment = self.get_alignment(line)
            
            self.add_heading(heading_content, level, alignment)
        else:
            # Это обычный абзац
            alignment = self.get_alignment(line)
            clean_line = self.remove_alignment_tags(line)
            
            if clean_line:
                self.add_paragraph(clean_line, alignment)
    
    def get_alignment(self, line):
        """Определить выравнивание из тегов"""
        if '<c>' in line and '</c>' in line:
            return 'center'
        elif '<l>' in line and '</l>' in line:
            return 'left'
        elif '<p>' in line and '</p>' in line:
            return 'right'
        elif '<j>' in line and '</j>' in line:
            return 'justify'
        else:
            return 'justify'  # По умолчанию
    
    def remove_alignment_tags(self, line):
        """Убрать теги выравнивания из строки"""
        line = re.sub(r'</?(c|l|p|j)>', '', line)
        return line.strip()
    
    def add_title(self, text):
        """Добавить заголовок документа"""
        title_match = re.search(r'<title>(.*?)</title>', text)
        if title_match:
            title_text = title_match.group(1).strip()
            
            paragraph = self.document.add_paragraph()
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            paragraph.paragraph_format.space_after = Cm(1.0)
            
            run = paragraph.add_run(title_text)
            run.font.name = 'Times New Roman'
            run.font.size = Pt(20)
            run.bold = True
            run.font.color.rgb = RGBColor(0, 0, 0)
            
            self.document.add_paragraph()  # Пустая строка
    
    def add_heading(self, text, level, alignment='left'):
        """Добавить заголовок"""
        params = {
            1: {'size': 14, 'bold': False},
            2: {'size': 14, 'bold': True},
            3: {'size': 16, 'bold': True},
            4: {'size': 18, 'bold': True}
        }.get(level, {'size': 14, 'bold': False})
        
        paragraph = self.document.add_paragraph()
        
        # Устанавливаем выравнивание
        if alignment == 'center':
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif alignment == 'left':
            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
        elif alignment == 'right':
            paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        elif alignment == 'justify':
            paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        
        # Настраиваем отступы
        paragraph.paragraph_format.space_before = Cm(0.3)
        paragraph.paragraph_format.space_after = Cm(0.3)
        
        # Обрабатываем inline-форматирование
        self.process_inline_formatting(paragraph, text, params['size'], params['bold'])
    
    def add_paragraph(self, text, alignment='justify'):
        """Добавить обычный абзац"""
        paragraph = self.document.add_paragraph()
        
        # Устанавливаем выравнивание
        if alignment == 'center':
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif alignment == 'left':
            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
        elif alignment == 'right':
            paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        elif alignment == 'justify':
            paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        
        # Настраиваем отступы
        paragraph.paragraph_format.first_line_indent = Cm(1.25)  # Красная строка
        paragraph.paragraph_format.space_after = Cm(0.5)
        
        # Обрабатываем inline-форматирование
        self.process_inline_formatting(paragraph, text, 12, False)
    
    def process_inline_formatting(self, paragraph, text, size, bold):
        """Обработать inline-форматирование"""
        parts = re.split(r'(<[^>]+>)', text)
        
        current_bold = bold
        current_italic = False
        current_underline = False
        
        for part in parts:
            if not part:
                continue
            
            if part.startswith('<') and part.endswith('>'):
                tag = part[1:-1].lower()
                if tag == 'b':
                    current_bold = True
                elif tag == '/b':
                    current_bold = False
                elif tag == 'i':
                    current_italic = True
                elif tag == '/i':
                    current_italic = False
                elif tag == 'z':
                    current_underline = True
                elif tag == '/z':
                    current_underline = False
            else:
                if part.strip():
                    run = paragraph.add_run(part)
                    run.font.name = 'Times New Roman'
                    run.font.size = Pt(size)
                    run.font.color.rgb = RGBColor(0, 0, 0)
                    run.bold = current_bold
                    run.italic = current_italic
                    run.underline = current_underline
    
    def save(self, filename):
        """Сохранить документ"""
        if not filename.endswith('.docx'):
            filename += '.docx'
        
        self.document.save(filename)
        return filename

# =============================================================================
# ГЛАВНАЯ ФУНКЦИЯ ПРОГРАММЫ
# =============================================================================

def main():
    """Главная функция программы"""
    # 1. Чистим терминал
    clear_screen()
    
    # 2. Выводим арт
    print_art()
    print()  # Пустая строка
    
    # 3. Запрашиваем текст для форматирования
    print("Введите текст для форматирования:")
    print("(Для завершения ввода нажмите Ctrl+D)")
    print("-" * 40)
    
    lines = []
    try:
        while True:
            line = input()
            lines.append(line)
    except EOFError:
        # Ctrl+D был нажат
        pass
    except KeyboardInterrupt:
        # Ctrl+C был нажат
        print("\n\n❌ Ввод прерван пользователем!")
        return
    
    text = '\n'.join(lines)
    
    if not text.strip():
        print("\n❌ Текст не введен!")
        return
    
    # 4. Запрашиваем название файла
    print("\nВведите название файла (без .docx):")
    filename = input("> ").strip()
    
    if not filename:
        import datetime
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"документ_{timestamp}"
    
    # 5. Создаем документ
    print("\n🔄 Создание документа...")
    
    creator = DocxCreator()
    creator.create_document(text)
    saved_file = creator.save(filename)
    
    # 6. Выводим результат
    print(f"\n✅ Файл создан: {saved_file}")
    print("Программа завершена.")

# =============================================================================
# ЗАПУСК ПРОГРАММЫ
# =============================================================================

if __name__ == "__main__":
    main()
