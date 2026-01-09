from http.server import BaseHTTPRequestHandler
import json
import io
import re
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from latex2mathml.converter import convert as latex_to_mathml
from lxml import etree

class handler(BaseHTTPRequestHandler):
    
    def _set_cors_headers(self):
        """Устанавливает CORS заголовки"""
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'POST, OPTIONS, GET')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type, Authorization')
        self.send_header('Access-Control-Max-Age', '3600')
    
    def do_OPTIONS(self):
        """Обрабатывает preflight запрос"""
        self.send_response(200)
        self._set_cors_headers()
        self.end_headers()

    def do_GET(self):
        """Тестовый GET запрос"""
        self.send_response(200)
        self._set_cors_headers()
        self.send_header('Content-Type', 'application/json')
        self.end_headers()
        response = json.dumps({
            'status': 'OK',
            'message': 'Gemini Chat Export API is running',
            'version': '1.0'
        })
        self.wfile.write(response.encode())

    def do_POST(self):
        try:
            # Читаем данные
            content_length = int(self.headers.get('Content-Length', 0))
            post_data = self.rfile.read(content_length)
            data = json.loads(post_data.decode('utf-8'))
            
            messages = data.get('messages', [])
            title = data.get('title', 'Чат с Gemini')
            
            if not messages:
                self.send_response(400)
                self._set_cors_headers()
                self.send_header('Content-Type', 'application/json')
                self.end_headers()
                error = json.dumps({'error': 'No messages provided'})
                self.wfile.write(error.encode())
                return
            
            # Создаём документ
            doc = Document()
            
            # Заголовок
            heading = doc.add_heading(title, level=1)
            heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            # Дата
            date_para = doc.add_paragraph()
            date_run = date_para.add_run(self.get_current_date())
            date_run.font.size = Pt(10)
            date_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            doc.add_paragraph()  # Отступ
            
            # Обрабатываем сообщения
            for i, message in enumerate(messages):
                role = message.get('role', 'user')
                content = message.get('content', '')
                
                # Имя роли
                role_name = 'Вы' if role == 'user' else 'Gemini'
                role_para = doc.add_paragraph()
                role_run = role_para.add_run(role_name)
                role_run.bold = True
                role_run.font.size = Pt(14)
                
                if role == 'user':
                    role_run.font.color.rgb = RGBColor(33, 150, 243)  # Синий
                else:
                    role_run.font.color.rgb = RGBColor(76, 175, 80)   # Зелёный
                
                # Контент
                self.process_content(doc, content)
                
                # Разделитель
                if i < len(messages) - 1:
                    doc.add_paragraph('─' * 60)
            
            # Сохраняем в буфер
            buffer = io.BytesIO()
            doc.save(buffer)
            buffer.seek(0)
            
            # Отправляем
            self.send_response(200)
            self._set_cors_headers()
            self.send_header('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')
            self.send_header('Content-Disposition', f'attachment; filename="{title}.docx"')
            self.end_headers()
            self.wfile.write(buffer.getvalue())
            
        except Exception as e:
            self.send_response(500)
            self._set_cors_headers()
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            error_response = json.dumps({'error': str(e)})
            self.wfile.write(error_response.encode())

    def process_content(self, doc, content):
        """Обрабатывает контент с формулами, кодом и таблицами"""
        lines = content.split('\n')
        
        i = 0
        while i < len(lines):
            line = lines[i]
            
            # Блоки кода ```
            if line.strip().startswith('```'):
                code_lines = []
                i += 1
                while i < len(lines) and not lines[i].strip().startswith('```'):
                    code_lines.append(lines[i])
                    i += 1
                
                # Добавляем код
                code_para = doc.add_paragraph()
                code_run = code_para.add_run('\n'.join(code_lines))
                code_run.font.name = 'Courier New'
                code_run.font.size = Pt(10)
                
                # Фон (серый)
                shading = OxmlElement('w:shd')
                shading.set(qn('w:fill'), 'F5F5F5')
                code_para._element.get_or_add_pPr().append(shading)
                
                i += 1
                continue
            
            # Таблицы markdown
            if '|' in line and line.strip().startswith('|'):
                table_lines = []
                while i < len(lines) and '|' in lines[i]:
                    if '---' not in lines[i]:  # Пропускаем разделитель
                        table_lines.append(lines[i])
                    i += 1
                
                if table_lines:
                    self.add_markdown_table(doc, table_lines)
                continue
            
            # LaTeX формулы: $...$ или $$...$$
            inline_pattern = r'\$([^\$]+)\$'
            block_pattern = r'\$\$([^\$]+)\$\$'
            
            # Проверяем блочные формулы
            if '$$' in line:
                match = re.search(block_pattern, line)
                if match:
                    latex = match.group(1).strip()
                    self.add_math_formula(doc, latex, block=True)
                    i += 1
                    continue
            
            # Обычный текст с возможными inline формулами
            if line.strip():
                self.add_text_with_inline_math(doc, line)
            else:
                doc.add_paragraph()  # Пустая строка
            
            i += 1

    def add_text_with_inline_math(self, doc, text):
        """Добавляет текст с inline LaTeX формулами"""
        inline_pattern = r'\$([^\$]+)\$'
        parts = re.split(inline_pattern, text)
        
        para = doc.add_paragraph()
        
        for i, part in enumerate(parts):
            if i % 2 == 0:
                # Обычный текст
                if part:
                    self.add_formatted_text(para, part)
            else:
                # Формула
                try:
                    self.add_inline_math_to_paragraph(para, part.strip())
                except:
                    para.add_run(f'[{part}]').font.italic = True

    def add_formatted_text(self, para, text):
        """Добавляет текст с форматированием (жирный, курсив)"""
        # Жирный текст **text**
        bold_pattern = r'\*\*(.+?)\*\*'
        parts = re.split(bold_pattern, text)
        
        for i, part in enumerate(parts):
            if i % 2 == 0:
                if part:
                    para.add_run(part)
            else:
                run = para.add_run(part)
                run.bold = True

    def add_math_formula(self, doc, latex, block=True):
        """Добавляет математическую формулу (блочную)"""
        try:
            mathml = latex_to_mathml(latex)
            
            para = doc.add_paragraph()
            if block:
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            run = para.add_run(f'[Формула: {latex}]')
            run.italic = True
            run.font.color.rgb = RGBColor(100, 100, 100)
            
        except Exception as e:
            para = doc.add_paragraph()
            run = para.add_run(f'[Формула: {latex}]')
            run.italic = True

    def add_inline_math_to_paragraph(self, para, latex):
        """Добавляет inline формулу в параграф"""
        run = para.add_run(f'[{latex}]')
        run.italic = True
        run.font.color.rgb = RGBColor(100, 100, 100)

    def add_markdown_table(self, doc, table_lines):
        """Добавляет таблицу из markdown"""
        if not table_lines:
            return
        
        # Парсим строки
        rows = []
        for line in table_lines:
            cells = [cell.strip() for cell in line.split('|') if cell.strip()]
            if cells:
                rows.append(cells)
        
        if not rows:
            return
        
        # Создаём таблицу
        table = doc.add_table(rows=len(rows), cols=len(rows[0]))
        table.style = 'Table Grid'
        
        for i, row_data in enumerate(rows):
            for j, cell_text in enumerate(row_data):
                cell = table.rows[i].cells[j]
                para = cell.paragraphs[0]
                run = para.add_run(cell_text)
                run.font.size = Pt(11)
                
                # Заголовок жирным
                if i == 0:
                    run.bold = True

    def get_current_date(self):
        from datetime import datetime
        return datetime.now().strftime('%d.%m.%Y %H:%M')
