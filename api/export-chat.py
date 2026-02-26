from http.server import BaseHTTPRequestHandler
import json
import io
import re
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from lxml import etree
import urllib.request

# =============================================
# XSLT для конвертации MathML → OMML (Word Math)
# =============================================
MATHML_TO_OMML_XSLT = None

def get_omml_xslt():
    """Загружает и кеширует XSLT трансформацию MathML→OMML"""
    global MATHML_TO_OMML_XSLT
    if MATHML_TO_OMML_XSLT is not None:
        return MATHML_TO_OMML_XSLT
    
    try:
        # Используем встроенный XSLT из python-docx или создаём свой
        import importlib.resources
        # Пробуем найти MML2OMML.XSL
        xslt_path = None
        
        # Вариант: скачиваем XSLT с GitHub
        xslt_url = 'https://raw.githubusercontent.com/nicjohnson145/planern/master/planern/MML2OMML.XSL'
        req = urllib.request.Request(xslt_url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req, timeout=10) as response:
            xslt_content = response.read()
        
        xslt_tree = etree.fromstring(xslt_content)
        MATHML_TO_OMML_XSLT = etree.XSLT(xslt_tree)
        return MATHML_TO_OMML_XSLT
    except Exception as e:
        print(f'Failed to load XSLT: {e}')
        return None


def latex_to_omml(latex_str):
    """Конвертирует LaTeX → MathML → OMML (нативный Word формат)"""
    try:
        from latex2mathml.converter import convert as latex_to_mathml
        
        # Шаг 1: LaTeX → MathML
        mathml_str = latex_to_mathml(latex_str)
        
        # Шаг 2: MathML → OMML через XSLT
        xslt = get_omml_xslt()
        if xslt is None:
            return None
        
        # Парсим MathML
        mathml_tree = etree.fromstring(mathml_str.encode('utf-8'))
        
        # Применяем XSLT трансформацию
        omml_tree = xslt(mathml_tree)
        
        # Получаем корневой элемент OMML
        omml_root = omml_tree.getroot()
        
        return omml_root
    except Exception as e:
        print(f'LaTeX to OMML conversion failed for "{latex_str}": {e}')
        return None


def add_omml_to_paragraph(paragraph, omml_element):
    """Вставляет OMML элемент в параграф Word"""
    try:
        # oMath должен быть дочерним элементом параграфа
        paragraph._element.append(omml_element)
        return True
    except Exception as e:
        print(f'Failed to add OMML to paragraph: {e}')
        return False


class handler(BaseHTTPRequestHandler):
    
    def _set_cors_headers(self):
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'POST, OPTIONS, GET')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type, Authorization')
        self.send_header('Access-Control-Max-Age', '3600')
    
    def do_OPTIONS(self):
        self.send_response(200)
        self._set_cors_headers()
        self.end_headers()

    def do_GET(self):
        self.send_response(200)
        self._set_cors_headers()
        self.send_header('Content-Type', 'application/json')
        self.end_headers()
        response = json.dumps({
            'status': 'OK',
            'message': 'Gemini Chat Export API with OMML math support',
            'version': '2.0'
        })
        self.wfile.write(response.encode())

    def do_POST(self):
        try:
            content_length = int(self.headers.get('Content-Length', 0))
            post_data = self.rfile.read(content_length)
            data = json.loads(post_data.decode('utf-8'))
            
            messages = data.get('messages', [])
            title = data.get('title', 'Gemini Chat')
            
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
            
            doc.add_paragraph()
            
            # Обрабатываем сообщения
            for i, message in enumerate(messages):
                role = message.get('role', 'user')
                content = message.get('content', '')
                
                role_name = 'You' if role == 'user' else 'Gemini'
                role_para = doc.add_paragraph()
                role_run = role_para.add_run(role_name)
                role_run.bold = True
                role_run.font.size = Pt(14)
                
                if role == 'user':
                    role_run.font.color.rgb = RGBColor(33, 150, 243)
                else:
                    role_run.font.color.rgb = RGBColor(76, 175, 80)
                
                self.process_content(doc, content)
                
                if i < len(messages) - 1:
                    separator = doc.add_paragraph()
                    sep_run = separator.add_run('─' * 60)
                    sep_run.font.color.rgb = RGBColor(200, 200, 200)
            
            # Сохраняем
            buffer = io.BytesIO()
            doc.save(buffer)
            buffer.seek(0)
            
            self.send_response(200)
            self._set_cors_headers()
            self.send_header('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')
            self.send_header('Content-Disposition', f'attachment; filename="{title}.docx"')
            self.end_headers()
            self.wfile.write(buffer.getvalue())
            
        except Exception as e:
            import traceback
            traceback.print_exc()
            self.send_response(500)
            self._set_cors_headers()
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            error_response = json.dumps({'error': str(e)})
            self.wfile.write(error_response.encode())

    def process_content(self, doc, content):
        """Обрабатывает контент с формулами, кодом, таблицами и изображениями"""
        lines = content.split('\n')
        
        i = 0
        while i < len(lines):
            line = lines[i]
            
            # Изображения
            image_pattern = r'!\[([^\]]*)\]\(([^\)]+)\)'
            if re.search(image_pattern, line):
                match = re.search(image_pattern, line)
                if match:
                    alt_text = match.group(1)
                    image_data = match.group(2)
                    try:
                        if image_data.startswith('data:image'):
                            self.add_image_from_base64(doc, image_data, alt_text)
                        else:
                            self.add_image_from_url(doc, image_data, alt_text)
                    except Exception as e:
                        para = doc.add_paragraph()
                        run = para.add_run(f'[Image: {alt_text}]')
                        run.italic = True
                    i += 1
                    continue
            
            # Блоки кода
            if line.strip().startswith('```'):
                code_lines = []
                i += 1
                while i < len(lines) and not lines[i].strip().startswith('```'):
                    code_lines.append(lines[i])
                    i += 1
                
                code_para = doc.add_paragraph()
                code_run = code_para.add_run('\n'.join(code_lines))
                code_run.font.name = 'Courier New'
                code_run.font.size = Pt(10)
                
                shading = OxmlElement('w:shd')
                shading.set(qn('w:fill'), 'F5F5F5')
                code_para._element.get_or_add_pPr().append(shading)
                
                i += 1
                continue
            
            # Таблицы
            if '|' in line and line.strip().startswith('|'):
                table_lines = []
                while i < len(lines) and '|' in lines[i]:
                    if '---' not in lines[i]:
                        table_lines.append(lines[i])
                    i += 1
                
                if table_lines:
                    self.add_markdown_table(doc, table_lines)
                continue
            
            # =============================================
            # БЛОЧНЫЕ ФОРМУЛЫ: $$...$$
            # =============================================
            block_match = re.match(r'^\s*\$\$(.+?)\$\$\s*$', line)
            if block_match:
                latex = block_match.group(1).strip()
                self.add_block_math(doc, latex)
                i += 1
                continue
            
            # Многострочная блочная формула: $$ на отдельной строке
            if line.strip() == '$$':
                formula_lines = []
                i += 1
                while i < len(lines) and lines[i].strip() != '$$':
                    formula_lines.append(lines[i])
                    i += 1
                latex = '\n'.join(formula_lines).strip()
                if latex:
                    self.add_block_math(doc, latex)
                i += 1
                continue
            
            # Обычный текст с inline формулами $...$
            if line.strip():
                self.add_text_with_inline_math(doc, line)
            else:
                doc.add_paragraph()
            
            i += 1

    def add_block_math(self, doc, latex):
        """Добавляет блочную формулу как нативную формулу Word"""
        para = doc.add_paragraph()
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        omml = latex_to_omml(latex)
        if omml is not None:
            add_omml_to_paragraph(para, omml)
        else:
            # Fallback: добавляем как стилизованный текст
            run = para.add_run(latex)
            run.font.name = 'Cambria Math'
            run.font.size = Pt(12)
            run.italic = True

    def add_text_with_inline_math(self, doc, text):
        """Добавляет текст с inline LaTeX формулами как нативными формулами Word"""
        # Разделяем текст на части: обычный текст и формулы
        # Паттерн: $...$ но НЕ $$...$$
        parts = re.split(r'(?<!\$)\$(?!\$)(.+?)(?<!\$)\$(?!\$)', text)
        
        para = doc.add_paragraph()
        
        for idx, part in enumerate(parts):
            if idx % 2 == 0:
                # Обычный текст
                if part:
                    self.add_formatted_text(para, part)
            else:
                # Inline формула
                latex = part.strip()
                omml = latex_to_omml(latex)
                if omml is not None:
                    add_omml_to_paragraph(para, omml)
                else:
                    # Fallback
                    run = para.add_run(latex)
                    run.font.name = 'Cambria Math'
                    run.italic = True

    def add_formatted_text(self, para, text):
        """Добавляет текст с форматированием (жирный, курсив)"""
        # Жирный **text**
        bold_pattern = r'\*\*(.+?)\*\*'
        parts = re.split(bold_pattern, text)
        
        for i, part in enumerate(parts):
            if i % 2 == 0:
                # Проверяем курсив *text*
                italic_pattern = r'\*(.+?)\*'
                italic_parts = re.split(italic_pattern, part)
                for j, ipart in enumerate(italic_parts):
                    if j % 2 == 0:
                        if ipart:
                            para.add_run(ipart)
                    else:
                        run = para.add_run(ipart)
                        run.italic = True
            else:
                run = para.add_run(part)
                run.bold = True

    def add_image_from_base64(self, doc, base64_data, alt_text=''):
        try:
            import base64
            if 'base64,' in base64_data:
                base64_data = base64_data.split('base64,')[1]
            image_bytes = base64.b64decode(base64_data)
            image_stream = io.BytesIO(image_bytes)
            para = doc.add_paragraph()
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run()
            run.add_picture(image_stream, width=Inches(5.0))
            if alt_text and alt_text != 'Image':
                caption_para = doc.add_paragraph()
                caption_run = caption_para.add_run(alt_text)
                caption_run.italic = True
                caption_run.font.size = Pt(10)
                caption_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        except Exception as e:
            para = doc.add_paragraph()
            run = para.add_run(f'[Failed to load image: {alt_text}]')
            run.italic = True
            run.font.color.rgb = RGBColor(150, 150, 150)

    def add_image_from_url(self, doc, url, alt_text=''):
        try:
            req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
            with urllib.request.urlopen(req, timeout=30) as response:
                image_data = response.read()
            if len(image_data) < 100:
                raise Exception('Image too small')
            image_stream = io.BytesIO(image_data)
            para = doc.add_paragraph()
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run()
            run.add_picture(image_stream, width=Inches(5.0))
            if alt_text and alt_text != 'Image':
                caption_para = doc.add_paragraph()
                caption_run = caption_para.add_run(alt_text)
                caption_run.italic = True
                caption_run.font.size = Pt(10)
                caption_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        except Exception as e:
            para = doc.add_paragraph()
            run = para.add_run(f'[Failed to load image]')
            run.italic = True
            run.font.color.rgb = RGBColor(150, 150, 150)
            link_para = doc.add_paragraph()
            link_run = link_para.add_run(f'Link: {url}')
            link_run.font.size = Pt(9)
            link_run.font.color.rgb = RGBColor(0, 0, 255)

    def add_markdown_table(self, doc, table_lines):
        if not table_lines:
            return
        rows = []
        for line in table_lines:
            cells = [cell.strip() for cell in line.split('|') if cell.strip()]
            if cells:
                rows.append(cells)
        if not rows:
            return
        
        max_cols = max(len(row) for row in rows)
        table = doc.add_table(rows=len(rows), cols=max_cols)
        table.style = 'Table Grid'
        
        for i, row_data in enumerate(rows):
            for j, cell_text in enumerate(row_data):
                if j < max_cols:
                    cell = table.rows[i].cells[j]
                    para = cell.paragraphs[0]
                    run = para.add_run(cell_text)
                    run.font.size = Pt(11)
                    if i == 0:
                        run.bold = True

    def get_current_date(self):
        from datetime import datetime
        return datetime.now().strftime('%d.%m.%Y %H:%M')
