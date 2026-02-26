from http.server import BaseHTTPRequestHandler
import json
import io
import re
import traceback
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn, nsmap
from lxml import etree
import urllib.request

# =============================================
# LaTeX → OMML (нативные формулы Word)
# =============================================

# Word namespace для математики
WORD_MATH_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math'

# Кешируем XSLT
_XSLT_TRANSFORM = None

def get_xslt_transform():
    """Загружает и кеширует XSLT для MathML→OMML"""
    global _XSLT_TRANSFORM
    if _XSLT_TRANSFORM is not None:
        return _XSLT_TRANSFORM
    
    # Список зеркал для скачивания MML2OMML.XSL
    xslt_urls = [
        'https://raw.githubusercontent.com/nicjohnson145/planern/master/planern/MML2OMML.XSL',
        'https://raw.githubusercontent.com/oerpub/mathconverter/master/MML2OMML.XSL',
        'https://raw.githubusercontent.com/pjheslin/diogenes/master/server/MML2OMML.XSL',
    ]
    
    for url in xslt_urls:
        try:
            req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
            with urllib.request.urlopen(req, timeout=15) as response:
                xslt_content = response.read()
            xslt_tree = etree.fromstring(xslt_content)
            _XSLT_TRANSFORM = etree.XSLT(xslt_tree)
            print(f'XSLT loaded successfully from {url}')
            return _XSLT_TRANSFORM
        except Exception as e:
            print(f'Failed to load XSLT from {url}: {e}')
            continue
    
    print('All XSLT sources failed')
    return None


def latex_to_omml(latex_str):
    """LaTeX → MathML → OMML"""
    try:
        from latex2mathml.converter import convert as latex_to_mathml
        
        # LaTeX → MathML
        mathml_str = latex_to_mathml(latex_str)
        
        # MathML → OMML
        xslt = get_xslt_transform()
        if xslt is None:
            return None
        
        mathml_tree = etree.fromstring(mathml_str.encode('utf-8'))
        omml_tree = xslt(mathml_tree)
        omml_root = omml_tree.getroot()
        
        return omml_root
    except Exception as e:
        print(f'LaTeX→OMML failed for "{latex_str}": {e}')
        return None


def add_math_to_paragraph(para, latex_str):
    """Добавляет формулу в параграф. Возвращает True если успешно."""
    omml = latex_to_omml(latex_str)
    if omml is not None:
        try:
            para._element.append(omml)
            return True
        except:
            pass
    
    # Fallback: красивый текст формулы
    run = para.add_run(latex_str)
    run.font.name = 'Cambria Math'
    run.font.size = Pt(12)
    run.italic = True
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
        
        # Тестируем конвертацию при GET запросе
        test_result = 'not tested'
        try:
            omml = latex_to_omml(r'\frac{a}{b}')
            test_result = 'OK' if omml is not None else 'XSLT failed'
        except Exception as e:
            test_result = f'Error: {str(e)}'
        
        response = json.dumps({
            'status': 'OK',
            'message': 'Gemini Chat Export API v2 - OMML Math',
            'math_test': test_result
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
                self.wfile.write(json.dumps({'error': 'No messages'}).encode())
                return
            
            doc = Document()
            
            # Заголовок
            heading = doc.add_heading(title, level=1)
            heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            date_para = doc.add_paragraph()
            date_run = date_para.add_run(self._get_date())
            date_run.font.size = Pt(10)
            date_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            doc.add_paragraph()
            
            for i, msg in enumerate(messages):
                role = msg.get('role', 'user')
                content = msg.get('content', '')
                
                # Роль
                role_para = doc.add_paragraph()
                role_run = role_para.add_run('You' if role == 'user' else 'Gemini')
                role_run.bold = True
                role_run.font.size = Pt(14)
                role_run.font.color.rgb = RGBColor(33, 150, 243) if role == 'user' else RGBColor(76, 175, 80)
                
                # Контент
                self._process_content(doc, content)
                
                # Разделитель
                if i < len(messages) - 1:
                    sep = doc.add_paragraph()
                    sep_run = sep.add_run('─' * 60)
                    sep_run.font.color.rgb = RGBColor(200, 200, 200)
            
            buffer = io.BytesIO()
            doc.save(buffer)
            buffer.seek(0)
            
            self.send_response(200)
            self._set_cors_headers()
            self.send_header('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')
            self.send_header('Content-Disposition', f'attachment; filename="gemini-chat.docx"')
            self.end_headers()
            self.wfile.write(buffer.getvalue())
            
        except Exception as e:
            traceback.print_exc()
            self.send_response(500)
            self._set_cors_headers()
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({'error': str(e)}).encode())

    def _process_content(self, doc, content):
        """Обрабатывает контент сообщения"""
        lines = content.split('\n')
        i = 0
        
        while i < len(lines):
            line = lines[i]
            
            # --- Изображения ---
            img_match = re.search(r'!\[([^\]]*)\]\(([^\)]+)\)', line)
            if img_match:
                self._add_image(doc, img_match.group(2), img_match.group(1))
                i += 1
                continue
            
            # --- Блоки кода ---
            if line.strip().startswith('```'):
                code_lines = []
                i += 1
                while i < len(lines) and not lines[i].strip().startswith('```'):
                    code_lines.append(lines[i])
                    i += 1
                self._add_code_block(doc, '\n'.join(code_lines))
                i += 1
                continue
            
            # --- Таблицы ---
            if '|' in line and line.strip().startswith('|'):
                table_lines = []
                while i < len(lines) and '|' in lines[i]:
                    if '---' not in lines[i]:
                        table_lines.append(lines[i])
                    i += 1
                if table_lines:
                    self._add_table(doc, table_lines)
                continue
            
            # --- Блочная формула $$...$$ на одной строке ---
            block_match = re.match(r'^\s*\$\$(.+?)\$\$\s*$', line)
            if block_match:
                latex = block_match.group(1).strip()
                para = doc.add_paragraph()
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                add_math_to_paragraph(para, latex)
                i += 1
                continue
            
            # --- Блочная формула $$\n...\n$$ многострочная ---
            if line.strip() == '$$':
                formula_lines = []
                i += 1
                while i < len(lines) and lines[i].strip() != '$$':
                    formula_lines.append(lines[i])
                    i += 1
                latex = ' '.join(formula_lines).strip()
                if latex:
                    para = doc.add_paragraph()
                    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    add_math_to_paragraph(para, latex)
                i += 1
                continue
            
            # --- Обычный текст (может содержать inline $...$) ---
            if line.strip():
                self._add_text_with_math(doc, line)
            else:
                doc.add_paragraph()
            
            i += 1

    def _add_text_with_math(self, doc, text):
        """Текст с inline формулами $...$"""
        # Разбиваем по $...$ (но не $$)
        parts = re.split(r'(?<!\$)\$(?!\$)(.+?)(?<!\$)\$(?!\$)', text)
        
        if len(parts) == 1:
            # Нет формул — просто текст
            para = doc.add_paragraph()
            self._add_formatted_run(para, text)
            return
        
        para = doc.add_paragraph()
        for idx, part in enumerate(parts):
            if idx % 2 == 0:
                # Обычный текст
                if part:
                    self._add_formatted_run(para, part)
            else:
                # Inline формула
                add_math_to_paragraph(para, part.strip())

    def _add_formatted_run(self, para, text):
        """Текст с **жирным** и *курсивом*"""
        # Сначала жирный
        bold_parts = re.split(r'\*\*(.+?)\*\*', text)
        for i, part in enumerate(bold_parts):
            if i % 2 == 0:
                # Проверяем курсив
                italic_parts = re.split(r'\*(.+?)\*', part)
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

    def _add_code_block(self, doc, code):
        """Блок кода с серым фоном"""
        para = doc.add_paragraph()
        run = para.add_run(code)
        run.font.name = 'Courier New'
        run.font.size = Pt(10)
        shading = OxmlElement('w:shd')
        shading.set(qn('w:fill'), 'F5F5F5')
        para._element.get_or_add_pPr().append(shading)

    def _add_image(self, doc, src, alt=''):
        """Изображение"""
        try:
            if src.startswith('data:image'):
                import base64
                b64 = src.split('base64,')[1]
                img_bytes = base64.b64decode(b64)
                stream = io.BytesIO(img_bytes)
            else:
                req = urllib.request.Request(src, headers={'User-Agent': 'Mozilla/5.0'})
                with urllib.request.urlopen(req, timeout=30) as resp:
                    img_bytes = resp.read()
                stream = io.BytesIO(img_bytes)
            
            para = doc.add_paragraph()
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run()
            run.add_picture(stream, width=Inches(5.0))
            
            if alt and alt != 'Image':
                cap = doc.add_paragraph()
                cap_run = cap.add_run(alt)
                cap_run.italic = True
                cap_run.font.size = Pt(10)
                cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        except:
            para = doc.add_paragraph()
            run = para.add_run(f'[Image: {alt}]')
            run.italic = True
            run.font.color.rgb = RGBColor(150, 150, 150)

    def _add_table(self, doc, table_lines):
        """Markdown таблица"""
        rows = []
        for line in table_lines:
            cells = [c.strip() for c in line.split('|') if c.strip()]
            if cells:
                rows.append(cells)
        if not rows:
            return
        max_cols = max(len(r) for r in rows)
        table = doc.add_table(rows=len(rows), cols=max_cols)
        table.style = 'Table Grid'
        for i, row_data in enumerate(rows):
            for j, cell_text in enumerate(row_data):
                if j < max_cols:
                    cell = table.rows[i].cells[j]
                    run = cell.paragraphs[0].add_run(cell_text)
                    run.font.size = Pt(11)
                    if i == 0:
                        run.bold = True

    def _get_date(self):
        from datetime import datetime
        return datetime.now().strftime('%d.%m.%Y %H:%M')
