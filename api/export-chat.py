from http.server import BaseHTTPRequestHandler
import json
import io
import re
import traceback
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from lxml import etree
import copy

# =============================================
# OMML builder — строим Word Math XML напрямую
# =============================================

MATH_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

def M(tag):
    """Создаёт элемент в math namespace"""
    return etree.SubElement.__func__  # не используем, ниже вручную

def make_el(ns, tag):
    return etree.Element(f'{{{ns}}}{tag}')

def sub_el(parent, ns, tag):
    el = etree.SubElement(parent, f'{{{ns}}}{tag}')
    return el

def make_run(text, italic=True):
    """Создаёт m:r с текстом"""
    r = make_el(MATH_NS, 'r')
    # Свойства шрифта
    rpr = sub_el(r, MATH_NS, 'rPr')
    sty = sub_el(rpr, MATH_NS, 'sty')
    sty.set(f'{{{MATH_NS}}}val', 'i' if italic else 'p')
    # Текст
    t = sub_el(r, MATH_NS, 't')
    t.text = text
    t.set(f'{{{W_NS}}}space', 'preserve')
    return r

def make_frac(num_elements, den_elements):
    """Создаёт дробь m:f"""
    f = make_el(MATH_NS, 'f')
    # fPr
    fpr = sub_el(f, MATH_NS, 'fPr')
    ftype = sub_el(fpr, MATH_NS, 'type')
    ftype.set(f'{{{MATH_NS}}}val', 'bar')
    # числитель
    num = sub_el(f, MATH_NS, 'num')
    for el in num_elements:
        num.append(el)
    # знаменатель
    den = sub_el(f, MATH_NS, 'den')
    for el in den_elements:
        den.append(el)
    return f

def make_sup(base_elements, sup_elements):
    """Верхний индекс m:sSup"""
    ssup = make_el(MATH_NS, 'sSup')
    e = sub_el(ssup, MATH_NS, 'e')
    for el in base_elements:
        e.append(el)
    s = sub_el(ssup, MATH_NS, 'sup')
    for el in sup_elements:
        s.append(el)
    return ssup

def make_sub(base_elements, sub_elements):
    """Нижний индекс m:sSub"""
    ssub = make_el(MATH_NS, 'sSub')
    e = sub_el(ssub, MATH_NS, 'e')
    for el in base_elements:
        e.append(el)
    s = sub_el(ssub, MATH_NS, 'sub')
    for el in sub_elements:
        s.append(el)
    return ssub

def make_sqrt(content_elements):
    """Корень m:rad"""
    rad = make_el(MATH_NS, 'rad')
    radpr = sub_el(rad, MATH_NS, 'radPr')
    deghide = sub_el(radpr, MATH_NS, 'degHide')
    deghide.set(f'{{{MATH_NS}}}val', '1')
    deg = sub_el(rad, MATH_NS, 'deg')
    e = sub_el(rad, MATH_NS, 'e')
    for el in content_elements:
        e.append(el)
    return rad

def make_accent(base_elements, accent_char='\u0302'):
    """Акцент (шапочка и т.д.) m:acc"""
    acc = make_el(MATH_NS, 'acc')
    accpr = sub_el(acc, MATH_NS, 'accPr')
    ch = sub_el(accpr, MATH_NS, 'chr')
    ch.set(f'{{{MATH_NS}}}val', accent_char)
    e = sub_el(acc, MATH_NS, 'e')
    for el in base_elements:
        e.append(el)
    return acc

def make_delim(content_elements, beg='(', end=')'):
    """Скобки m:d"""
    d = make_el(MATH_NS, 'd')
    dpr = sub_el(d, MATH_NS, 'dPr')
    begchr = sub_el(dpr, MATH_NS, 'begChr')
    begchr.set(f'{{{MATH_NS}}}val', beg)
    endchr = sub_el(dpr, MATH_NS, 'endChr')
    endchr.set(f'{{{MATH_NS}}}val', end)
    e = sub_el(d, MATH_NS, 'e')
    for el in content_elements:
        e.append(el)
    return d

# Таблица греческих букв
GREEK = {
    r'\alpha': 'α', r'\beta': 'β', r'\gamma': 'γ', r'\delta': 'δ',
    r'\epsilon': 'ε', r'\zeta': 'ζ', r'\eta': 'η', r'\theta': 'θ',
    r'\iota': 'ι', r'\kappa': 'κ', r'\lambda': 'λ', r'\mu': 'μ',
    r'\nu': 'ν', r'\xi': 'ξ', r'\pi': 'π', r'\rho': 'ρ',
    r'\sigma': 'σ', r'\tau': 'τ', r'\upsilon': 'υ', r'\phi': 'φ',
    r'\chi': 'χ', r'\psi': 'ψ', r'\omega': 'ω',
    r'\Gamma': 'Γ', r'\Delta': 'Δ', r'\Theta': 'Θ', r'\Lambda': 'Λ',
    r'\Xi': 'Ξ', r'\Pi': 'Π', r'\Sigma': 'Σ', r'\Phi': 'Φ',
    r'\Psi': 'Ψ', r'\Omega': 'Ω',
    r'\hbar': 'ℏ', r'\infty': '∞', r'\partial': '∂',
    r'\nabla': '∇', r'\pm': '±', r'\mp': '∓',
    r'\times': '×', r'\cdot': '·', r'\leq': '≤', r'\geq': '≥',
    r'\neq': '≠', r'\approx': '≈', r'\equiv': '≡',
    r'\sum': '∑', r'\prod': '∏', r'\int': '∫',
    r'\leftarrow': '←', r'\rightarrow': '→', r'\Rightarrow': '⇒',
    r'\Leftarrow': '⇐', r'\leftrightarrow': '↔',
}

def parse_latex_to_omml(latex):
    """
    Парсит LaTeX строку и возвращает список OMML элементов.
    Поддерживает: \frac, ^, _, \hat, \sqrt, \left \right, греческие буквы.
    """
    elements = []
    i = 0
    s = latex.strip()
    
    while i < len(s):
        c = s[i]
        
        # Пробелы
        if c == ' ':
            i += 1
            continue
        
        # Группа в {}
        if c == '{':
            # Находим соответствующую }
            depth = 1
            j = i + 1
            while j < len(s) and depth > 0:
                if s[j] == '{': depth += 1
                elif s[j] == '}': depth -= 1
                j += 1
            inner = s[i+1:j-1]
            inner_els = parse_latex_to_omml(inner)
            elements.extend(inner_els)
            i = j
            continue
        
        # Команды LaTeX
        if c == '\\':
            # Считываем команду
            j = i + 1
            while j < len(s) and s[j].isalpha():
                j += 1
            cmd = s[i:j]
            
            # \frac{num}{den}
            if cmd == r'\frac':
                num_content, after_num = _read_group(s, j)
                den_content, after_den = _read_group(s, after_num)
                num_els = parse_latex_to_omml(num_content)
                den_els = parse_latex_to_omml(den_content)
                if not num_els: num_els = [make_run(' ')]
                if not den_els: den_els = [make_run(' ')]
                elements.append(make_frac(num_els, den_els))
                i = after_den
                continue
            
            # \hat{x}
            if cmd == r'\hat':
                content, after = _read_group(s, j)
                inner_els = parse_latex_to_omml(content)
                if not inner_els: inner_els = [make_run(' ')]
                elements.append(make_accent(inner_els, '\u0302'))
                i = after
                continue
            
            # \vec{x}
            if cmd == r'\vec':
                content, after = _read_group(s, j)
                inner_els = parse_latex_to_omml(content)
                if not inner_els: inner_els = [make_run(' ')]
                elements.append(make_accent(inner_els, '\u20D7'))
                i = after
                continue
            
            # \bar{x}
            if cmd == r'\bar':
                content, after = _read_group(s, j)
                inner_els = parse_latex_to_omml(content)
                if not inner_els: inner_els = [make_run(' ')]
                elements.append(make_accent(inner_els, '\u0305'))
                i = after
                continue
            
            # \sqrt{x}
            if cmd == r'\sqrt':
                content, after = _read_group(s, j)
                inner_els = parse_latex_to_omml(content)
                if not inner_els: inner_els = [make_run(' ')]
                elements.append(make_sqrt(inner_els))
                i = after
                continue
            
            # \left( ... \right)
            if cmd == r'\left':
                beg_char = s[j] if j < len(s) else '('
                # Ищем \right
                right_pos = s.find(r'\right', j+1)
                if right_pos >= 0:
                    inner = s[j+1:right_pos]
                    end_pos = right_pos + 6
                    end_char = s[end_pos] if end_pos < len(s) else ')'
                    inner_els = parse_latex_to_omml(inner)
                    if not inner_els: inner_els = [make_run(' ')]
                    elements.append(make_delim(inner_els, beg_char, end_char))
                    i = end_pos + 1
                else:
                    elements.append(make_run(beg_char))
                    i = j + 1
                continue
            
            # \right — обрабатывается в \left
            if cmd == r'\right':
                i = j + 1
                continue
            
            # Греческие и спец символы
            if cmd in GREEK:
                elements.append(make_run(GREEK[cmd], italic=False))
                i = j
                continue
            
            # Неизвестная команда — выводим как текст
            elements.append(make_run(cmd[1:]))
            i = j
            continue
        
        # Верхний индекс ^
        if c == '^':
            sup_content, after = _read_group_or_char(s, i+1)
            sup_els = parse_latex_to_omml(sup_content)
            if not sup_els: sup_els = [make_run(' ')]
            
            # Берём последний элемент как базу
            if elements:
                base = elements.pop()
                elements.append(make_sup([base], sup_els))
            else:
                elements.append(make_sup([make_run(' ')], sup_els))
            i = after
            continue
        
        # Нижний индекс _
        if c == '_':
            sub_content, after = _read_group_or_char(s, i+1)
            sub_els = parse_latex_to_omml(sub_content)
            if not sub_els: sub_els = [make_run(' ')]
            
            if elements:
                base = elements.pop()
                elements.append(make_sub([base], sub_els))
            else:
                elements.append(make_sub([make_run(' ')], sub_els))
            i = after
            continue
        
        # Обычные символы — группируем текст
        text = ''
        while i < len(s) and s[i] not in '\\{}^_$ ':
            if s[i] in '+-=()[]|<>,.:;!?':
                if text:
                    elements.append(make_run(text))
                    text = ''
                elements.append(make_run(s[i], italic=False))
                i += 1
                continue
            text += s[i]
            i += 1
        if text:
            elements.append(make_run(text))
        continue
    
    return elements


def _read_group(s, pos):
    """Читает {content} начиная с pos. Возвращает (content, pos_after)"""
    # Пропускаем пробелы
    while pos < len(s) and s[pos] == ' ':
        pos += 1
    
    if pos >= len(s):
        return ('', pos)
    
    if s[pos] == '{':
        depth = 1
        j = pos + 1
        while j < len(s) and depth > 0:
            if s[j] == '{': depth += 1
            elif s[j] == '}': depth -= 1
            j += 1
        return (s[pos+1:j-1], j)
    else:
        # Один символ
        return (s[pos], pos+1)


def _read_group_or_char(s, pos):
    """Читает {content} или один символ"""
    while pos < len(s) and s[pos] == ' ':
        pos += 1
    if pos >= len(s):
        return ('', pos)
    if s[pos] == '{':
        return _read_group(s, pos)
    elif s[pos] == '\\':
        # Команда
        j = pos + 1
        while j < len(s) and s[j].isalpha():
            j += 1
        return (s[pos:j], j)
    else:
        return (s[pos], pos+1)


def build_omath(latex):
    """Строит полный m:oMath элемент из LaTeX"""
    omath = make_el(MATH_NS, 'oMath')
    elements = parse_latex_to_omml(latex)
    for el in elements:
        omath.append(el)
    return omath


def insert_math(paragraph, latex):
    """Вставляет формулу в параграф"""
    try:
        omath = build_omath(latex)
        paragraph._element.append(omath)
        return True
    except Exception as e:
        print(f'Math error "{latex}": {e}')
        traceback.print_exc()
        r = paragraph.add_run(latex)
        r.font.name = 'Cambria Math'
        r.italic = True
        return False


def add_block_formula(doc, latex):
    """Блочная формула"""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    insert_math(p, latex)


# =============================================
# HTTP Handler
# =============================================

class handler(BaseHTTPRequestHandler):
    
    def _cors(self):
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'POST, OPTIONS, GET')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.send_header('Access-Control-Max-Age', '3600')
    
    def do_OPTIONS(self):
        self.send_response(200)
        self._cors()
        self.end_headers()

    def do_GET(self):
        self.send_response(200)
        self._cors()
        self.send_header('Content-Type', 'application/json')
        self.end_headers()
        
        test = 'not tested'
        try:
            omath = build_omath(r'\frac{a}{b}')
            tag = etree.QName(omath.tag).localname
            children = len(list(omath))
            test = f'OK (tag={tag}, children={children})'
        except Exception as e:
            test = f'Error: {str(e)}'
        
        r = json.dumps({
            'status': 'OK',
            'version': '4.0-direct-omml',
            'math_test': test
        })
        self.wfile.write(r.encode())

    def do_POST(self):
        try:
            length = int(self.headers.get('Content-Length', 0))
            data = json.loads(self.rfile.read(length).decode('utf-8'))
            
            messages = data.get('messages', [])
            title = data.get('title', 'Gemini Chat')
            
            if not messages:
                self.send_response(400)
                self._cors()
                self.send_header('Content-Type', 'application/json')
                self.end_headers()
                self.wfile.write(b'{"error":"No messages"}')
                return
            
            doc = Document()
            
            h = doc.add_heading(title, level=1)
            h.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            dp = doc.add_paragraph()
            dr = dp.add_run(self._date())
            dr.font.size = Pt(10)
            dp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            doc.add_paragraph()
            
            for i, msg in enumerate(messages):
                role = msg.get('role', 'user')
                content = msg.get('content', '')
                
                rp = doc.add_paragraph()
                rr = rp.add_run('You' if role == 'user' else 'Gemini')
                rr.bold = True
                rr.font.size = Pt(14)
                rr.font.color.rgb = RGBColor(33, 150, 243) if role == 'user' else RGBColor(76, 175, 80)
                
                self._process(doc, content)
                
                if i < len(messages) - 1:
                    sp = doc.add_paragraph()
                    sr = sp.add_run('─' * 60)
                    sr.font.color.rgb = RGBColor(200, 200, 200)
            
            buf = io.BytesIO()
            doc.save(buf)
            buf.seek(0)
            
            self.send_response(200)
            self._cors()
            self.send_header('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')
            self.send_header('Content-Disposition', 'attachment; filename="gemini-chat.docx"')
            self.end_headers()
            self.wfile.write(buf.getvalue())
            
        except Exception as e:
            traceback.print_exc()
            self.send_response(500)
            self._cors()
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({'error': str(e)}).encode())

    def _process(self, doc, content):
        lines = content.split('\n')
        i = 0
        while i < len(lines):
            line = lines[i]
            
            img = re.search(r'!\[([^\]]*)\]\(([^\)]+)\)', line)
            if img:
                self._img(doc, img.group(2), img.group(1))
                i += 1
                continue
            
            if line.strip().startswith('```'):
                code = []
                i += 1
                while i < len(lines) and not lines[i].strip().startswith('```'):
                    code.append(lines[i])
                    i += 1
                self._code(doc, '\n'.join(code))
                i += 1
                continue
            
            if '|' in line and line.strip().startswith('|'):
                tlines = []
                while i < len(lines) and '|' in lines[i]:
                    if '---' not in lines[i]:
                        tlines.append(lines[i])
                    i += 1
                if tlines:
                    self._table(doc, tlines)
                continue
            
            bm = re.match(r'^\s*\$\$(.+?)\$\$\s*$', line)
            if bm:
                add_block_formula(doc, bm.group(1).strip())
                i += 1
                continue
            
            if line.strip() == '$$':
                fl = []
                i += 1
                while i < len(lines) and lines[i].strip() != '$$':
                    fl.append(lines[i])
                    i += 1
                latex = ' '.join(fl).strip()
                if latex:
                    add_block_formula(doc, latex)
                i += 1
                continue
            
            if line.strip():
                self._text_math(doc, line)
            else:
                doc.add_paragraph()
            i += 1

    def _text_math(self, doc, text):
        parts = re.split(r'(?<!\$)\$(?!\$)(.+?)(?<!\$)\$(?!\$)', text)
        if len(parts) <= 1:
            p = doc.add_paragraph()
            self._fmt(p, text)
            return
        p = doc.add_paragraph()
        for idx, part in enumerate(parts):
            if idx % 2 == 0:
                if part:
                    self._fmt(p, part)
            else:
                insert_math(p, part.strip())

    def _fmt(self, para, text):
        bparts = re.split(r'\*\*(.+?)\*\*', text)
        for i, bp in enumerate(bparts):
            if i % 2 == 0:
                iparts = re.split(r'\*(.+?)\*', bp)
                for j, ip in enumerate(iparts):
                    if j % 2 == 0:
                        if ip: para.add_run(ip)
                    else:
                        r = para.add_run(ip)
                        r.italic = True
            else:
                r = para.add_run(bp)
                r.bold = True

    def _code(self, doc, code):
        p = doc.add_paragraph()
        r = p.add_run(code)
        r.font.name = 'Courier New'
        r.font.size = Pt(10)
        s = OxmlElement('w:shd')
        s.set(qn('w:fill'), 'F5F5F5')
        p._element.get_or_add_pPr().append(s)

    def _img(self, doc, src, alt=''):
        try:
            import urllib.request, base64
            if src.startswith('data:image'):
                b64 = src.split('base64,')[1]
                stream = io.BytesIO(base64.b64decode(b64))
            else:
                req = urllib.request.Request(src, headers={'User-Agent': 'Mozilla/5.0'})
                with urllib.request.urlopen(req, timeout=30) as resp:
                    stream = io.BytesIO(resp.read())
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.add_run().add_picture(stream, width=Inches(5.0))
        except:
            p = doc.add_paragraph()
            r = p.add_run(f'[Image: {alt}]')
            r.italic = True

    def _table(self, doc, tlines):
        rows = []
        for l in tlines:
            cells = [c.strip() for c in l.split('|') if c.strip()]
            if cells: rows.append(cells)
        if not rows: return
        mc = max(len(r) for r in rows)
        t = doc.add_table(rows=len(rows), cols=mc)
        t.style = 'Table Grid'
        for i, rd in enumerate(rows):
            for j, ct in enumerate(rd):
                if j < mc:
                    r = t.rows[i].cells[j].paragraphs[0].add_run(ct)
                    r.font.size = Pt(11)
                    if i == 0: r.bold = True

    def _date(self):
        from datetime import datetime
        return datetime.now().strftime('%d.%m.%Y %H:%M')
