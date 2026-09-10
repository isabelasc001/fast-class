from pathlib import Path
import sys

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor, Twips


SKILL_ROOT = Path('/Users/jaissonmonteiro/.codex/plugins/cache/openai-primary-runtime/documents/26.826.12353/skills/documents')
sys.path.insert(0, str(SKILL_ROOT / 'scripts'))
from table_geometry import apply_table_geometry


OUT = Path('/Users/jaissonmonteiro/Documents/Projeto Gerador de Aulas/Guia de Design System — Gerador de Aulas.docx')

PAGE_WIDTH_DXA = 9360
TABLE_INDENT_DXA = 120
MARGINS_DXA = {'top': 100, 'bottom': 100, 'start': 140, 'end': 140}

COLORS = {
    'terracotta-700': '9A4F3D',
    'terracotta-600': 'B85C47',
    'terracotta-500': 'C96F57',
    'orange-500': 'D9823B',
    'orange-400': 'E6A05F',
    'brown-900': '3B2923',
    'brown-700': '5B4036',
    'brown-500': '806158',
    'off-white': 'FAF7F2',
    'cream': 'F3EDE4',
    'warm-gray-100': 'E9E4DE',
    'warm-gray-300': 'CFC7BF',
    'gray-500': '77716C',
    'gray-700': '4A4643',
    'black': '211D1B',
    'white': 'FFFFFF',
    'success': '5F7F67',
    'warning': 'C38A3A',
    'error': 'B64B45',
    'info': '6B7C86',
}


def rgb(value):
    return RGBColor.from_string(value.replace('#', ''))


def set_run_font(run, name='Calibri', size=None, color=None, bold=None, italic=None):
    run.font.name = name
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.rFonts
    if rfonts is None:
        rfonts = OxmlElement('w:rFonts')
        rpr.append(rfonts)
    rfonts.set(qn('w:ascii'), name)
    rfonts.set(qn('w:hAnsi'), name)
    rfonts.set(qn('w:eastAsia'), name)
    if size is not None:
        run.font.size = Pt(size)
    if color is not None:
        run.font.color.rgb = rgb(color)
    if bold is not None:
        run.bold = bold
    if italic is not None:
        run.italic = italic
    return run


def set_style_font(style, name='Calibri', size=11, color='4A4643', bold=False, italic=False):
    style.font.name = name
    rpr = style._element.get_or_add_rPr()
    rfonts = rpr.rFonts
    if rfonts is None:
        rfonts = OxmlElement('w:rFonts')
        rpr.append(rfonts)
    for attr in ('ascii', 'hAnsi', 'eastAsia'):
        rfonts.set(qn(f'w:{attr}'), name)
    style.font.size = Pt(size)
    style.font.color.rgb = rgb(color)
    style.font.bold = bold
    style.font.italic = italic


def set_cell_shading(cell, fill):
    tcpr = cell._tc.get_or_add_tcPr()
    shd = tcpr.find(qn('w:shd'))
    if shd is None:
        shd = OxmlElement('w:shd')
        tcpr.append(shd)
    shd.set(qn('w:fill'), fill.replace('#', ''))
    shd.set(qn('w:val'), 'clear')


def set_cell_border(cell, color='D8D1CB', size='6', sides=('top', 'left', 'bottom', 'right', 'insideH', 'insideV')):
    tcpr = cell._tc.get_or_add_tcPr()
    borders = tcpr.first_child_found_in('w:tcBorders')
    if borders is None:
        borders = OxmlElement('w:tcBorders')
        tcpr.append(borders)
    for side in sides:
        tag = f'w:{side}'
        element = borders.find(qn(tag))
        if element is None:
            element = OxmlElement(tag)
            borders.append(element)
        element.set(qn('w:val'), 'single')
        element.set(qn('w:sz'), size)
        element.set(qn('w:space'), '0')
        element.set(qn('w:color'), color.replace('#', ''))


def set_paragraph_shading(paragraph, fill):
    ppr = paragraph._p.get_or_add_pPr()
    shd = ppr.find(qn('w:shd'))
    if shd is None:
        shd = OxmlElement('w:shd')
        ppr.append(shd)
    shd.set(qn('w:fill'), fill.replace('#', ''))
    shd.set(qn('w:val'), 'clear')


def set_paragraph_border(paragraph, side='left', color='B85C47', size='18', space='8'):
    ppr = paragraph._p.get_or_add_pPr()
    pbdr = ppr.find(qn('w:pBdr'))
    if pbdr is None:
        pbdr = OxmlElement('w:pBdr')
        ppr.append(pbdr)
    border = pbdr.find(qn(f'w:{side}'))
    if border is None:
        border = OxmlElement(f'w:{side}')
        pbdr.append(border)
    border.set(qn('w:val'), 'single')
    border.set(qn('w:sz'), size)
    border.set(qn('w:space'), space)
    border.set(qn('w:color'), color.replace('#', ''))


def set_paragraph_bottom_rule(paragraph, color='D8D1CB', size='6', space='1'):
    set_paragraph_border(paragraph, side='bottom', color=color, size=size, space=space)


def set_cell_text(cell, text, *, color='4A4643', bold=False, size=9.5, align=WD_ALIGN_PARAGRAPH.LEFT, font='Calibri'):
    cell.text = ''
    p = cell.paragraphs[0]
    p.alignment = align
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.12
    p.paragraph_format.keep_together = True
    run = p.add_run(text)
    set_run_font(run, name=font, size=size, color=color, bold=bold)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    return p


def mark_header_row(row):
    trpr = row._tr.get_or_add_trPr()
    hdr = trpr.find(qn('w:tblHeader'))
    if hdr is None:
        hdr = OxmlElement('w:tblHeader')
        trpr.append(hdr)
    hdr.set(qn('w:val'), 'true')


def set_keep_with_next(paragraph):
    paragraph.paragraph_format.keep_with_next = True


def setup_styles(doc):
    normal = doc.styles['Normal']
    set_style_font(normal, size=10.5, color=COLORS['gray-700'])
    normal.paragraph_format.space_before = Pt(0)
    normal.paragraph_format.space_after = Pt(6)
    normal.paragraph_format.line_spacing = 1.25

    for name, size, color, before, after in [
        ('Heading 1', 20, COLORS['brown-900'], 18, 8),
        ('Heading 2', 14, COLORS['terracotta-700'], 14, 6),
        ('Heading 3', 11.5, COLORS['brown-700'], 10, 4),
    ]:
        style = doc.styles[name]
        set_style_font(style, size=size, color=color, bold=True)
        style.paragraph_format.space_before = Pt(before)
        style.paragraph_format.space_after = Pt(after)
        style.paragraph_format.line_spacing = 1.1
        style.paragraph_format.keep_with_next = True

    styles = {
        'Guide Kicker': ('Calibri', 9, COLORS['terracotta-700'], True, False),
        'Guide Title': ('Georgia', 30, COLORS['brown-900'], True, False),
        'Guide Subtitle': ('Calibri', 13, COLORS['brown-700'], False, False),
        'Lead': ('Calibri', 12, COLORS['brown-700'], False, False),
        'Small': ('Calibri', 9, COLORS['gray-500'], False, False),
        'Caption': ('Calibri', 9, COLORS['gray-500'], False, True),
        'Code Block': ('Courier New', 8.6, COLORS['brown-900'], False, False),
        'Table Header': ('Calibri', 9.4, COLORS['brown-900'], True, False),
        'Table Body': ('Calibri', 9.2, COLORS['gray-700'], False, False),
        'Callout': ('Calibri', 10.5, COLORS['gray-700'], False, False),
    }
    for name, (font, size, color, bold, italic) in styles.items():
        if name in doc.styles:
            style = doc.styles[name]
        else:
            style = doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
        set_style_font(style, name=font, size=size, color=color, bold=bold, italic=italic)
        style.paragraph_format.space_before = Pt(0)
        style.paragraph_format.space_after = Pt(6)
        style.paragraph_format.line_spacing = 1.15 if name in ('Small', 'Caption') else 1.25

    for name in ('List Bullet', 'List Number'):
        style = doc.styles[name]
        set_style_font(style, size=10.5, color=COLORS['gray-700'])
        style.paragraph_format.left_indent = Inches(0.375)
        style.paragraph_format.first_line_indent = Inches(-0.188)
        style.paragraph_format.space_after = Pt(4)
        style.paragraph_format.line_spacing = 1.25


def configure_page(doc):
    section = doc.sections[0]
    section.page_width = Inches(8.5)
    section.page_height = Inches(11)
    section.top_margin = Inches(0.78)
    section.bottom_margin = Inches(0.72)
    section.left_margin = Inches(0.85)
    section.right_margin = Inches(0.85)
    section.header_distance = Inches(0.35)
    section.footer_distance = Inches(0.35)

    header = section.header
    header.is_linked_to_previous = False
    hp = header.paragraphs[0]
    hp.text = ''
    hp.alignment = WD_ALIGN_PARAGRAPH.LEFT
    hp.paragraph_format.space_after = Pt(0)
    left = hp.add_run('GERADOR DE AULAS')
    set_run_font(left, size=8.2, color=COLORS['gray-500'], bold=True)
    tab = hp.add_run('\t')
    set_run_font(tab, size=8.2, color=COLORS['gray-500'])
    right = hp.add_run('DESIGN SYSTEM  /  V0.1')
    set_run_font(right, size=8.2, color=COLORS['gray-500'], bold=True)
    hp.paragraph_format.tab_stops.add_tab_stop(Inches(6.65))

    footer = section.footer
    footer.is_linked_to_previous = False
    fp = footer.paragraphs[0]
    fp.text = ''
    fp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    fp.paragraph_format.space_before = Pt(4)
    set_paragraph_top_border(fp, COLORS['warm-gray-100'], '4')
    label = fp.add_run('Base visual e comportamental para o MVP  |  ')
    set_run_font(label, size=8.2, color=COLORS['gray-500'])
    add_page_field(fp, color=COLORS['gray-500'])


def set_paragraph_top_border(paragraph, color='D8D1CB', size='4'):
    ppr = paragraph._p.get_or_add_pPr()
    pbdr = ppr.find(qn('w:pBdr'))
    if pbdr is None:
        pbdr = OxmlElement('w:pBdr')
        ppr.append(pbdr)
    top = pbdr.find(qn('w:top'))
    if top is None:
        top = OxmlElement('w:top')
        pbdr.append(top)
    top.set(qn('w:val'), 'single')
    top.set(qn('w:sz'), size)
    top.set(qn('w:space'), '4')
    top.set(qn('w:color'), color.replace('#', ''))


def add_page_field(paragraph, color='77716C'):
    run = paragraph.add_run()
    set_run_font(run, size=8.2, color=color)
    fld_char_begin = OxmlElement('w:fldChar')
    fld_char_begin.set(qn('w:fldCharType'), 'begin')
    instr = OxmlElement('w:instrText')
    instr.set(qn('xml:space'), 'preserve')
    instr.text = ' PAGE '
    fld_char_end = OxmlElement('w:fldChar')
    fld_char_end.set(qn('w:fldCharType'), 'end')
    run._r.append(fld_char_begin)
    run._r.append(instr)
    run._r.append(fld_char_end)


def add_para(doc, text='', style='Normal', *, align=None, before=None, after=None, keep=False):
    p = doc.add_paragraph(style=style)
    if text:
        p.add_run(text)
    if align is not None:
        p.alignment = align
    if before is not None:
        p.paragraph_format.space_before = Pt(before)
    if after is not None:
        p.paragraph_format.space_after = Pt(after)
    if keep:
        set_keep_with_next(p)
    return p


def add_rich_para(doc, parts, style='Normal', *, align=None, before=None, after=None):
    p = doc.add_paragraph(style=style)
    if align is not None:
        p.alignment = align
    if before is not None:
        p.paragraph_format.space_before = Pt(before)
    if after is not None:
        p.paragraph_format.space_after = Pt(after)
    for item in parts:
        if isinstance(item, str):
            text, kwargs = item, {}
        else:
            text, kwargs = item
        run = p.add_run(text)
        set_run_font(run, name=kwargs.get('font', 'Calibri'), size=kwargs.get('size', None), color=kwargs.get('color', COLORS['gray-700']), bold=kwargs.get('bold', False), italic=kwargs.get('italic', False))
    return p


def add_kicker(doc, text):
    p = add_para(doc, '', style='Guide Kicker', after=8)
    p.paragraph_format.keep_with_next = True
    run = p.add_run(text.upper())
    set_run_font(run, size=9, color=COLORS['terracotta-700'], bold=True)
    return p


def add_title(doc, text, subtitle=None):
    p = add_para(doc, '', style='Guide Title', after=8, keep=True)
    p.paragraph_format.line_spacing = 1.0
    r = p.add_run(text)
    set_run_font(r, name='Georgia', size=30, color=COLORS['brown-900'], bold=True)
    if subtitle:
        sp = add_para(doc, '', style='Guide Subtitle', after=16, keep=True)
        r = sp.add_run(subtitle)
        set_run_font(r, size=13, color=COLORS['brown-700'])
    return p


def add_callout(doc, title, text, *, fill='F3EDE4', accent='B85C47'):
    p = doc.add_paragraph(style='Callout')
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(12)
    p.paragraph_format.left_indent = Inches(0.14)
    p.paragraph_format.right_indent = Inches(0.10)
    p.paragraph_format.line_spacing = 1.2
    set_paragraph_shading(p, fill)
    set_paragraph_border(p, side='left', color=accent, size='22', space='8')
    t = p.add_run(title + '  ')
    set_run_font(t, size=10.5, color=accent, bold=True)
    b = p.add_run(text)
    set_run_font(b, size=10.5, color=COLORS['brown-700'])
    return p


def add_rule(doc, before=0, after=10):
    p = add_para(doc, '', before=before, after=after)
    set_paragraph_bottom_rule(p, color=COLORS['warm-gray-300'], size='6', space='1')
    return p


def add_code_block(doc, code, *, label=None):
    if label:
        add_para(doc, label, style='Small', before=2, after=3, keep=True)
    p = doc.add_paragraph(style='Code Block')
    p.paragraph_format.left_indent = Inches(0.12)
    p.paragraph_format.right_indent = Inches(0.12)
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(10)
    p.paragraph_format.line_spacing = 1.0
    set_paragraph_shading(p, 'F5EFE8')
    set_paragraph_border(p, side='left', color=COLORS['orange-500'], size='12', space='8')
    for i, line in enumerate(code.splitlines()):
        run = p.add_run(line)
        set_run_font(run, name='Courier New', size=8.6, color=COLORS['brown-900'])
        if i < len(code.splitlines()) - 1:
            run.add_break()
    return p


def add_table(doc, headers, rows, widths, *, header_fill='E8EEF5', font_size=9.2, header_color='3B2923', swatch_col=None, swatch_colors=None):
    table = doc.add_table(rows=1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    table.autofit = False
    for idx, header in enumerate(headers):
        cell = table.rows[0].cells[idx]
        set_cell_text(cell, header, color=header_color, bold=True, size=9.3)
        set_cell_shading(cell, header_fill)
        set_cell_border(cell, color='D3DDE5', size='6')
    mark_header_row(table.rows[0])
    for row_idx, row in enumerate(rows):
        cells = table.add_row().cells
        for idx, value in enumerate(row):
            if swatch_col is not None and idx == swatch_col:
                set_cell_text(cells[idx], '', align=WD_ALIGN_PARAGRAPH.CENTER, size=9.0)
                fill = (swatch_colors or [])[row_idx]
                set_cell_shading(cells[idx], fill)
                set_cell_border(cells[idx], color='D8D1CB', size='6')
                continue
            align = WD_ALIGN_PARAGRAPH.CENTER if idx == 0 and len(headers) >= 3 and len(str(value)) <= 16 else WD_ALIGN_PARAGRAPH.LEFT
            set_cell_text(cells[idx], str(value), color=COLORS['gray-700'], size=font_size, align=align)
            set_cell_border(cells[idx], color='D8D1CB', size='6')
    apply_table_geometry(table, widths, table_width_dxa=PAGE_WIDTH_DXA, indent_dxa=TABLE_INDENT_DXA, cell_margins_dxa=MARGINS_DXA)
    after = doc.add_paragraph(style='Caption')
    after.paragraph_format.space_before = Pt(3)
    after.paragraph_format.space_after = Pt(8)
    return table


def add_bullet(doc, text):
    p = doc.add_paragraph(style='List Bullet')
    p.add_run(text)
    return p


def add_number(doc, text):
    p = doc.add_paragraph(style='List Number')
    p.add_run(text)
    return p


def add_section(doc, level, text):
    p = doc.add_paragraph(text, style=f'Heading {level}')
    return p


def add_text_cell_table(doc, rows, widths=(2700, 6660), fill='F3EDE4'):
    table = doc.add_table(rows=0, cols=2)
    for left, right in rows:
        cells = table.add_row().cells
        set_cell_text(cells[0], left, color=COLORS['terracotta-700'], bold=True, size=9.3)
        set_cell_text(cells[1], right, color=COLORS['gray-700'], size=9.3)
        set_cell_shading(cells[0], fill)
        set_cell_border(cells[0], color='E6D5CA', size='6')
        set_cell_border(cells[1], color='E6D5CA', size='6')
    apply_table_geometry(table, list(widths), table_width_dxa=PAGE_WIDTH_DXA, indent_dxa=TABLE_INDENT_DXA, cell_margins_dxa=MARGINS_DXA)
    doc.add_paragraph(style='Caption')
    return table


def page_break(doc):
    doc.add_page_break()


def build():
    doc = Document()
    setup_styles(doc)
    configure_page(doc)
    props = doc.core_properties
    props.title = 'Guia de Design System — Gerador de Aulas'
    props.subject = 'Base visual e comportamental para o MVP'
    props.author = 'Gerador de Aulas'
    props.keywords = 'design system, gerador de aulas, MVP, UI, UX'

    # Cover
    add_kicker(doc, 'Gerador de Aulas  /  Produto')
    add_title(doc, 'Guia de Design System', 'Base visual e comportamental para o MVP')
    add_para(doc, 'Uma referência única para transformar a proposta de produto em uma experiência acolhedora, estruturada e confiável para professores.', style='Lead', after=16)
    add_callout(doc, 'DECISÃO CENTRAL', 'O produto é um workspace de planejamento de aulas com IA: a interface ajuda o professor a formular, revisar e aplicar uma aula; não entrega apenas um campo de prompt.', fill='F3EDE4', accent='B85C47')
    add_rule(doc, before=8, after=12)
    add_text_cell_table(doc, [
        ('Versão', '0.1  |  31 de agosto de 2026'),
        ('Base analisada', 'Guia de Desenvolvimento — AI Lesson Generator + Proposta de MVP - Gerador de Aulas'),
        ('Público primário', 'Professores da educação básica brasileira preparando aulas e avaliações'),
        ('Superfície inicial', 'Aplicação web responsiva, desktop-first, com suporte consistente a tablet'),
    ], widths=(2200, 7160), fill='FAF7F2')
    add_para(doc, 'Status: referência inicial de design. Tokens e comportamentos desta versão são a base normativa; nome, slogan, idioma final e recursos além do MVP continuam hipóteses a validar.', style='Small', before=12, after=0)

    # Page 2: reading guide
    page_break(doc)
    add_kicker(doc, 'Como usar este guia')
    add_section(doc, 1, 'De intenção de produto a decisão de interface')
    add_para(doc, 'Os documentos analisados têm papéis complementares. O guia de desenvolvimento traz a linguagem visual e os padrões de interação; a proposta de MVP define o contexto pedagógico, o fluxo de geração, os modos Fast e Criativo e os limites técnicos. Este documento consolida os dois em regras aplicáveis ao dia a dia do produto.', style='Lead')
    add_section(doc, 2, 'O que é normativo e o que é hipótese')
    add_text_cell_table(doc, [
        ('Normativo', 'Tokens de cor, escala tipográfica, espaçamento, estados, acessibilidade, estrutura de componentes e regras de interação.'),
        ('Recomendado', 'Arquitetura de telas, prioridades do dashboard, comportamento do editor, uso de ilustrações e microcopy de referência.'),
        ('Hipótese', 'Nome da marca, slogan, textos em inglês, dark mode, modo Criativo completo, compartilhamento e recursos escolares.'),
    ], widths=(2100, 7260), fill='F3EDE4')
    add_section(doc, 2, 'Princípios do sistema')
    add_bullet(doc, 'Professor no controle: a IA sugere, o professor revisa, edita, salva ou descarta.')
    add_bullet(doc, 'Estrutura antes da mágica: perguntas bem agrupadas reduzem o trabalho de formular prompts.')
    add_bullet(doc, 'Clareza editorial: títulos, etapas e objetivos devem ser lidos como um plano utilizável, não como uma resposta de chatbot.')
    add_bullet(doc, 'Calma por padrão: superfícies claras, sombras discretas e movimento curto preservam foco durante horas de uso.')
    add_bullet(doc, 'Rastreabilidade visível: objetivos, habilidades, fontes e materiais de origem devem aparecer quando forem relevantes para a decisão.')
    add_callout(doc, 'REGRA DE OURO', 'Ideia do professor → perguntas certas → aula estruturada → revisão → refinamento assistido → salvar, exportar e ensinar.', fill='FFF8EA', accent='D9823B')

    # Page 3: colors
    page_break(doc)
    add_kicker(doc, 'Foundations  /  Cor')
    add_section(doc, 1, 'Paleta de cores')
    add_para(doc, 'A paleta combina terracota, laranja, creme, marrom e cinzas quentes. Ela deve comunicar criatividade e acolhimento sem transformar a interface em uma superfície colorida demais.', style='Lead')
    palette_rows = [
        ('terracotta-700', '', '#9A4F3D', 'Botões principais, links e texto sobre superfícies claras'),
        ('terracotta-600', '', '#B85C47', 'Ação primária e identidade'),
        ('terracotta-500', '', '#C96F57', 'Hover, destaque e dark mode'),
        ('orange-500', '', '#D9823B', 'Ações secundárias, progresso e foco editorial'),
        ('orange-400', '', '#E6A05F', 'Ilustrações e detalhes; não usar como texto pequeno'),
        ('brown-900', '', '#3B2923', 'Títulos e conteúdo de alta prioridade'),
        ('brown-700', '', '#5B4036', 'Texto secundário e controles'),
        ('brown-500', '', '#806158', 'Texto terciário e metadados'),
    ]
    add_table(doc, ['Token', 'Amostra', 'Valor', 'Uso'], palette_rows, [2200, 900, 1500, 4760], header_fill='F3EDE4', swatch_col=1, swatch_colors=[COLORS[r[0]] for r in palette_rows])
    add_section(doc, 2, 'Neutros e estados semânticos')
    neutral_rows = [
        ('off-white', '', '#FAF7F2', 'Background principal'),
        ('cream', '', '#F3EDE4', 'Sidebar, seções e superfícies de apoio'),
        ('white', '', '#FFFFFF', 'Cards, inputs e superfícies elevadas'),
        ('warm-gray-100', '', '#E9E4DE', 'Divisores e bordas suaves'),
        ('warm-gray-300', '', '#CFC7BF', 'Bordas de inputs'),
        ('gray-700', '', '#4A4643', 'Texto de leitura'),
        ('gray-500', '', '#77716C', 'Texto secundário'),
        ('success / warning / error / info', '', '#5F7F67 / #C38A3A / #B64B45 / #6B7C86', 'Estados e mensagens'),
    ]
    add_table(doc, ['Token', 'Amostra', 'Valor', 'Uso'], neutral_rows, [2200, 900, 2500, 3760], header_fill='F3EDE4', swatch_col=1, swatch_colors=[COLORS['off-white'], COLORS['cream'], COLORS['white'], COLORS['warm-gray-100'], COLORS['warm-gray-300'], COLORS['gray-700'], COLORS['gray-500'], COLORS['success']])
    add_callout(doc, 'PROPORÇÃO 70 / 20 / 10', 'Use aproximadamente 70% de off-white e neutros, 20% de marrons e cinzas e 10% de terracota e laranja. Terracota é personalidade e ação, não o background de toda a aplicação.', fill='F3EDE4', accent='B85C47')

    # Page 4: typography and spacing
    page_break(doc)
    add_kicker(doc, 'Foundations  /  Tipografia e ritmo')
    add_section(doc, 1, 'Tipografia')
    add_para(doc, 'A recomendação de produto é DM Serif Display para títulos e Inter para a interface. O contraste entre serifada e sans-serif cria a sensação editorial sem sacrificar leitura, densidade ou precisão nos controles.', style='Lead')
    add_text_cell_table(doc, [
        ('DM Serif Display', 'Títulos de páginas, hero, nomes de aulas e momentos de marca. Usar com parcimônia e sem pesos extremos.'),
        ('Inter', 'Botões, menus, formulários, labels, tabelas, mensagens da IA e todo texto operacional.'),
        ('Fallback', 'Georgia para títulos e Arial ou system-ui para interface quando as fontes não estiverem disponíveis.'),
    ], widths=(2600, 6760), fill='F3EDE4')
    add_section(doc, 2, 'Escala tipográfica da interface')
    type_rows = [
        ('Display', '48 / 56', 'Hero de landing ou momentos de marca'),
        ('H1', '36 / 44', 'Título principal de página'),
        ('H2', '28 / 36', 'Seção ou título de resultado'),
        ('H3', '22 / 30', 'Subseção, nome de atividade ou painel'),
        ('Body Large', '18 / 28', 'Lead e texto introdutório'),
        ('Body', '16 / 24', 'Leitura principal e mensagens'),
        ('Body Small', '14 / 20', 'Labels, controles, metadados'),
        ('Caption', '12 / 16', 'Ajuda, timestamp e informação auxiliar'),
    ]
    add_table(doc, ['Estilo', 'Tamanho / linha', 'Uso'], type_rows, [1900, 1900, 5560], header_fill='F3EDE4', font_size=9.5)
    add_para(doc, 'Pesos permitidos: 400 regular, 500 medium, 600 semibold e 700 bold. Priorize contraste de tamanho, espaço e posição antes de aumentar peso.', style='Small', before=0, after=8)
    add_section(doc, 2, 'Espaçamento e forma')
    add_text_cell_table(doc, [
        ('Base de spacing', 'Escala de 4 px: 4, 8, 12, 16, 20, 24, 32, 40, 48, 64.'),
        ('Radius', 'xs 4 px · sm 8 px · md 12 px · lg 16 px · xl 24 px. Inputs e botões: 8 px; cards: 16 px; modal: 20 px.'),
        ('Borda', '1 px, warm-gray-300 em inputs; warm-gray-100 em divisores; evitar grades pesadas.'),
        ('Sombra', 'sm 0 1px 3px rgba(59,41,35,.06) · md 0 4px 12px rgba(59,41,35,.08) · lg 0 12px 30px rgba(59,41,35,.10).'),
    ], widths=(2200, 7160), fill='FAF7F2')

    # Page 5: layout
    page_break(doc)
    add_kicker(doc, 'Foundations  /  Layout')
    add_section(doc, 1, 'Estrutura da aplicação')
    add_para(doc, 'A shell principal deve manter navegação persistente e conteúdo focado. A sidebar orienta; o header confirma contexto; o PageContainer controla a largura de leitura.', style='Lead')
    add_code_block(doc, '┌──────────────────────────────────────────────────────────────┐\n│ Logo / workspace                              Profile / Menu │\n├───────────────┬──────────────────────────────────────────────┤\n│ Dashboard     │                                              │\n│ My Lessons    │                 Main Content                 │\n│ Templates     │                                              │\n│ AI Generator  │                                              │\n│ Settings      │                                              │\n└───────────────┴──────────────────────────────────────────────┘', label='Blueprint da shell')
    layout_rows = [
        ('Desktop', 'Sidebar fixa de 240 px · conteúdo com max-width de 1200 px · ações importantes visíveis sem rolagem quando possível.'),
        ('Tablet', 'Sidebar recolhível · manter duas colunas somente quando a leitura continuar confortável.'),
        ('Mobile', 'Sidebar vira drawer ou navegação inferior · AI Assistant vira bottom sheet · editor usa tabs Lesson / AI Assistant.'),
        ('Grid', 'Usar spacing de 4 px e gaps consistentes de 16 ou 24 px; não criar uma cardização para cada parágrafo.'),
    ]
    add_table(doc, ['Contexto', 'Regra'], layout_rows, [1800, 7560], header_fill='F3EDE4', font_size=9.5)
    add_section(doc, 2, 'Dashboard')
    add_para(doc, 'A primeira tela deve responder rapidamente: “o que posso fazer agora?”. O hero apresenta o próximo passo; Recent Lessons dá continuidade; o vazio conduz para Create lesson.', style='Normal')
    add_bullet(doc, 'Hero: saudação curta, contexto e uma ação primária “Create lesson”.')
    add_bullet(doc, 'Recent Lessons: cards somente para unidades lógicas, com disciplina, tema, duração e última edição.')
    add_bullet(doc, 'Nunca esconder a ação principal dentro de uma grade de cards ou em um menu secundário.')
    add_callout(doc, 'SINAL DE QUALIDADE', 'Se o dashboard parece uma coleção de cards, remova contêineres até que a hierarquia de tarefas volte a ficar evidente.', fill='FFF8EA', accent='D9823B')

    # Page 6: component principles
    page_break(doc)
    add_kicker(doc, 'Componentes  /  Base')
    add_section(doc, 1, 'Componentes e contratos')
    add_para(doc, 'Componentes devem ser pequenos, previsíveis e compostos. Cada componente interativo precisa declarar visualmente seus estados e expor labels acessíveis. O sistema deve permitir que o fluxo Fast seja enxuto e que o Criativo cresça sem criar uma interface nova.', style='Lead')
    component_rows = [
        ('Button', 'primary, secondary, ghost, danger', 'idle, hover, focus, disabled, loading'),
        ('Input / Textarea', 'label, hint, placeholder opcional', 'default, focus, error, disabled'),
        ('Select / Radio / Checkbox', 'opções claras e agrupadas', 'selected, hover, focus, error'),
        ('Card', 'unidade lógica: aula, arquivo ou resumo', 'default, hover quando clicável, selected'),
        ('Badge', 'status, modo ou classificação', 'semantic color + texto, nunca cor sozinha'),
        ('Dialog / Dropdown', 'ação contextual ou confirmação', 'open, focus trap, escape, destructive'),
        ('Tabs / Progress', 'editor e etapas do fluxo', 'active, completed, current, blocked'),
        ('Skeleton', 'carregamento estrutural', 'shape que antecipa o conteúdo, sem layout shift'),
    ]
    add_table(doc, ['Componente', 'Uso', 'Estados mínimos'], component_rows, [1900, 3850, 3610], header_fill='E8EEF5', font_size=9.0)
    add_section(doc, 2, 'Botões')
    add_text_cell_table(doc, [
        ('Primary', 'background terracotta-600 · texto branco · hover terracotta-700 · usar para Create lesson, Generate, Save.'),
        ('Secondary', 'background cream · texto brown-700 · usar para Export PDF, Add activity ou ações equivalentes.'),
        ('Ghost', 'fundo transparente · texto brown-700 · usar em ações auxiliares e navegação.'),
        ('Danger', 'usar vermelho apenas em exclusão e ações realmente destrutivas; sempre pedir confirmação quando houver perda.'),
    ], widths=(1900, 7460), fill='FAF7F2')
    add_section(doc, 2, 'Inputs')
    add_bullet(doc, 'Label sempre visível; placeholder é exemplo, nunca substituto do label.')
    add_bullet(doc, 'Focus: borda terracotta-600 e ring rgba(184,92,71,.15), com indicador perceptível sem depender apenas de cor.')
    add_bullet(doc, 'Error: explicar o problema e a forma de corrigir; posicionar a mensagem próxima ao campo.')

    # Page 7: AI and core flows
    page_break(doc)
    add_kicker(doc, 'Experiência central  /  AI Lesson Generator')
    add_section(doc, 1, 'O builder estrutura o pedido')
    add_para(doc, 'O fluxo principal converte intenção pedagógica em contexto suficiente para uma geração útil. A interface deve perguntar apenas o que muda a qualidade da aula e mostrar o avanço sem transformar o processo em um formulário intimidante.', style='Lead')
    builder_rows = [
        ('01', 'Disciplina e tema', 'O que ensinar?'),
        ('02', 'Turma', 'série, nível, quantidade de alunos e conhecimentos prévios'),
        ('03', 'Duração e recursos', 'tempo, conectividade, impressão e materiais disponíveis'),
        ('04', 'Objetivos', 'o que os estudantes devem conseguir fazer'),
        ('05', 'Materiais', 'texto, imagem, PDF ou DOCX próprios'),
        ('06', 'Modo e metodologia', 'Fast para velocidade; Criativo para contexto e refinamento'),
    ]
    add_table(doc, ['Etapa', 'Campo', 'Intenção'], builder_rows, [950, 2250, 6160], header_fill='F3EDE4', font_size=9.2)
    add_section(doc, 2, 'Fast e Criativo')
    mode_rows = [
        ('Fast', 'Poucas entradas, uma geração principal, pacote enxuto e tempo até a primeira saída utilizável.'),
        ('Criativo', 'Mais arquivos e referências, planejamento em etapas, alternativas, adaptações e recuperação.'),
    ]
    add_table(doc, ['Modo', 'Comportamento visual'], mode_rows, [1700, 7660], header_fill='E8EEF5', font_size=9.5)
    add_section(doc, 2, 'O momento da IA')
    add_para(doc, 'O carregamento deve ser editorial e transparente. Evite robôs, cérebros digitais e promessas vagas. Use sparkle como sinal de assistência e descreva etapas reais do processamento.', style='Normal')
    add_code_block(doc, '✦ Criando sua aula\n  Entendendo seus objetivos\n  Organizando as atividades\n  Preparando as notas do professor', label='Copy de referência para loading')
    add_callout(doc, 'FEEDBACK', 'O usuário precisa saber se a aula está sendo criada, salva, pronta para revisão ou bloqueada por um erro. A IA nunca deve parecer silenciosa.', fill='F3EDE4', accent='B85C47')

    # Page 8: output and editor
    page_break(doc)
    add_kicker(doc, 'Experiência central  /  Resultado e refinamento')
    add_section(doc, 1, 'Resultado gerado')
    add_para(doc, 'O output deve parecer um documento que o professor realmente poderia revisar, exportar e aplicar. A leitura precisa revelar título, objetivo, etapas, duração, materiais, instruções, avaliação, gabarito, fontes e adaptações.', style='Lead')
    output_rows = [
        ('Cabeçalho', 'tema, disciplina, turma, duração, modo e status de revisão'),
        ('Objetivos', 'resultado de aprendizagem e habilidades relacionadas'),
        ('Sequência', 'atividades numeradas, tempo, materiais e instruções ao professor'),
        ('Avaliação', 'questões, gabarito e vínculo com o objetivo correspondente'),
        ('Continuidade', 'sugestões de recuperação organizadas por habilidade ou erro recorrente'),
        ('Ações', 'Edit, Regenerate, Save, Export PDF e Share conforme disponibilidade'),
    ]
    add_table(doc, ['Bloco', 'Conteúdo visível'], output_rows, [2100, 7260], header_fill='F3EDE4', font_size=9.3)
    add_section(doc, 2, 'Editor + AI Assistant')
    add_para(doc, 'O editor é o ponto onde a proposta deixa de ser um gerador de texto e vira um workspace. A aula ocupa o espaço principal; o assistente atua como painel lateral e oferece ações rápidas antes da caixa livre.', style='Normal')
    add_code_block(doc, '┌──────────────────────────────┬────────────────────────────┐\n│ Lesson                       │ AI Assistant                │\n│ Introduction                 │ How can I help?             │\n│ Learning objectives          │ [ Make this simpler ]       │\n│ Activities                   │ [ Add activity ]            │\n│ Homework                     │ [ Adapt for beginners ]    │\n│                              │ [ Ask AI... ]               │\n└──────────────────────────────┴────────────────────────────┘', label='Blueprint do editor')
    add_bullet(doc, 'Ações sugeridas: simplificar, tornar mais interativa, criar exercício de fala, adaptar nível, adicionar dever de casa.')
    add_bullet(doc, 'A resposta da IA deve informar o que mudou e permitir desfazer ou comparar quando o impacto for relevante.')
    add_bullet(doc, 'No mobile, transformar as duas colunas em tabs preserva foco e reduz a disputa por espaço.')

    # Page 9: states and motion
    page_break(doc)
    add_kicker(doc, 'Comportamento  /  Estados e movimento')
    add_section(doc, 1, 'Estados que o sistema deve prever')
    state_rows = [
        ('Empty', 'Explique o próximo passo: “Your lessons start here. Create your first lesson.” + CTA.'),
        ('Loading', 'Mostre etapas conhecidas do pipeline e preserve a estrutura com skeleton quando aplicável.'),
        ('Success', 'Confirme com clareza: “Lesson created” ou “Saved”. Não interrompa o fluxo.'),
        ('Error', 'Diga o que aconteceu, se os dados foram preservados e qual ação resolve o problema.'),
        ('Blocked', 'Explique limite, quota ou arquivo inválido; ofereça alternativa quando houver.'),
        ('Unsaved', 'Indique alterações pendentes e proteja contra perda ao sair ou fechar.'),
    ]
    add_table(doc, ['Estado', 'Regra'], state_rows, [1700, 7660], header_fill='F3EDE4', font_size=9.3)
    add_section(doc, 2, 'Microinterações')
    add_text_cell_table(doc, [
        ('Duração', '150–250 ms para transições de interface; usar ease-out.'),
        ('Generate', 'idle “✦ Generate lesson” → loading “✦ Creating lesson...” → success “✓ Lesson created”.'),
        ('Save', '“Saving...” → “Saved”; usar feedback próximo ao item, sem toast persistente desnecessário.'),
        ('Motion', 'Sutil, funcional e orientado a feedback. Evitar animações decorativas constantes.'),
    ], widths=(1800, 7560), fill='FAF7F2')
    add_section(doc, 2, 'Ícones e ilustração')
    add_para(doc, 'Usar Lucide Icons como biblioteca única, com stroke de 1.8–2 px. Base recomendada: BookOpen, Sparkles, Plus, Pencil, Trash2, Settings, User, Search, Clock, GraduationCap, MessageSquare, Download, Share2 e ChevronRight.', style='Normal')
    add_para(doc, 'Ilustrações devem ser pontuais: hand-drawn editorial ou minimal line art, com livros, lápis, folhas, cadernos, café, estrelas e lâmpadas. O objetivo é dar memória ao produto, não preencher cada estado.', style='Normal')

    # Page 10: accessibility and dark mode
    page_break(doc)
    add_kicker(doc, 'Qualidade  /  Acessibilidade e temas')
    add_section(doc, 1, 'Acessibilidade desde o primeiro componente')
    add_para(doc, 'Professores podem passar horas no produto. Acessibilidade aqui não é uma camada de acabamento: ela é parte da promessa de uma ferramenta confiável e confortável.', style='Lead')
    add_bullet(doc, 'Atender contraste WCAG AA; validar combinações de texto, botão, borda e estado em cada tema.')
    add_bullet(doc, 'Garantir foco visível, navegação completa por teclado e ordem de foco previsível.')
    add_bullet(doc, 'Usar labels associados a inputs e aria-label apenas quando o nome acessível não estiver visível.')
    add_bullet(doc, 'Nunca comunicar estado apenas por cor; combinar cor, texto, ícone ou padrão.')
    add_bullet(doc, 'Manter áreas clicáveis confortáveis e mensagens de erro específicas, próximas do campo.')
    add_bullet(doc, 'Respeitar prefers-reduced-motion e oferecer alternativa sem animações não essenciais.')
    add_callout(doc, 'CRITÉRIO DE ACEITAÇÃO', 'Todo componente interativo precisa demonstrar default, hover, focus, disabled, loading e error quando esses estados fizerem sentido.', fill='F3EDE4', accent='B85C47')
    add_section(doc, 2, 'Dark mode')
    dark_rows = [
        ('Background', '#211D1B'),
        ('Surface', '#2C2421'),
        ('Card', '#352B27'),
        ('Text', '#F7F2EC'),
        ('Secondary', '#CFC3BB'),
        ('Terracotta', '#C96F57'),
    ]
    add_table(doc, ['Token', 'Valor'], dark_rows, [2600, 6760], header_fill='352B27', font_size=9.5, header_color='F7F2EC')
    add_para(doc, 'O dark mode é uma extensão de identidade, não uma obrigação do primeiro slice. Quando entrar, manter marrom muito escuro em vez de preto puro e revisar contraste, estados e imagens.', style='Small', before=0, after=6)

    # Page 11: implementation
    page_break(doc)
    add_kicker(doc, 'Implementação  /  Tokens e organização')
    add_section(doc, 1, 'Tokens no código')
    add_para(doc, 'Não espalhar hexadecimais em componentes. Tokens semânticos permitem trocar tema, corrigir contraste e evoluir a marca sem reescrever a interface.', style='Lead')
    add_code_block(doc, ':root {\n  --color-bg: #FAF7F2;\n  --color-surface: #FFFFFF;\n  --color-surface-muted: #F3EDE4;\n  --color-text: #4A4643;\n  --color-heading: #3B2923;\n  --color-text-muted: #77716C;\n  --color-border: #CFC7BF;\n  --color-primary: #B85C47;\n  --color-primary-hover: #9A4F3D;\n  --color-accent: #D9823B;\n  --color-success: #5F7F67;\n  --color-warning: #C38A3A;\n  --color-error: #B64B45;\n  --radius-sm: 8px;\n  --radius-md: 12px;\n  --radius-lg: 16px;\n  --radius-xl: 24px;\n  --shadow-sm: 0 1px 3px rgba(59,41,35,.06);\n  --shadow-md: 0 4px 12px rgba(59,41,35,.08);\n}', label='Token sheet inicial')
    add_section(doc, 2, 'Organização recomendada')
    add_code_block(doc, 'components/\n├── ui/\n│   ├── Button\n│   ├── Input\n│   ├── Select\n│   ├── Textarea\n│   ├── Checkbox\n│   ├── RadioGroup\n│   ├── Badge\n│   ├── Card\n│   ├── Dialog\n│   ├── Dropdown\n│   ├── Tabs\n│   ├── Progress\n│   └── Skeleton\n├── layout/\n│   ├── Sidebar\n│   ├── Header\n│   └── PageContainer\n├── lessons/\n│   ├── LessonCard\n│   ├── LessonBuilder\n│   ├── LessonPreview\n│   ├── LessonEditor\n│   └── LessonSection\n└── ai/\n    ├── AIAssistant\n    ├── AIGenerateButton\n    ├── AILoadingState\n    └── AIAction', label='Mapa de componentes')
    add_para(doc, 'No stack proposta (Next.js, React, TypeScript e Tailwind CSS), mapear tokens para theme utilities e manter a integração com shadcn ou outra base de UI atrás dos componentes do produto.', style='Normal')

    # Page 12: delivery checklist
    page_break(doc)
    add_kicker(doc, 'Handoff  /  Critérios de pronto')
    add_section(doc, 1, 'Checklist para design e desenvolvimento')
    add_para(doc, 'Usar esta lista em cada tela ou componente novo. O objetivo é manter consistência sem impedir exploração quando houver uma razão clara.', style='Lead')
    checklist = [
        'A tela tem uma ação principal evidente e uma hierarquia que responde ao próximo passo do professor.',
        'As cores usadas são tokens semânticos; nenhum hexadecimal está espalhado no componente.',
        'A tipografia segue a escala e usa DM Serif Display apenas para títulos expressivos.',
        'Todos os controles têm label, foco visível, feedback de erro e estado disabled quando aplicável.',
        'Loading, success, error, empty e unsaved foram considerados para o fluxo.',
        'A interface continua compreensível em 1280 px, 768 px e 375 px.',
        'A solução não cria cards, modais ou sombras sem uma unidade lógica clara.',
        'A IA é apresentada como assistência, com ações reversíveis e explicação do que está acontecendo.',
        'O texto usa idioma consistente, verbos de ação e mensagens que orientam a recuperação.',
        'O contraste foi validado e nenhum estado depende somente de cor ou animação.',
    ]
    for item in checklist:
        add_bullet(doc, item)
    add_section(doc, 2, 'Copy de referência')
    copy_rows = [
        ('Primária', 'Create lesson / Criar aula'),
        ('Geração', 'Generate lesson / Gerar aula'),
        ('Processando', 'Creating your lesson... / Criando sua aula...'),
        ('Sucesso', 'Lesson created / Aula criada'),
        ('Salvamento', 'Saving... → Saved / Salvando... → Salvo'),
        ('Vazio', 'Your lessons start here. / Suas aulas começam aqui.'),
        ('Erro', 'We could not create this lesson. Your inputs are safe. / Não foi possível criar esta aula. Seus dados foram preservados.'),
    ]
    add_table(doc, ['Momento', 'Texto'], copy_rows, [1800, 7560], header_fill='F3EDE4', font_size=9.2)
    add_callout(doc, 'IDIOMA', 'Os arquivos-base usam exemplos em inglês, mas o MVP é direcionado à educação básica brasileira. Recomenda-se definir PT-BR como idioma de produto e manter copy em inglês apenas em protótipos ou como referência de benchmark.', fill='FFF8EA', accent='D9823B')

    # Page 13: appendix / decision log
    page_break(doc)
    add_kicker(doc, 'Apêndice  /  Registro de decisões')
    add_section(doc, 1, 'Decisões consolidadas')
    decision_rows = [
        ('Identidade', 'Acolhedora, editorial, profissional, criativa e simples.'),
        ('Cor', 'Terracota como ação e personalidade; neutros quentes como base.'),
        ('Tipografia', 'DM Serif Display + Inter como par recomendado de produto.'),
        ('IA', 'Sparkle e feedback por etapas; sem estética de robô ou cérebro digital.'),
        ('Layout', 'Workspace com sidebar, conteúdo principal e assistente contextual.'),
        ('MVP', 'Fluxo Fast primeiro; contrato preparado para o Criativo e para o ciclo de recuperação.'),
        ('Rastreabilidade', 'Objetivos, habilidades, fontes e origem do material devem ser exibidos quando sustentarem a revisão.'),
        ('Persistência', 'Histórico e arquivos devem ser opt-in, com retenção e exclusão claras.'),
    ]
    add_table(doc, ['Tema', 'Decisão'], decision_rows, [1900, 7460], header_fill='E8EEF5', font_size=9.4)
    add_section(doc, 2, 'Pontos a validar com usuários')
    for item in [
        'Qual formato de planejamento o professor considera pronto para aplicar?',
        'Quais campos do contexto da turma realmente melhoram o resultado sem aumentar fricção?',
        'O pacote completo reduz retrabalho em comparação com um plano isolado?',
        'A rastreabilidade aumenta confiança e diminui dúvidas sobre fontes e alinhamento?',
        'O modo Criativo justifica maior tempo, volume de arquivos e disposição a pagar?',
    ]:
        add_number(doc, item)
    add_section(doc, 2, 'Referências internas analisadas')
    add_bullet(doc, 'Guia de Desenvolvimento — AI Lesson Generator.docx')
    add_bullet(doc, 'Proposta de MVP - Gerador de Aulas.docx')
    add_para(doc, 'Este guia é uma consolidação de design para uso no projeto. Quando houver conflito entre estética e clareza, acessibilidade ou controle do professor, priorizar clareza, acessibilidade e controle.', style='Lead', before=12, after=0)

    doc.save(OUT)
    print(OUT)


if __name__ == '__main__':
    build()
