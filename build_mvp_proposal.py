from pathlib import Path
import sys

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor, Twips

SKILL_ROOT = Path("/Users/jaissonmonteiro/.codex/plugins/cache/openai-primary-runtime/documents/26.826.12353/skills/documents")
sys.path.insert(0, str(SKILL_ROOT / "scripts"))
from table_geometry import apply_table_geometry


OUT = Path("/Users/jaissonmonteiro/Documents/ChatGPT/gerador de aulas/gerador_aulas_mvp_draft.docx")

NAVY = RGBColor(32, 55, 72)
BLUE = RGBColor(46, 116, 181)
DARK_BLUE = RGBColor(31, 77, 120)
GRAY = RGBColor(82, 82, 82)
LIGHT_GRAY = "F4F6F9"
BLUE_GRAY = "E8EEF5"
PALE_BLUE = "F2F7FC"
GOLD = RGBColor(122, 90, 0)
WHITE = RGBColor(255, 255, 255)


def set_run_font(run, name="Calibri", size=None, color=None, bold=None, italic=None):
    run.font.name = name
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.rFonts
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    rfonts.set(qn("w:ascii"), name)
    rfonts.set(qn("w:hAnsi"), name)
    if size is not None:
        run.font.size = Pt(size)
    if color is not None:
        run.font.color.rgb = color
    if bold is not None:
        run.bold = bold
    if italic is not None:
        run.italic = italic


def set_cell_shading(cell, fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill)
    shd.set(qn("w:val"), "clear")


def set_cell_border(cell, **kwargs):
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    tc_borders = tc_pr.first_child_found_in("w:tcBorders")
    if tc_borders is None:
        tc_borders = OxmlElement("w:tcBorders")
        tc_pr.append(tc_borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        if edge in kwargs:
            tag = "w:{}".format(edge)
            element = tc_borders.find(qn(tag))
            if element is None:
                element = OxmlElement(tag)
                tc_borders.append(element)
            for key in ["val", "sz", "space", "color"]:
                if key in kwargs[edge]:
                    element.set(qn("w:{}".format(key)), str(kwargs[edge][key]))


def set_cell_text(cell, text, *, bold=False, color=GRAY, size=9.4, align=WD_ALIGN_PARAGRAPH.LEFT):
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = align
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.08
    r = p.add_run(text)
    set_run_font(r, size=size, color=color, bold=bold)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


def mark_header_row(row):
    tr_pr = row._tr.get_or_add_trPr()
    tbl_header = OxmlElement("w:tblHeader")
    tbl_header.set(qn("w:val"), "true")
    tr_pr.append(tbl_header)


def set_paragraph_border(paragraph, *, bottom=None, left=None):
    p = paragraph._p
    p_pr = p.get_or_add_pPr()
    p_bdr = p_pr.find(qn("w:pBdr"))
    if p_bdr is None:
        p_bdr = OxmlElement("w:pBdr")
        p_pr.append(p_bdr)
    for edge, spec in (("bottom", bottom), ("left", left)):
        if spec:
            tag = qn("w:" + edge)
            border = p_bdr.find(tag)
            if border is None:
                border = OxmlElement("w:" + edge)
                p_bdr.append(border)
            for key, value in spec.items():
                border.set(qn("w:" + key), str(value))


def set_paragraph_shading(paragraph, fill):
    p_pr = paragraph._p.get_or_add_pPr()
    shd = p_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        p_pr.append(shd)
    shd.set(qn("w:fill"), fill)
    shd.set(qn("w:val"), "clear")


def add_page_field(paragraph):
    run = paragraph.add_run()
    fld_char1 = OxmlElement("w:fldChar")
    fld_char1.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    fld_char2 = OxmlElement("w:fldChar")
    fld_char2.set(qn("w:fldCharType"), "end")
    run._r.append(fld_char1)
    run._r.append(instr)
    run._r.append(fld_char2)
    set_run_font(run, size=9, color=GRAY)


def add_hyperlink(paragraph, text, url):
    part = paragraph.part
    r_id = part.relate_to(url, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", is_external=True)
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)
    new_run = OxmlElement("w:r")
    r_pr = OxmlElement("w:rPr")
    color = OxmlElement("w:color")
    color.set(qn("w:val"), "2E74B5")
    r_pr.append(color)
    underline = OxmlElement("w:u")
    underline.set(qn("w:val"), "single")
    r_pr.append(underline)
    fonts = OxmlElement("w:rFonts")
    fonts.set(qn("w:ascii"), "Calibri")
    fonts.set(qn("w:hAnsi"), "Calibri")
    r_pr.append(fonts)
    size = OxmlElement("w:sz")
    size.set(qn("w:val"), "19")
    r_pr.append(size)
    new_run.append(r_pr)
    text_node = OxmlElement("w:t")
    text_node.text = text
    new_run.append(text_node)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)


def set_style_font(style, name="Calibri", size=11, color=GRAY, bold=None, italic=None):
    style.font.name = name
    style._element.rPr.rFonts.set(qn("w:ascii"), name)
    style._element.rPr.rFonts.set(qn("w:hAnsi"), name)
    style.font.size = Pt(size)
    style.font.color.rgb = color
    if bold is not None:
        style.font.bold = bold
    if italic is not None:
        style.font.italic = italic


def add_body(doc, text, *, bold_prefix=None, italic=False, after=8, keep=False):
    p = doc.add_paragraph(style="Normal")
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(after)
    p.paragraph_format.line_spacing = 1.333
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    if keep:
        p.paragraph_format.keep_with_next = True
    if bold_prefix and text.startswith(bold_prefix):
        r1 = p.add_run(bold_prefix)
        set_run_font(r1, size=11, color=GRAY, bold=True)
        r2 = p.add_run(text[len(bold_prefix):])
        set_run_font(r2, size=11, color=GRAY, italic=italic)
    else:
        r = p.add_run(text)
        set_run_font(r, size=11, color=GRAY, italic=italic)
    return p


def add_bullet(doc, text):
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing = 1.208
    r = p.add_run(text)
    set_run_font(r, size=11, color=GRAY)
    return p


def add_number(doc, text):
    p = doc.add_paragraph(style="List Number")
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing = 1.208
    r = p.add_run(text)
    set_run_font(r, size=11, color=GRAY)
    return p


def create_numbering_instance(doc):
    numbering = doc.part.numbering_part.element
    abstract_ids = [int(x.get(qn("w:abstractNumId"))) for x in numbering.findall(qn("w:abstractNum")) if x.get(qn("w:abstractNumId"))]
    num_ids = [int(x.get(qn("w:numId"))) for x in numbering.findall(qn("w:num")) if x.get(qn("w:numId"))]
    abstract_id = max(abstract_ids or [0]) + 1
    num_id = max(num_ids or [0]) + 1

    abstract = OxmlElement("w:abstractNum")
    abstract.set(qn("w:abstractNumId"), str(abstract_id))
    multi = OxmlElement("w:multiLevelType")
    multi.set(qn("w:val"), "singleLevel")
    abstract.append(multi)
    lvl = OxmlElement("w:lvl")
    lvl.set(qn("w:ilvl"), "0")
    start = OxmlElement("w:start")
    start.set(qn("w:val"), "1")
    lvl.append(start)
    num_fmt = OxmlElement("w:numFmt")
    num_fmt.set(qn("w:val"), "decimal")
    lvl.append(num_fmt)
    lvl_text = OxmlElement("w:lvlText")
    lvl_text.set(qn("w:val"), "%1.")
    lvl.append(lvl_text)
    lvl_jc = OxmlElement("w:lvlJc")
    lvl_jc.set(qn("w:val"), "left")
    lvl.append(lvl_jc)
    p_pr = OxmlElement("w:pPr")
    tabs = OxmlElement("w:tabs")
    tab = OxmlElement("w:tab")
    tab.set(qn("w:val"), "num")
    tab.set(qn("w:pos"), "540")
    tabs.append(tab)
    p_pr.append(tabs)
    ind = OxmlElement("w:ind")
    ind.set(qn("w:left"), "540")
    ind.set(qn("w:hanging"), "280")
    p_pr.append(ind)
    lvl.append(p_pr)
    abstract.append(lvl)
    numbering.append(abstract)

    num = OxmlElement("w:num")
    num.set(qn("w:numId"), str(num_id))
    abs_ref = OxmlElement("w:abstractNumId")
    abs_ref.set(qn("w:val"), str(abstract_id))
    num.append(abs_ref)
    numbering.append(num)
    return num_id


def add_number_with_id(doc, text, num_id):
    p = doc.add_paragraph(style="List Number")
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing = 1.208
    p_pr = p._p.get_or_add_pPr()
    num_pr = p_pr.find(qn("w:numPr"))
    if num_pr is None:
        num_pr = OxmlElement("w:numPr")
        p_pr.append(num_pr)
    ilvl = num_pr.find(qn("w:ilvl"))
    if ilvl is None:
        ilvl = OxmlElement("w:ilvl")
        num_pr.append(ilvl)
    ilvl.set(qn("w:val"), "0")
    num_ref = num_pr.find(qn("w:numId"))
    if num_ref is None:
        num_ref = OxmlElement("w:numId")
        num_pr.append(num_ref)
    num_ref.set(qn("w:val"), str(num_id))
    r = p.add_run(text)
    set_run_font(r, size=11, color=GRAY)
    return p


def add_numbered_list(doc, items):
    num_id = create_numbering_instance(doc)
    for item in items:
        add_number_with_id(doc, item, num_id)


def add_callout(doc, label, text, fill=PALE_BLUE, accent="2E74B5"):
    p = doc.add_paragraph(style="Normal")
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(12)
    p.paragraph_format.left_indent = Inches(0.12)
    p.paragraph_format.right_indent = Inches(0.08)
    p.paragraph_format.line_spacing = 1.22
    set_paragraph_shading(p, fill)
    set_paragraph_border(p, left={"val": "single", "sz": "18", "space": "8", "color": accent})
    r1 = p.add_run(label + " ")
    set_run_font(r1, size=11, color=DARK_BLUE, bold=True)
    r2 = p.add_run(text)
    set_run_font(r2, size=11, color=GRAY)
    return p


def add_heading(doc, text, level=1):
    p = doc.add_paragraph(style=f"Heading {level}")
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    set_run_font(r, size={1: 16, 2: 13, 3: 12}[level], color={1: BLUE, 2: BLUE, 3: DARK_BLUE}[level], bold=True)
    return p


def add_table(doc, headers, rows, widths, *, header_fill=BLUE_GRAY, font_size=9.4):
    table = doc.add_table(rows=1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    table.style = "Table Grid"
    hdr = table.rows[0]
    mark_header_row(hdr)
    for idx, text in enumerate(headers):
        cell = hdr.cells[idx]
        set_cell_shading(cell, header_fill)
        set_cell_text(cell, text, bold=True, color=NAVY, size=font_size)
    for row_values in rows:
        cells = table.add_row().cells
        for idx, text in enumerate(row_values):
            set_cell_text(cells[idx], text, size=font_size)
    apply_table_geometry(table, widths, table_width_dxa=9360, indent_dxa=120,
                         cell_margins_dxa={"top": 90, "bottom": 90, "start": 120, "end": 120})
    for row in table.rows:
        for cell in row.cells:
            set_cell_border(cell,
                            top={"val": "single", "sz": "4", "color": "D7DEE8"},
                            bottom={"val": "single", "sz": "4", "color": "D7DEE8"},
                            left={"val": "single", "sz": "4", "color": "D7DEE8"},
                            right={"val": "single", "sz": "4", "color": "D7DEE8"})
    spacer = doc.add_paragraph()
    spacer.paragraph_format.space_after = Pt(2)
    return table


def add_source(doc, name, url):
    p = doc.add_paragraph(style="Normal")
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing = 1.1
    r = p.add_run(name + " — ")
    set_run_font(r, size=9.5, color=GRAY, bold=True)
    add_hyperlink(p, url, url)


def configure_document(doc):
    section = doc.sections[0]
    section.page_width = Inches(8.5)
    section.page_height = Inches(11)
    section.top_margin = Inches(1)
    section.bottom_margin = Inches(1)
    section.left_margin = Inches(1)
    section.right_margin = Inches(1)
    section.header_distance = Inches(0.492)
    section.footer_distance = Inches(0.492)

    normal = doc.styles["Normal"]
    set_style_font(normal, size=11, color=GRAY)
    normal.paragraph_format.space_before = Pt(0)
    normal.paragraph_format.space_after = Pt(8)
    normal.paragraph_format.line_spacing = 1.333

    title = doc.styles["Title"]
    set_style_font(title, size=24, color=NAVY, bold=True)
    title.paragraph_format.space_before = Pt(0)
    title.paragraph_format.space_after = Pt(4)
    title.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER

    subtitle = doc.styles["Subtitle"]
    set_style_font(subtitle, size=14, color=GRAY)
    subtitle.paragraph_format.space_before = Pt(0)
    subtitle.paragraph_format.space_after = Pt(8)
    subtitle.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER

    for level, size, color, before, after in [
        (1, 16, BLUE, 18, 10),
        (2, 13, BLUE, 12, 6),
        (3, 12, DARK_BLUE, 8, 4),
    ]:
        style = doc.styles[f"Heading {level}"]
        set_style_font(style, size=size, color=color, bold=True)
        style.paragraph_format.space_before = Pt(before)
        style.paragraph_format.space_after = Pt(after)
        style.paragraph_format.line_spacing = 1.15
        style.paragraph_format.keep_with_next = True

    for style_name in ["List Bullet", "List Number"]:
        style = doc.styles[style_name]
        set_style_font(style, size=11, color=GRAY)
        style.paragraph_format.left_indent = Inches(0.375)
        style.paragraph_format.first_line_indent = Inches(-0.194)
        style.paragraph_format.space_after = Pt(4)
        style.paragraph_format.line_spacing = 1.208

    header = section.header
    hp = header.paragraphs[0]
    hp.alignment = WD_ALIGN_PARAGRAPH.LEFT
    hp.paragraph_format.space_after = Pt(0)
    hp.paragraph_format.line_spacing = 1
    r = hp.add_run("PROPOSTA DE PRODUTO")
    set_run_font(r, size=8.5, color=GRAY, bold=True)
    tab = hp.add_run("\tGERADOR DE AULAS — MVP")
    set_run_font(tab, size=8.5, color=GRAY)
    hp.paragraph_format.tab_stops.add_tab_stop(Inches(6.5))

    footer = section.footer
    fp = footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    fp.paragraph_format.space_before = Pt(0)
    fp.paragraph_format.space_after = Pt(0)
    r = fp.add_run("Gerador de Aulas — MVP  |  ")
    set_run_font(r, size=9, color=GRAY)
    add_page_field(fp)


def build():
    doc = Document()
    configure_document(doc)

    # First-page proposal centerpiece.
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(32)
    p.paragraph_format.space_after = Pt(8)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("PROPOSTA DE MVP")
    set_run_font(r, size=11, color=GOLD, bold=True)

    p = doc.add_paragraph(style="Title")
    p.paragraph_format.space_after = Pt(5)
    r = p.add_run("Plataforma de Preparação de Aulas com IA")
    set_run_font(r, size=24, color=NAVY, bold=True)

    p = doc.add_paragraph(style="Subtitle")
    r = p.add_run("Do material real da turma à sequência didática aplicável")
    set_run_font(r, size=14, color=GRAY)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(24)
    r = p.add_run("Documento de definição inicial | Versão 0.1 | 27 de agosto de 2026")
    set_run_font(r, size=9.5, color=GRAY, italic=True)

    add_callout(doc, "Direção recomendada:", "começar com um único segmento e validar um fluxo completo — planejamento, material, avaliação e recuperação — antes de expandir para todos os tipos de cursos.", fill="FFF8E8", accent="C59A27")

    add_heading(doc, "Resumo executivo", 1)
    add_body(doc, "A proposta é criar uma plataforma de apoio ao professor capaz de receber texto, imagens e documentos, compreender a intenção pedagógica e gerar materiais prontos para uso em diferentes formatos. O MVP deve ser apresentado como um copiloto de preparação, e não apenas como um gerador de textos.")
    add_body(doc, "O diferencial recomendado é combinar contexto real da turma, rastreabilidade pedagógica e continuidade entre aula, atividade, avaliação e recuperação. Assim, o produto entrega uma decisão pedagógica mais completa e útil do que uma resposta genérica produzida a partir de um único prompt.")

    add_heading(doc, "1. Contexto e oportunidade", 1)
    add_body(doc, "O mercado já possui ferramentas maduras para gerar planos de aula, apresentações, atividades e avaliações. Portanto, a proposta precisa competir pela qualidade do fluxo e pela adequação à realidade do professor, e não apenas pela quantidade de formatos oferecidos.")
    add_table(doc,
              ["Referência", "Posicionamento observado", "Implicação para o MVP"],
              [
                  ("Teachy", "Ecossistema amplo: planos, aulas, slides, avaliações, correção, acessibilidade e BNCC.", "Evitar competir como plataforma generalista desde o primeiro dia."),
                  ("Chalkie", "Aulas, séries de aulas, fichas, atividades interativas, uploads e adaptação para alunos.", "Coerência entre materiais e personalização precisam ser requisitos reais."),
                  ("Avalea", "Planos alinhados à BNCC, provas, questões e exportação de materiais.", "Banco de questões sozinho não é diferencial suficiente."),
              ],
              [1500, 4500, 3360], font_size=9.0)
    add_body(doc, "A análise acima foi feita a partir das páginas públicas consultadas em 27 de agosto de 2026. As referências completas estão ao final do documento.", italic=True, after=10)

    add_heading(doc, "2. Problema a resolver", 1)
    add_body(doc, "Professores gastam tempo reunindo materiais, adaptando conteúdos para uma turma específica, preparando atividades e tentando garantir coerência entre o que foi ensinado e o que será avaliado. Geradores genéricos reduzem parte do trabalho, mas podem produzir materiais pouco adequados ao tempo, aos recursos e ao nível real dos alunos.")
    add_callout(doc, "Problema central:", "o professor não precisa apenas de um texto pronto; precisa de um pacote pedagógico coerente, revisável e aplicável nas condições concretas da sua turma.")

    add_heading(doc, "3. Posicionamento e diferencial", 1)
    add_body(doc, "Posicionamento proposto:")
    add_callout(doc, "O produto será:", "um copiloto que transforma os materiais reais do professor e as necessidades da turma em uma sequência didática coerente, aplicável e verificável — com plano, atividade, avaliação, gabarito e sugestões de recuperação.", fill=PALE_BLUE, accent="2E74B5")
    add_heading(doc, "Pilares de diferenciação", 2)
    add_number(doc, "Contexto real da turma: considerar duração, quantidade de alunos, conhecimentos prévios, dificuldades, recursos disponíveis, conectividade, necessidade de impressão e adaptações de acessibilidade.")
    add_number(doc, "Rastreabilidade pedagógica: vincular objetivos, atividades e questões à habilidade curricular, ao material de origem e ao objetivo de aprendizagem correspondente.")
    add_number(doc, "Ciclo de aprendizagem: permitir que uma avaliação gere recomendações de reforço e novas atividades, organizadas por habilidade, dificuldade e erro recorrente.")

    add_heading(doc, "4. Público e recorte inicial", 1)
    add_body(doc, "O documento original considera professores particulares, escolas de idiomas, escolas públicas, escolas privadas, cursos técnicos e diferentes níveis. Esse escopo é amplo demais para um primeiro produto. A recomendação é escolher um segmento onde existam usuários acessíveis para entrevistas e testes.")
    add_callout(doc, "Recorte recomendado:", "professores da educação básica brasileira que precisam criar aulas e avaliações a partir dos próprios materiais, levando em conta as limitações reais da turma.")
    add_body(doc, "A BNCC deve ser usada como referência para a educação básica. Para escolas de idiomas, ensino superior e cursos técnicos, o produto deverá oferecer posteriormente currículos, competências ou referenciais específicos — não assumir que a BNCC se aplica a todos os contextos.")

    add_heading(doc, "5. Escopo do MVP", 1)
    add_table(doc,
              ["Módulo", "Entrega do MVP", "Prioridade"],
              [
                  ("Entrada de contexto", "Texto, PDF, DOCX, imagem; disciplina, série, duração, metodologia, perfil da turma e recursos disponíveis.", "Essencial"),
                  ("Geração do pacote", "Plano de aula, roteiro, atividade, avaliação, gabarito e sugestão de recuperação.", "Essencial"),
                  ("BNCC", "Indicação das habilidades relacionadas, com possibilidade de revisão manual.", "Essencial"),
                  ("Modos de uso", "Fast para geração rápida; Criativo para mais contexto, referências e refinamento.", "Essencial"),
                  ("Banco de questões", "Salvar questões e organizá-las por tema, habilidade, dificuldade e tipo.", "Essencial"),
                  ("Exportação", "PDF e DOCX inicialmente; PowerPoint e Google Slides em etapa posterior.", "Essencial"),
                  ("Histórico", "Histórico opcional e retenção configurável, com exclusão pelo usuário.", "Recomendado"),
                  ("Analytics escolar", "Indicadores de turma, coordenação e desempenho em escala.", "Fora do MVP"),
                  ("Tutor do aluno", "Chat, gamificação e acompanhamento individual do estudante.", "Fora do MVP"),
              ],
              [1700, 5350, 2310], font_size=9.0)

    add_heading(doc, "6. Experiência principal do usuário", 1)
    add_body(doc, "O fluxo deve conduzir o professor de uma intenção pedagógica até um pacote pronto para revisar e usar:")
    add_numbered_list(doc, [
        "Selecionar o segmento, disciplina, série, duração e objetivo da aula.",
        "Adicionar texto, imagens ou documentos próprios e descrever o contexto da turma.",
        "Escolher o modo Fast ou Criativo e indicar a metodologia desejada.",
        "Gerar o pacote pedagógico com referências, habilidades, atividades e avaliação.",
        "Revisar, editar, salvar itens no banco de questões e exportar o material.",
        "Após aplicar a avaliação, informar resultados para receber uma sugestão de recuperação.",
    ])

    add_heading(doc, "Fast e Criativo", 2)
    add_table(doc,
              ["Aspecto", "Fast", "Criativo"],
              [
                  ("Objetivo", "Obter um material utilizável rapidamente.", "Explorar uma solução pedagógica mais completa."),
                  ("Entradas", "Texto curto, metodologia, tipo de material e poucos arquivos.", "Vários arquivos, referências, contexto detalhado e objetivos."),
                  ("Processamento", "Uma geração com limite de tokens e poucas etapas.", "Extração, planejamento, geração, revisão e validação."),
                  ("Saída", "Material único ou pacote enxuto.", "Sequência didática com alternativas, adaptações e recuperação."),
              ],
              [1700, 3830, 3830], font_size=9.0)

    add_heading(doc, "7. Requisitos funcionais iniciais", 1)
    add_table(doc,
              ["ID", "Requisito", "Critério de aceitação"],
              [
                  ("RF-01", "Receber entradas multimodais", "Usuário consegue enviar texto, PDF, DOCX e imagem dentro dos limites definidos."),
                  ("RF-02", "Gerar pacote coerente", "Objetivos, aula, atividade e avaliação tratam o mesmo tema e nível."),
                  ("RF-03", "Relacionar à BNCC", "Sistema sugere habilidades e permite edição antes da exportação."),
                  ("RF-04", "Preservar estrutura interna", "A saída é criada em um modelo canônico antes de virar PDF ou DOCX."),
                  ("RF-05", "Salvar questões", "Usuário consegue revisar, classificar, editar e reutilizar questões."),
                  ("RF-06", "Limitar consumo", "Sistema informa limites, bloqueia excesso e registra custo por operação."),
                  ("RF-07", "Proteger dados", "Arquivos têm política de retenção, exclusão e permissões claras."),
              ],
              [1100, 3300, 4960], font_size=9.0)

    add_heading(doc, "8. Arquitetura conceitual", 1)
    add_body(doc, "A arquitetura deve tratar a aula como uma entidade estruturada, e não como um texto único. Um modelo canônico pode conter objetivos, habilidades, conteúdos, etapas, materiais, questões, gabarito, adaptações, fontes e metadados de geração.")
    add_heading(doc, "Pipeline proposto", 2)
    add_numbered_list(doc, [
        "Recepção e validação do arquivo ou texto.",
        "Extração de texto, OCR quando necessário e identificação de metadados.",
        "Divisão do conteúdo em trechos e recuperação apenas das partes relevantes.",
        "Geração de um plano estruturado com schema validado.",
        "Verificação de coerência, referências curriculares e campos obrigatórios.",
        "Renderização do modelo em PDF, DOCX e, futuramente, apresentações.",
    ])
    add_heading(doc, "Tecnologia", 2)
    add_body(doc, "TypeScript, React Native e AWS são opções viáveis, mas a decisão deve seguir o fluxo de uso. Para o MVP, uma aplicação web responsiva pode validar a proposta mais rapidamente; o aplicativo React Native pode entrar quando houver evidência de uso recorrente em dispositivos móveis.")

    add_heading(doc, "9. Contexto, custo e limites", 1)
    add_body(doc, "A estimativa registrada no levantamento — média de R$ 1,30 por chamada — deve ser considerada apenas uma hipótese. O custo real inclui geração, leitura de arquivos, OCR, armazenamento, exportação, reprocessamentos e chamadas de validação.")
    add_callout(doc, "Regra de controle:", "não enviar documentos inteiros em todas as chamadas. Extrair, resumir, dividir em trechos e recuperar somente o contexto relevante para cada etapa.")
    add_table(doc,
              ["Controle", "Aplicação"],
              [
                  ("Limite de entrada", "Definir tamanho máximo de arquivo, páginas, imagens e caracteres por modo."),
                  ("Orçamento de tokens", "Separar limites para extração, planejamento, geração e revisão."),
                  ("Cache", "Reutilizar texto extraído, resumos e classificações de arquivos."),
                  ("Roteamento de modelo", "Usar modelo rápido para classificação e formatação; reservar modelo mais caro para tarefas complexas."),
                  ("Quotas", "Limitar gerações por usuário e plano, com indicação transparente do consumo."),
                  ("Observabilidade", "Medir custo por pacote concluído, não apenas por chamada individual."),
              ],
              [2300, 7060], font_size=9.2)

    add_heading(doc, "10. Planos e monetização", 1)
    add_body(doc, "Os planos devem limitar volume e recursos de valor, não apenas esconder funcionalidades. Uma primeira hipótese de estrutura é:")
    add_table(doc,
              ["Plano", "Proposta inicial", "Objetivo"],
              [
                  ("Gratuito", "Fast, limite mensal de gerações, PDF e pacote básico.", "Aquisição e validação."),
                  ("Professor", "Criativo, mais arquivos, DOCX, banco de questões e histórico opcional.", "Monetizar uso individual recorrente."),
                  ("Escola", "Espaço compartilhado, currículo institucional, permissões e relatórios.", "Venda B2B em etapa posterior."),
              ],
              [1500, 5100, 2760], font_size=9.2)

    add_heading(doc, "11. Privacidade e persistência", 1)
    add_body(doc, "A decisão de não guardar histórico reduz riscos de privacidade, mas também reduz reutilização, continuidade e capacidade de melhorar o produto. Como o banco de questões já exige persistência, a recomendação é permitir que o usuário escolha o que deseja salvar.")
    add_bullet(doc, "Histórico opcional, com exclusão manual e prazo de retenção configurável.")
    add_bullet(doc, "Arquivos temporários excluídos automaticamente após o processamento, quando não forem salvos.")
    add_bullet(doc, "Separação entre materiais privados do professor e conteúdo compartilhado com a escola.")
    add_bullet(doc, "Aviso claro de que o material gerado precisa de revisão profissional antes do uso.")

    add_heading(doc, "12. Validação do MVP", 1)
    add_body(doc, "Antes de ampliar o produto, é necessário testar se o professor percebe valor no pacote completo e se a geração reduz trabalho sem aumentar o tempo de revisão.")
    add_table(doc,
              ["Hipótese", "Como testar", "Sinal inicial de sucesso"],
              [
                  ("O contexto da turma melhora a utilidade", "Comparar uma geração genérica com uma geração contextualizada.", "Maior avaliação de utilidade e menos edições estruturais."),
                  ("O pacote completo é melhor que um plano isolado", "Observar o que o usuário realmente exporta e aplica.", "Uso recorrente de atividade e avaliação no mesmo fluxo."),
                  ("A rastreabilidade aumenta confiança", "Apresentar habilidades e fontes junto ao material.", "Menos dúvidas sobre origem e alinhamento do conteúdo."),
                  ("O Fast reduz fricção", "Medir tempo até a primeira saída útil.", "Primeiro resultado em poucos minutos, sem suporte manual."),
                  ("O Criativo justifica pagamento", "Testar disposição de pagar com professores ativos.", "Uso recorrente de referências, adaptações e recuperação."),
              ],
              [2350, 4110, 2900], font_size=9.0)

    add_heading(doc, "Métricas sugeridas", 2)
    add_bullet(doc, "Tempo até a primeira saída utilizável.")
    add_bullet(doc, "Percentual de gerações exportadas ou salvas.")
    add_bullet(doc, "Quantidade média de edições antes da aplicação.")
    add_bullet(doc, "Nota de utilidade atribuída pelo professor.")
    add_bullet(doc, "Custo médio por pacote concluído.")
    add_bullet(doc, "Taxa de retorno do usuário em até 30 dias.")

    add_heading(doc, "13. Riscos principais", 1)
    add_table(doc,
              ["Risco", "Impacto", "Mitigação inicial"],
              [
                  ("Escopo amplo", "Produto superficial e difícil de validar.", "Escolher um segmento e um fluxo principal."),
                  ("Conteúdo incorreto", "Perda de confiança ou uso pedagógico inadequado.", "Fontes, validação estruturada e revisão humana obrigatória."),
                  ("Custo elevado", "Margem negativa em planos com uso intenso.", "Quotas, roteamento, cache e medição por pacote."),
                  ("Banco de questões sem qualidade", "Atividades repetitivas ou mal alinhadas.", "Metadados, revisão, versionamento e controle de duplicatas."),
                  ("Exportação inconsistente", "Material visualmente ruim ou difícil de usar.", "Modelo canônico e testes de renderização por formato."),
                  ("Privacidade de materiais", "Risco para professores e instituições.", "Retenção clara, exclusão, permissões e consentimento."),
              ],
              [2200, 3000, 4160], font_size=9.0)

    add_heading(doc, "14. Próximas decisões", 1)
    add_numbered_list(doc, [
        "Escolher o segmento inicial e entrevistar de cinco a dez professores.",
        "Mapear o formato de planejamento mais usado por esse segmento.",
        "Definir o pacote mínimo que o professor considera pronto para aplicar.",
        "Construir um protótipo navegável do fluxo Fast e Criativo.",
        "Testar geração com materiais reais e medir tempo, qualidade e custo.",
        "Decidir se a evidência de uso justifica banco de questões avançado, aplicativo mobile e venda para escolas.",
    ])

    add_heading(doc, "Conclusão e recomendação", 1)
    add_body(doc, "A proposta tem potencial, mas o caminho mais seguro é não começar como uma plataforma para todos os cursos e públicos. O MVP deve provar uma promessa específica: ajudar um professor a transformar seus próprios materiais e o contexto da turma em um pacote pedagógico coerente, revisável e exportável.")
    add_callout(doc, "Decisão recomendada:", "validar primeiro a combinação entre contexto da turma, rastreabilidade pedagógica e ciclo aula–avaliação–recuperação. Esses elementos formam uma proposta mais defensável do que simplesmente oferecer mais formatos de arquivo.", fill="FFF8E8", accent="C59A27")

    add_heading(doc, "Fontes de benchmark", 1)
    add_body(doc, "Páginas públicas consultadas em 27 de agosto de 2026:", after=4)
    add_source(doc, "Teachy — professores e ferramentas", "https://www.teachy.com.br/pt-BR/teachers")
    add_source(doc, "Chalkie — assistente de planejamento de aulas", "https://chalkie.ai/pt-br")
    add_source(doc, "Avalea — gerador de plano de aula", "https://avalea.qconcursos.com/ferramentas/gerador-de-plano-de-aula")

    # Keep the final paragraph together with its content where possible.
    for p in doc.paragraphs[-5:]:
        p.paragraph_format.widow_control = True

    OUT.parent.mkdir(parents=True, exist_ok=True)
    doc.save(OUT)
    print(OUT)


if __name__ == "__main__":
    build()
