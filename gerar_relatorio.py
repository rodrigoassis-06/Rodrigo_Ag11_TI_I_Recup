import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
import os

def gerar_relatorio():
    doc = docx.Document()

    # Margens padrao
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    # Cabecalho
    p_head = doc.add_paragraph()
    r_head = p_head.add_run(
        "CENTRO PAULA SOUZA - EAD TEC\n"
        "CURSO TECNICO EM DESENVOLVIMENTO DE SISTEMAS\n"
        "AGENDA 11 - TECNOLOGIAS DA INFORMACAO I (RECUPERACAO)"
    )
    r_head.bold = True
    r_head.font.size = Pt(11)
    r_head.font.color.rgb = RGBColor(70, 70, 70)
    p_head.alignment = WD_ALIGN_PARAGRAPH.CENTER

    p_div = doc.add_paragraph()
    r_div = p_div.add_run("-" * 70)
    r_div.font.color.rgb = RGBColor(180, 180, 180)
    p_div.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Titulo Principal
    p_title = doc.add_paragraph()
    r_title = p_title.add_run("RELATORIO DE MODELAGEM UML: DIAGRAMA DE CASOS DE USO COM IA")
    r_title.bold = True
    r_title.font.size = Pt(14)
    r_title.font.color.rgb = RGBColor(9, 105, 218)
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Identificacao do Aluno
    p_id = doc.add_paragraph()
    r_id = p_id.add_run("Aluno: Rodrigo Assis | Formato: Rodrigo_Ag11_TI_I_Recup")
    r_id.bold = True
    r_id.font.size = Pt(10.5)
    p_id.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph()

    # 1. Introducao e Cenario
    doc.add_heading("1. Cenario do Projeto e Analise Inicial", level=1)
    doc.add_paragraph(
        "A ETEC propos o desenvolvimento de um Sistema de Matricula Online com objetivo de integrar os processos "
        "academicos e administrativos entre alunos, professores, secretaria academica e servicos externos de cobranca.\n\n"
        "Na etapa preliminar de analise, foram levantados os seguintes atores e funcionalidades candidatas:"
    )

    table = doc.add_table(rows=1, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = "Ator Identificado"
    hdr_cells[1].text = "Casos de Uso Preliminares"
    hdr_cells[0].paragraphs[0].runs[0].bold = True
    hdr_cells[1].paragraphs[0].runs[0].bold = True

    dados_iniciais = [
        ("Aluno", "Cadastrar-se, Visualizar Cursos, Matricular-se, Ver Notas e Historico"),
        ("Professor", "Acessar Lista de Alunos, Registrar Notas, Registrar Frequencia, Disponibilizar Materiais"),
        ("Secretaria Academica", "Gerenciar Alunos/Professores, Criar e Editar Cursos, Aprovar Matriculas, Gerar Relatorios"),
        ("Sistema Externo de Pagamentos", "Processar taxas de matricula")
    ]

    for ator, funcs in dados_iniciais:
        row_cells = table.add_row().cells
        row_cells[0].text = ator
        row_cells[1].text = funcs

    doc.add_paragraph()

    # 2. Refinamento com IA
    doc.add_heading("2. Refinamento e Validacao com Inteligencia Artificial", level=1)
    doc.add_paragraph(
        "Como assistente de apoio a tomada de decisao tecnica, utilizou-se a Inteligencia Artificial para validar e "
        "refinar os seguintes pontos criticos da modelagem UML:\n"
        "a) Papel do Sistema de Pagamentos: definido como Ator Secundario (ator sistema externo), que interage com o caso de uso Processar Taxas.\n"
        "b) Relacionamento <<include>>: estabelecido entre Realizar Matricula e Processar Taxas, evidenciando a obrigatoriedade da taxa no processo de matricula.\n"
        "c) Independencia de Atores: descartada generalizacao desnecessaria de atores, preservando a especificidade das permissoes de Aluno e Professor.\n"
        "d) Padronizacao de Nomenclatura: aplicacao estrita de verbos no infinitivo para os casos de uso."
    )

    img_ia = os.path.join(os.path.dirname(__file__), "interacao_ia.png")
    if os.path.exists(img_ia):
        doc.add_paragraph().add_run("Evidencia da Interacao com a IA:").bold = True
        p_img_ia = doc.add_paragraph()
        p_img_ia.alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_picture(img_ia, width=Inches(5.8))
        p_cap_ia = doc.add_paragraph("Figura 1: Registro do dialogo de validacao dos requisitos e modelagem com IA.")
        p_cap_ia.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cap_ia.runs[0].font.size = Pt(9)
        p_cap_ia.runs[0].italic = True

    doc.add_paragraph()

    # 3. Diagrama Final
    doc.add_heading("3. Diagrama Final de Casos de Uso (UML)", level=1)
    doc.add_paragraph(
        "O diagrama final foi desenhado contemplando os limites do sistema (retangulo), os atores devidamente posicionados "
        "externamente, os casos de uso representados por elipses e suas respectivas associacoes e relacionamentos estereotipados."
    )

    img_diag = os.path.join(os.path.dirname(__file__), "diagrama_caso_de_uso.png")
    if os.path.exists(img_diag):
        p_img_d = doc.add_paragraph()
        p_img_d.alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_picture(img_diag, width=Inches(6.0))
        p_cap_d = doc.add_paragraph("Figura 2: Diagrama de Casos de Uso completo do Sistema de Matricula Online da ETEC.")
        p_cap_d.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cap_d.runs[0].font.size = Pt(9)
        p_cap_d.runs[0].italic = True

    doc.add_paragraph()

    # 4. Detalhamento
    doc.add_heading("4. Detalhamento dos Componentes Modelados", level=1)
    p_at = doc.add_paragraph()
    p_at.add_run("Atores do Sistema:\n").bold = True
    p_at.add_run(
        "- Aluno: ator primario que realiza cadastro, visualiza cursos, solicita matricula e consulta notas/historico.\n"
        "- Professor: ator primario responsavel pelo acesso a lista de alunos, lancamento de notas/frequencia e disponibilizacao de materiais.\n"
        "- Secretaria Academica: ator primario administrativo que aprova matriculas, mantem cursos/usuarios e gera relatorios.\n"
        "- Sistema de Pagamentos: ator secundario externo responsavel pela liquidacao financeira das taxas de matricula."
    )

    p_uc = doc.add_paragraph()
    p_uc.add_run("Casos de Uso:\n").bold = True
    p_uc.add_run(
        "- UC01 - Cadastrar-se: inclusao de novos estudantes no sistema.\n"
        "- UC02 - Visualizar Cursos: consulta aos cursos e modulos ofertados.\n"
        "- UC03 - Realizar Matricula: inscricao nos componentes curriculares.\n"
        "- UC04 - Ver Notas e Historico: acompanhamento do progresso escolar.\n"
        "- UC05 - Processar Taxas: cobranca e verificacao de pagamento (incluido via <<include>> na matricula).\n"
        "- UC06 - Aprovar Matriculas: avaliacao e deferimento pela secretaria.\n"
        "- UC07 - Criar/Editar Cursos: gestao da grade curricular e vagas.\n"
        "- UC08 - Gerenciar Alunos/Professores: manutencao cadastral de usuarios.\n"
        "- UC09 - Gerar Relatorios: emissao de dados quantitativos e academicos.\n"
        "- UC10 - Acessar Lista de Alunos: visualizacao das turmas atribuidas ao docente.\n"
        "- UC11 - Registrar Notas e Frequencia: apontamento continuo no diario de classe.\n"
        "- UC12 - Disponibilizar Materiais: envio de arquivos e orientacoes didaticas."
    )

    # 5. Conclusao
    doc.add_heading("5. Conclusao", level=1)
    doc.add_paragraph(
        "A modelagem UML alcancou total aderencia ao cenario proposto para o Sistema de Matricula Online da ETEC. "
        "A utilizacao da IA agilizou a validacao dos conceitos formais de atores secundarios e relacionamentos de inclusao, "
        "resultando em uma especificacao consistente e pronta para a implementacao."
    )

    out_file = os.path.join(os.path.dirname(__file__), "Rodrigo_Ag11_TI_I_Recup.docx")
    doc.save(out_file)
    print("Documento salvo em:", out_file)
    print("Tamanho:", os.path.getsize(out_file), "bytes")

if __name__ == "__main__":
    gerar_relatorio()
