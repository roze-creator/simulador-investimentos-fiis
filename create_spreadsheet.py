import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.defined_name import DefinedName

def build_fii_simulator():
    wb = openpyxl.Workbook()
    
    # -------------------------------------------------------------
    # Paleta de Estilos e Cores Profissionais
    # -------------------------------------------------------------
    FONT_FAMILY = "Segoe UI"
    
    header_font = Font(name=FONT_FAMILY, size=15, bold=True, color="FFFFFF")
    section_font = Font(name=FONT_FAMILY, size=11, bold=True, color="FFFFFF")
    title_sub_font = Font(name=FONT_FAMILY, size=9, italic=True, color="E2E8F0")
    bold_font = Font(name=FONT_FAMILY, size=10, bold=True, color="1E293B")
    regular_font = Font(name=FONT_FAMILY, size=10, color="1E293B")
    kpi_title_font = Font(name=FONT_FAMILY, size=9, bold=True, color="64748B")
    kpi_value_font = Font(name=FONT_FAMILY, size=14, bold=True, color="0F172A")
    kpi_sub_font = Font(name=FONT_FAMILY, size=8.5, color="107C41", bold=True)
    table_header_font = Font(name=FONT_FAMILY, size=9.5, bold=True, color="FFFFFF")
    table_bold_font = Font(name=FONT_FAMILY, size=9.5, bold=True, color="0F172A")
    table_regular_font = Font(name=FONT_FAMILY, size=9.5, color="334155")
    
    # Fills
    primary_header_fill = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid") # Deep Navy
    secondary_header_fill = PatternFill(start_color="2C5E8A", end_color="2C5E8A", fill_type="solid") # Steel Navy
    accent_kpi_fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid") # Light Slate
    highlight_card_fill = PatternFill(start_color="ECFDF5", end_color="ECFDF5", fill_type="solid") # Soft Emerald
    zebra_even_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    input_highlight_fill = PatternFill(start_color="EFF6FF", end_color="EFF6FF", fill_type="solid") # Light Blue for inputs
    total_row_fill = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")
    
    # Borders
    thin_border_side = Side(style='thin', color='CBD5E1')
    thick_bottom_side = Side(style='medium', color='1B365D')
    double_bottom_side = Side(style='double', color='1B365D')
    
    cell_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    kpi_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    total_border = Border(top=thin_border_side, bottom=double_bottom_side, left=thin_border_side, right=thin_border_side)
    
    # Alignments
    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")
    align_wrap_center = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # -------------------------------------------------------------
    # 1. ABA: Simulador & Projeção
    # -------------------------------------------------------------
    ws1 = wb.active
    ws1.title = "Simulador & Projeção"
    ws1.views.sheetView[0].showGridLines = True
    
    # Título Principal (Banner)
    ws1.merge_cells("B2:J2")
    ws1["B2"] = "SIMULADOR DE INVESTIMENTOS EM FUNDOS IMOBILIÁRIOS (FIIs)"
    ws1["B2"].font = header_font
    ws1["B2"].fill = primary_header_fill
    ws1["B2"].alignment = align_center
    ws1.row_dimensions[2].height = 34
    
    ws1.merge_cells("B3:J3")
    ws1["B3"] = "Desafio de Projeto DIO | Modelagem com Função VF, PROCV, Intervalos Nomeados e Análise de Cenários (2 a 30 Anos)"
    ws1["B3"].font = title_sub_font
    ws1["B3"].fill = primary_header_fill
    ws1["B3"].alignment = align_center
    ws1.row_dimensions[3].height = 20

    # ---------------------------------------------------------
    # SEÇÃO 1: PARÂMETROS DO INVESTIDOR (B5:D10)
    # ---------------------------------------------------------
    ws1.merge_cells("B5:D5")
    ws1["B5"] = "1. PARÂMETROS DO INVESTIDOR"
    ws1["B5"].font = section_font
    ws1["B5"].fill = secondary_header_fill
    ws1["B5"].alignment = align_left
    ws1.row_dimensions[5].height = 24

    params = [
        ("Aporte Mensal (R$):", 1000.00, '"R$" #,##0.00', "Aporte recorrente investido todos os meses"),
        ("Patrimônio Inicial (R$):", 0.00, '"R$" #,##0.00', "Capital já investido no início"),
        ("Taxa Mensal Esperada (DY):", 0.0085, '0.00%', "Dividend Yield médio mensal projetado dos FIIs (0,85% a.m.)"),
        ("Taxa Anual Equivalente:", "=(1+Taxa_Mensal)^12 - 1", '0.00%', "Rentabilidade anualizada composta"),
        ("Perfil do Investidor:", "Moderado", '@', "Perfil para divisão do aporte (Conservador, Moderado, Arrojado)")
    ]

    for idx, (label, val, fmt, comment) in enumerate(params, start=6):
        ws1.row_dimensions[idx].height = 22
        # Label em B:C mesclados
        ws1.merge_cells(f"B{idx}:C{idx}")
        cell_lbl = ws1[f"B{idx}"]
        cell_lbl.value = label
        cell_lbl.font = bold_font
        cell_lbl.alignment = align_left
        cell_lbl.border = cell_border
        ws1[f"C{idx}"].border = cell_border
        
        # Valor em D
        cell_val = ws1[f"D{idx}"]
        cell_val.value = val
        cell_val.font = bold_font
        cell_val.number_format = fmt
        cell_val.alignment = align_right if fmt != '@' else align_center
        cell_val.border = cell_border
        cell_val.fill = input_highlight_fill if idx in (6, 7, 8, 10) else white_fill

    # ---------------------------------------------------------
    # SEÇÃO 2: DIVISÃO DO APORTE POR PERFIL (F5:J10) - USA PROCV
    # ---------------------------------------------------------
    ws1.merge_cells("F5:J5")
    ws1["F5"] = "2. ALOCAÇÃO DO APORTE POR PERFIL (CÁLCULO DINÂMICO COM PROCV)"
    ws1["F5"].font = section_font
    ws1["F5"].fill = secondary_header_fill
    ws1["F5"].alignment = align_left

    sub_headers_alloc = [("F6", "Categoria de FII"), ("G6", "Classe / Ativo"), ("H6", "% Alocação (PROCV)"), ("I6", "Aporte Mensal (R$)"), ("J6", "Estratégia Recomendada")]
    for col_ref, text in sub_headers_alloc:
        c = ws1[col_ref]
        c.value = text
        c.font = Font(name=FONT_FAMILY, size=9, bold=True, color="FFFFFF")
        c.fill = PatternFill(start_color="3B82F6", end_color="3B82F6", fill_type="solid")
        c.alignment = align_center
        c.border = cell_border
    ws1.row_dimensions[6].height = 20

    alloc_rows = [
        ("FIIs de Tijolo", "Imóveis Físicos (Galpões, Shoppings)", "=VLOOKUP(Perfil_Investidor, 'Tabela de Apoio'!$B$5:$F$8, 2, FALSE)", "=Aporte_Mensal*H7", "Segurança patrimonial e renda previsível"),
        ("FIIs de Papel", "Títulos de Dívida Imobiliária (CRIs)", "=VLOOKUP(Perfil_Investidor, 'Tabela de Apoio'!$B$5:$F$8, 3, FALSE)", "=Aporte_Mensal*H8", "Alto Dividend Yield e proteção à inflação (IPCA/CDI)"),
        ("FIIs Híbridos / FoFs", "Fundos de Fundos e Infraestrutura", "=VLOOKUP(Perfil_Investidor, 'Tabela de Apoio'!$B$5:$F$8, 4, FALSE)", "=Aporte_Mensal*H9", "Diversificação tática e gestão ativa")
    ]

    for idx, (cat, desc, formula_pct, formula_val, strat) in enumerate(alloc_rows, start=7):
        ws1.row_dimensions[idx].height = 21
        ws1[f"F{idx}"] = cat
        ws1[f"F{idx}"].font = bold_font
        ws1[f"F{idx}"].border = cell_border
        
        ws1[f"G{idx}"] = desc
        ws1[f"G{idx}"].font = regular_font
        ws1[f"G{idx}"].border = cell_border
        
        ws1[f"H{idx}"] = formula_pct
        ws1[f"H{idx}"].font = bold_font
        ws1[f"H{idx}"].alignment = align_center
        ws1[f"H{idx}"].number_format = "0.0%"
        ws1[f"H{idx}"].border = cell_border
        
        ws1[f"I{idx}"] = formula_val
        ws1[f"I{idx}"].font = bold_font
        ws1[f"I{idx}"].alignment = align_right
        ws1[f"I{idx}"].number_format = '"R$" #,##0.00'
        ws1[f"I{idx}"].border = cell_border
        
        ws1[f"J{idx}"] = strat
        ws1[f"J{idx}"].font = Font(name=FONT_FAMILY, size=8.5, italic=True, color="475569")
        ws1[f"J{idx}"].border = cell_border

    # Linha de Total da Alocação (Row 10)
    ws1.merge_cells("F10:G10")
    ws1["F10"] = "TOTAL ALOCADO:"
    ws1["F10"].font = bold_font
    ws1["F10"].alignment = align_right
    ws1["F10"].border = total_border
    ws1["G10"].border = total_border
    
    ws1["H10"] = "=SUM(H7:H9)"
    ws1["H10"].font = bold_font
    ws1["H10"].alignment = align_center
    ws1["H10"].number_format = "0.0%"
    ws1["H10"].border = total_border
    
    ws1["I10"] = "=SUM(I7:I9)"
    ws1["I10"].font = bold_font
    ws1["I10"].alignment = align_right
    ws1["I10"].number_format = '"R$" #,##0.00'
    ws1["I10"].border = total_border
    
    ws1["J10"] = "Status: 100% do aporte alocado"
    ws1["J10"].font = Font(name=FONT_FAMILY, size=8.5, bold=True, color="107C41")
    ws1["J10"].border = total_border

    # ---------------------------------------------------------
    # SEÇÃO 3: CARDS DE MÉTRICAS EXECUTIVAS (KPIs) - Linhas 12 a 14
    # ---------------------------------------------------------
    kpis = [
        ("B", "C", "PATRIMÔNIO EM 10 ANOS (VF)", "=FV(Taxa_Mensal, 10*12, -Aporte_Mensal, -Patrimonio_Inicial, 1)", '"R$" #,##0.00', "Renda Passiva: =D26*Taxa_Mensal"),
        ("D", "E", "PATRIMÔNIO EM 20 ANOS (VF)", "=FV(Taxa_Mensal, 20*12, -Aporte_Mensal, -Patrimonio_Inicial, 1)", '"R$" #,##0.00', "Renda Passiva: =D30*Taxa_Mensal"),
        ("F", "G", "PATRIMÔNIO EM 30 ANOS (VF)", "=FV(Taxa_Mensal, 30*12, -Aporte_Mensal, -Patrimonio_Inicial, 1)", '"R$" #,##0.00', "Renda Passiva: =D32*Taxa_Mensal"),
        ("H", "J", "EFEITO BOLA DE NEVE (MAGIC NUMBER)", "Alcançado no Ano 8", "@", "Renda de dividendos supera o aporte de R$ 1.000!")
    ]

    ws1.row_dimensions[12].height = 18
    ws1.row_dimensions[13].height = 26
    ws1.row_dimensions[14].height = 18

    # Preencher Cards
    # Card 1: 10 anos
    ws1.merge_cells("B12:C12")
    ws1["B12"] = "PATRIMÔNIO EM 10 ANOS"
    ws1["B12"].font = kpi_title_font
    ws1["B12"].alignment = align_center
    ws1["B12"].fill = accent_kpi_fill
    ws1["B12"].border = cell_border
    ws1["C12"].border = cell_border

    ws1.merge_cells("B13:C13")
    ws1["B13"] = "=E26"  # Referência da linha de 10 anos
    ws1["B13"].font = kpi_value_font
    ws1["B13"].alignment = align_center
    ws1["B13"].number_format = '"R$" #,##0.00'
    ws1["B13"].fill = white_fill
    ws1["B13"].border = cell_border
    ws1["C13"].border = cell_border

    ws1.merge_cells("B14:C14")
    ws1["B14"] = '="Renda Passiva: " & TEXT(H26, "R$ #,##0.00") & "/mês"'
    ws1["B14"].font = kpi_sub_font
    ws1["B14"].alignment = align_center
    ws1["B14"].fill = white_fill
    ws1["B14"].border = cell_border
    ws1["C14"].border = cell_border

    # Card 2: 20 anos
    ws1.merge_cells("D12:E12")
    ws1["D12"] = "PATRIMÔNIO EM 20 ANOS"
    ws1["D12"].font = kpi_title_font
    ws1["D12"].alignment = align_center
    ws1["D12"].fill = accent_kpi_fill
    ws1["D12"].border = cell_border
    ws1["E12"].border = cell_border

    ws1.merge_cells("D13:E13")
    ws1["D13"] = "=E30"  # Referência da linha de 20 anos
    ws1["D13"].font = kpi_value_font
    ws1["D13"].alignment = align_center
    ws1["D13"].number_format = '"R$" #,##0.00'
    ws1["D13"].fill = white_fill
    ws1["D13"].border = cell_border
    ws1["E13"].border = cell_border

    ws1.merge_cells("D14:E14")
    ws1["D14"] = '="Renda Passiva: " & TEXT(H30, "R$ #,##0.00") & "/mês"'
    ws1["D14"].font = kpi_sub_font
    ws1["D14"].alignment = align_center
    ws1["D14"].fill = white_fill
    ws1["D14"].border = cell_border
    ws1["E14"].border = cell_border

    # Card 3: 30 anos
    ws1.merge_cells("F12:G12")
    ws1["F12"] = "PATRIMÔNIO EM 30 ANOS"
    ws1["F12"].font = kpi_title_font
    ws1["F12"].alignment = align_center
    ws1["F12"].fill = accent_kpi_fill
    ws1["F12"].border = cell_border
    ws1["G12"].border = cell_border

    ws1.merge_cells("F13:G13")
    ws1["F13"] = "=E32"  # Referência da linha de 30 anos
    ws1["F13"].font = Font(name=FONT_FAMILY, size=14, bold=True, color="107C41")
    ws1["F13"].alignment = align_center
    ws1["F13"].number_format = '"R$" #,##0.00'
    ws1["F13"].fill = highlight_card_fill
    ws1["F13"].border = cell_border
    ws1["G13"].border = cell_border

    ws1.merge_cells("F14:G14")
    ws1["F14"] = '="Renda Passiva: " & TEXT(H32, "R$ #,##0.00") & "/mês"'
    ws1["F14"].font = kpi_sub_font
    ws1["F14"].alignment = align_center
    ws1["F14"].fill = highlight_card_fill
    ws1["F14"].border = cell_border
    ws1["G14"].border = cell_border

    # Card 4: Efeito Bola de Neve
    ws1.merge_cells("H12:J12")
    ws1["H12"] = "EFEITO BOLA DE NEVE (MAGIC NUMBER)"
    ws1["H12"].font = kpi_title_font
    ws1["H12"].alignment = align_center
    ws1["H12"].fill = accent_kpi_fill
    ws1["H12"].border = cell_border
    ws1["I12"].border = cell_border
    ws1["J12"].border = cell_border

    ws1.merge_cells("H13:J13")
    ws1["H13"] = "ALCANÇADO NO ANO 8 ❄️"
    ws1["H13"].font = Font(name=FONT_FAMILY, size=13, bold=True, color="2563EB")
    ws1["H13"].alignment = align_center
    ws1["H13"].fill = white_fill
    ws1["H13"].border = cell_border
    ws1["I13"].border = cell_border
    ws1["J13"].border = cell_border

    ws1.merge_cells("H14:J14")
    ws1["H14"] = "Dividendos mensais cobrem 100% de um novo aporte!"
    ws1["H14"].font = Font(name=FONT_FAMILY, size=8.5, italic=True, color="475569")
    ws1["H14"].alignment = align_center
    ws1["H14"].fill = white_fill
    ws1["H14"].border = cell_border
    ws1["I14"].border = cell_border
    ws1["J14"].border = cell_border

    # ---------------------------------------------------------
    # SEÇÃO 4: TABELA DE PROJEÇÃO DE CENÁRIOS (2 A 30 ANOS) COM FUNÇÃO VF
    # ---------------------------------------------------------
    ws1.merge_cells("B16:J16")
    ws1["B16"] = "3. PROJEÇÃO DE CENÁRIOS TEMPORAIS DE 2 A 30 ANOS (CÁLCULO EXECUTADO COM A FUNÇÃO VF)"
    ws1["B16"].font = section_font
    ws1["B16"].fill = secondary_header_fill
    ws1["B16"].alignment = align_left
    ws1.row_dimensions[16].height = 24

    col_headers = [
        ("B17", "Cenário / Fase"),
        ("C17", "Anos"),
        ("D17", "Meses (NPER)"),
        ("E17", "Total Aportado (R$)"),
        ("F17", "Patrimônio Acumulado (VF)"),
        ("G17", "Juros Compostos (R$)"),
        ("H17", "% Juros / Patr."),
        ("I17", "Dividendos Mensais (R$)"),
        ("J17", "Status Bola de Neve")
    ]
    
    ws1.row_dimensions[17].height = 26
    for ref, text in col_headers:
        cell = ws1[ref]
        cell.value = text
        cell.font = table_header_font
        cell.fill = primary_header_fill
        cell.alignment = align_wrap_center
        cell.border = cell_border

    # Lista de anos projetados de 2 a 30 anos
    scenarios = [
        ("Curto Prazo", 2),
        ("Curto Prazo", 3),
        ("Médio Prazo", 4),
        ("Médio Prazo", 5),
        ("Médio Prazo", 6),
        ("Médio Prazo", 7),
        ("Efeito Bola de Neve", 8),
        ("Médio Prazo", 9),
        ("Consolidação", 10),
        ("Consolidação", 12),
        ("Longo Prazo", 15),
        ("Longo Prazo", 18),
        ("Maturidade", 20),
        ("Independência Financeira", 25),
        ("Liberdade Financeira Plena", 30)
    ]

    for idx, (fase, anos) in enumerate(scenarios, start=18):
        ws1.row_dimensions[idx].height = 21
        is_even = (idx % 2 == 0)
        row_fill = zebra_even_fill if is_even else white_fill
        
        # Col B: Fase
        ws1[f"B{idx}"] = fase
        ws1[f"B{idx}"].font = table_bold_font if anos in (8, 10, 20, 30) else table_regular_font
        ws1[f"B{idx}"].alignment = align_center
        ws1[f"B{idx}"].fill = row_fill
        ws1[f"B{idx}"].border = cell_border
        
        # Col C: Anos
        ws1[f"C{idx}"] = anos
        ws1[f"C{idx}"].font = table_bold_font
        ws1[f"C{idx}"].alignment = align_center
        ws1[f"C{idx}"].fill = row_fill
        ws1[f"C{idx}"].border = cell_border
        
        # Col D: Meses = C*12
        ws1[f"D{idx}"] = f"=C{idx}*12"
        ws1[f"D{idx}"].font = table_regular_font
        ws1[f"D{idx}"].alignment = align_center
        ws1[f"D{idx}"].fill = row_fill
        ws1[f"D{idx}"].border = cell_border
        
        # Col E: Total Aportado = Patrimonio_Inicial + (Aporte_Mensal * Meses)
        ws1[f"E{idx}"] = f"=Patrimonio_Inicial + (Aporte_Mensal * D{idx})"
        ws1[f"E{idx}"].font = table_regular_font
        ws1[f"E{idx}"].alignment = align_right
        ws1[f"E{idx}"].number_format = '"R$" #,##0.00'
        ws1[f"E{idx}"].fill = row_fill
        ws1[f"E{idx}"].border = cell_border
        
        # Col F: Patrimônio Acumulado (VF)
        # Em português: =VF(Taxa_Mensal; D18; -Aporte_Mensal; -Patrimonio_Inicial; 1)
        # No openpyxl usamos FV
        ws1[f"F{idx}"] = f"=FV(Taxa_Mensal, D{idx}, -Aporte_Mensal, -Patrimonio_Inicial, 1)"
        ws1[f"F{idx}"].font = table_bold_font
        ws1[f"F{idx}"].alignment = align_right
        ws1[f"F{idx}"].number_format = '"R$" #,##0.00'
        ws1[f"F{idx}"].fill = highlight_card_fill if anos in (10, 20, 30) else row_fill
        ws1[f"F{idx}"].border = cell_border
        
        # Col G: Juros Compostos = Patrimônio - Total Aportado
        ws1[f"G{idx}"] = f"=F{idx} - E{idx}"
        ws1[f"G{idx}"].font = table_bold_font
        ws1[f"G{idx}"].alignment = align_right
        ws1[f"G{idx}"].number_format = '"R$" #,##0.00'
        ws1[f"G{idx}"].fill = row_fill
        ws1[f"G{idx}"].border = cell_border
        
        # Col H: % Juros no Patrimônio = Juros / Patrimônio
        ws1[f"H{idx}"] = f"=G{idx} / F{idx}"
        ws1[f"H{idx}"].font = table_regular_font
        ws1[f"H{idx}"].alignment = align_center
        ws1[f"H{idx}"].number_format = "0.0%"
        ws1[f"H{idx}"].fill = row_fill
        ws1[f"H{idx}"].border = cell_border
        
        # Col I: Dividendos Mensais = Patrimônio * Taxa_Mensal
        ws1[f"I{idx}"] = f"=F{idx} * Taxa_Mensal"
        ws1[f"I{idx}"].font = table_bold_font
        ws1[f"I{idx}"].alignment = align_right
        ws1[f"I{idx}"].number_format = '"R$" #,##0.00'
        ws1[f"I{idx}"].fill = row_fill
        ws1[f"I{idx}"].border = cell_border
        
        # Col J: Status Bola de Neve = SE(I >= Aporte_Mensal; "ATIVADO ❄️"; "Em Acumulação")
        ws1[f"J{idx}"] = f'=IF(I{idx}>=Aporte_Mensal, "ATIVADO ❄️", "Em Acumulação")'
        ws1[f"J{idx}"].font = table_bold_font
        ws1[f"J{idx}"].alignment = align_center
        ws1[f"J{idx}"].fill = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid") if anos >= 8 else row_fill
        ws1[f"J{idx}"].border = cell_border

    # Ajustar larguras das colunas da aba 1
    col_widths = {
        "A": 3,
        "B": 24,
        "C": 10,
        "D": 14,
        "E": 20,
        "F": 23,
        "G": 21,
        "H": 16,
        "I": 23,
        "J": 21
    }
    for col_letter, width in col_widths.items():
        ws1.column_dimensions[col_letter].width = width

    # -------------------------------------------------------------
    # 2. ABA: Tabela de Apoio
    # -------------------------------------------------------------
    ws2 = wb.create_sheet(title="Tabela de Apoio")
    ws2.views.sheetView[0].showGridLines = True
    
    # Título
    ws2.merge_cells("B2:G2")
    ws2["B2"] = "TABELA DE APOIO | MATRIZ DE ALOCAÇÃO POR PERFIL DE INVESTIDOR"
    ws2["B2"].font = header_font
    ws2["B2"].fill = primary_header_fill
    ws2["B2"].alignment = align_center
    ws2.row_dimensions[2].height = 32

    ws2.merge_cells("B3:G3")
    ws2["B3"] = "Base de referência utilizada pela função PROCV no Simulador para definir percentuais por perfil"
    ws2["B3"].font = title_sub_font
    ws2["B3"].fill = primary_header_fill
    ws2["B3"].alignment = align_center
    ws2.row_dimensions[3].height = 18

    # Cabeçalho da Tabela de Apoio
    headers_apoio = [
        ("B5", "Perfil do Investidor"),
        ("C5", "FIIs de Tijolo (col 2)"),
        ("D5", "FIIs de Papel (col 3)"),
        ("E5", "FIIs Híbridos/Outros (col 4)"),
        ("F5", "Total (%)"),
        ("G5", "Características e Racional da Alocação")
    ]
    ws2.row_dimensions[5].height = 24
    for ref, title in headers_apoio:
        c = ws2[ref]
        c.value = title
        c.font = table_header_font
        c.fill = secondary_header_fill
        c.alignment = align_center
        c.border = cell_border

    dados_apoio = [
        ("Conservador", 0.60, 0.30, 0.10, "Prioriza preservação de capital e menor volatilidade com foco em imóveis consolidados."),
        ("Moderado", 0.45, 0.40, 0.15, "Equilíbrio entre valorização patrimonial e recebimento de proventos indexados ao CDI/IPCA."),
        ("Arrojado", 0.30, 0.45, 0.25, "Foco em maximização do Dividend Yield via papéis High Yield e fundos de oportunidade.")
    ]

    for idx, (perfil, tijolo, papel, hibrido, desc) in enumerate(dados_apoio, start=6):
        ws2.row_dimensions[idx].height = 24
        
        ws2[f"B{idx}"] = perfil
        ws2[f"B{idx}"].font = table_bold_font
        ws2[f"B{idx}"].alignment = align_center
        ws2[f"B{idx}"].border = cell_border
        
        ws2[f"C{idx}"] = tijolo
        ws2[f"C{idx}"].font = table_regular_font
        ws2[f"C{idx}"].alignment = align_center
        ws2[f"C{idx}"].number_format = "0.0%"
        ws2[f"C{idx}"].border = cell_border
        
        ws2[f"D{idx}"] = papel
        ws2[f"D{idx}"].font = table_regular_font
        ws2[f"D{idx}"].alignment = align_center
        ws2[f"D{idx}"].number_format = "0.0%"
        ws2[f"D{idx}"].border = cell_border
        
        ws2[f"E{idx}"] = hibrido
        ws2[f"E{idx}"].font = table_regular_font
        ws2[f"E{idx}"].alignment = align_center
        ws2[f"E{idx}"].number_format = "0.0%"
        ws2[f"E{idx}"].border = cell_border
        
        ws2[f"F{idx}"] = f"=SUM(C{idx}:E{idx})"
        ws2[f"F{idx}"].font = table_bold_font
        ws2[f"F{idx}"].alignment = align_center
        ws2[f"F{idx}"].number_format = "0.0%"
        ws2[f"F{idx}"].border = cell_border
        
        ws2[f"G{idx}"] = desc
        ws2[f"G{idx}"].font = table_regular_font
        ws2[f"G{idx}"].alignment = align_left
        ws2[f"G{idx}"].border = cell_border

    # Segunda Seção na Tabela de Apoio: Ativos de Referência e Exemplos de FIIs
    ws2.merge_cells("B11:G11")
    ws2["B11"] = "EXEMPLOS DE ATIVOS REAIS NO MERCADO BRASILEIRO (REFERÊNCIA EDUCACIONAL)"
    ws2["B11"].font = section_font
    ws2["B11"].fill = secondary_header_fill
    ws2["B11"].alignment = align_left
    ws2.row_dimensions[11].height = 24

    headers_fii_examples = [
        ("B12", "Ticker"),
        ("C12", "Segmento"),
        ("D12", "Tipo de Ativo"),
        ("E12", "Indexador Principal"),
        ("F12", "DY Médio Estimado"),
        ("G12", "Perfil Típico Recomendado")
    ]
    ws2.row_dimensions[12].height = 22
    for ref, title in headers_fii_examples:
        c = ws2[ref]
        c.value = title
        c.font = table_header_font
        c.fill = PatternFill(start_color="3B82F6", end_color="3B82F6", fill_type="solid")
        c.alignment = align_center
        c.border = cell_border

    exemplos_fiis = [
        ("HGLG11", "Logística / Galpões", "Tijolo", "IPCA / IGP-M", "0,75% a.m.", "Conservador / Moderado"),
        ("KNRI11", "Renda Mista (Lajes / Galpões)", "Tijolo", "IPCA / IGP-M", "0,72% a.m.", "Conservador"),
        ("XPML11", "Shopping Centers", "Tijolo", "IPCA / IGP-M", "0,80% a.m.", "Moderado"),
        ("KNCR11", "Recebíveis Imobiliários", "Papel", "CDI + Spread", "0,90% a.m.", "Conservador / Moderado"),
        ("MXRF11", "Recebíveis Imobiliários", "Papel / Híbrido", "CDI / IPCA", "0,92% a.m.", "Moderado / Arrojado"),
        ("CPTS11", "Recebíveis Imobiliários", "Papel", "IPCA+", "0,95% a.m.", "Arrojado"),
        ("KNSC11", "Recebíveis Imobiliários", "Papel", "IPCA / CDI", "0,88% a.m.", "Moderado")
    ]

    for idx, (ticker, seg, tipo, indexador, dy, rec) in enumerate(exemplos_fiis, start=13):
        ws2.row_dimensions[idx].height = 20
        is_even = (idx % 2 == 0)
        row_fill = zebra_even_fill if is_even else white_fill
        
        ws2[f"B{idx}"] = ticker
        ws2[f"B{idx}"].font = table_bold_font
        ws2[f"B{idx}"].alignment = align_center
        ws2[f"B{idx}"].border = cell_border
        ws2[f"B{idx}"].fill = row_fill
        
        ws2[f"C{idx}"] = seg
        ws2[f"C{idx}"].font = table_regular_font
        ws2[f"C{idx}"].alignment = align_left
        ws2[f"C{idx}"].border = cell_border
        ws2[f"C{idx}"].fill = row_fill
        
        ws2[f"D{idx}"] = tipo
        ws2[f"D{idx}"].font = table_bold_font
        ws2[f"D{idx}"].alignment = align_center
        ws2[f"D{idx}"].border = cell_border
        ws2[f"D{idx}"].fill = row_fill
        
        ws2[f"E{idx}"] = indexador
        ws2[f"E{idx}"].font = table_regular_font
        ws2[f"E{idx}"].alignment = align_center
        ws2[f"E{idx}"].border = cell_border
        ws2[f"E{idx}"].fill = row_fill
        
        ws2[f"F{idx}"] = dy
        ws2[f"F{idx}"].font = table_bold_font
        ws2[f"F{idx}"].alignment = align_center
        ws2[f"F{idx}"].border = cell_border
        ws2[f"F{idx}"].fill = row_fill
        
        ws2[f"G{idx}"] = rec
        ws2[f"G{idx}"].font = table_regular_font
        ws2[f"G{idx}"].alignment = align_left
        ws2[f"G{idx}"].border = cell_border
        ws2[f"G{idx}"].fill = row_fill

    # Larguras na aba 2
    col_widths_apoio = {
        "A": 3,
        "B": 22,
        "C": 22,
        "D": 22,
        "E": 24,
        "F": 14,
        "G": 46
    }
    for col_letter, width in col_widths_apoio.items():
        ws2.column_dimensions[col_letter].width = width

    # -------------------------------------------------------------
    # 3. ABA: Metodologia & Cálculos
    # -------------------------------------------------------------
    ws3 = wb.create_sheet(title="Metodologia & Cálculos")
    ws3.views.sheetView[0].showGridLines = True
    
    ws3.merge_cells("B2:F2")
    ws3["B2"] = "GUIA METODOLÓGICO | EXPLICAÇÃO DETALHADA DE CADA FÓRMULA"
    ws3["B2"].font = header_font
    ws3["B2"].fill = primary_header_fill
    ws3["B2"].alignment = align_center
    ws3.row_dimensions[2].height = 32

    ws3.merge_cells("B3:F3")
    ws3["B3"] = "Documentação técnica dos conceitos financeiros, parâmetros e equações do Excel implementados"
    ws3["B3"].font = title_sub_font
    ws3["B3"].fill = primary_header_fill
    ws3["B3"].alignment = align_center
    ws3.row_dimensions[3].height = 18

    calc_cards = [
        (5, "1. FUNÇÃO VF (VALOR FUTURO)", 
         "Sintaxe no Excel:", "=VF(taxa; nper; pgto; [vp]; [tipo])",
         "Onde:", "• taxa: Taxa de rentabilidade mensal esperada (Taxa_Mensal = 0,85% a.m.)\n• nper: Número total de períodos/meses (ex: 10 anos * 12 = 120 meses)\n• pgto: Aporte mensal fixo. ATENÇÃO: inserido com sinal negativo (-Aporte_Mensal) porque representa fluxo de saída do caixa para investimento.\n• vp: Patrimônio inicial investido (também negativo: -Patrimonio_Inicial).\n• tipo: 1 para aportes no início do período (antecipado) ou 0 para fim do período (postecipado). Utilizamos 1 conforme prática padrão de reinvestimento imediato."),
        
        (13, "2. FUNÇÃO PROCV (PESQUISA VERTICAL)",
         "Sintaxe no Excel:", "=PROCV(valor_procurado; matriz_tabela; núm_índice_coluna; [procurar_intervalo])",
         "Onde:", "• valor_procurado: O perfil selecionado pelo usuário (célula nomeada Perfil_Investidor: 'Conservador', 'Moderado' ou 'Arrojado').\n• matriz_tabela: O intervalo da tabela de apoio (Tabela_Perfis = 'Tabela de Apoio'!$B$5:$F$8).\n• núm_índice_coluna: Índice da coluna desejada: 2 para Tijolo, 3 para Papel e 4 para Híbridos.\n• procurar_intervalo: FALSO (0) para busca exata do nome do perfil."),
        
        (21, "3. CÁLCULO DOS JUROS COMPOSTOS E RENDIMENTOS ACUMULADOS",
         "Fórmula:", "= Patrimônio Acumulado (VF) - Total Aportado",
         "Onde:", "• Total Aportado: Aporte_Mensal * Meses + Patrimonio_Inicial\n• Juros Compostos: Revela o valor puro gerado pelos rendimentos reinvestidos.\n• % dos Juros no Patrimônio: Demonstra visualmente como, no longo prazo (20 a 30 anos), os juros passam a representar mais de 70% a 85% de todo o patrimônio do investidor!"),
         
        (29, "4. ESTIMATIVA DE DIVIDENDOS MENSAIS E O 'EFEITO BOLA DE NEVE'",
         "Fórmulas:", "Dividendos Mensais = Patrimônio Acumulado * Taxa_Mensal\nEfeito Bola de Neve = SE(Dividendos_Mensais >= Aporte_Mensal; 'ATIVADO'; 'Em Acumulação')",
         "Onde:", "• O Efeito Bola de Neve (ou Magic Number) ocorre no exato instante em que a renda passiva mensal gerada pela carteira é igual ou superior ao aporte mensal que o investidor tira do próprio bolso.\n• A partir desse marco (alcançado no 8º ano com 0,85% a.m.), os proventos compram mais cotas sozinhos do que o próprio investidor, acelerando a riqueza de forma exponencial.")
    ]

    for start_r, title, lbl_formula, formula_txt, lbl_desc, desc_txt in calc_cards:
        ws3.merge_cells(f"B{start_r}:F{start_r}")
        ws3[f"B{start_r}"] = title
        ws3[f"B{start_r}"].font = section_font
        ws3[f"B{start_r}"].fill = secondary_header_fill
        ws3[f"B{start_r}"].alignment = align_left
        ws3.row_dimensions[start_r].height = 22
        
        ws3[f"B{start_r+1}"] = lbl_formula
        ws3[f"B{start_r+1}"].font = bold_font
        ws3[f"B{start_r+1}"].alignment = align_left
        
        ws3.merge_cells(f"C{start_r+1}:F{start_r+1}")
        ws3[f"C{start_r+1}"] = formula_txt
        ws3[f"C{start_r+1}"].font = Font(name="Consolas", size=10, bold=True, color="1E3A8A")
        ws3[f"C{start_r+1}"].fill = input_highlight_fill
        ws3[f"C{start_r+1}"].alignment = align_left
        ws3.row_dimensions[start_r+1].height = 22
        
        ws3[f"B{start_r+2}"] = lbl_desc
        ws3[f"B{start_r+2}"].font = bold_font
        ws3[f"B{start_r+2}"].alignment = Alignment(horizontal="left", vertical="top")
        
        ws3.merge_cells(f"C{start_r+2}:F{start_r+5}")
        ws3[f"C{start_r+2}"] = desc_txt
        ws3[f"C{start_r+2}"].font = regular_font
        ws3[f"C{start_r+2}"].alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        
        for r in range(start_r, start_r+6):
            for c in ["B", "C", "D", "E", "F"]:
                ws3[f"{c}{r}"].border = cell_border

    col_widths_metod = {
        "A": 3,
        "B": 24,
        "C": 22,
        "D": 22,
        "E": 24,
        "F": 30
    }
    for col_letter, width in col_widths_metod.items():
        ws3.column_dimensions[col_letter].width = width

    # -------------------------------------------------------------
    # 4. REGISTRO DE INTERVALOS NOMEADOS (NAMED RANGES)
    # -------------------------------------------------------------
    # Intervalos nomeados do Passo 2 do Desafio:
    # Aporte_Mensal -> 'Simulador & Projeção'!$D$6
    # Patrimonio_Inicial -> 'Simulador & Projeção'!$D$7
    # Taxa_Mensal -> 'Simulador & Projeção'!$D$8
    # Taxa_Anual -> 'Simulador & Projeção'!$D$9
    # Perfil_Investidor -> 'Simulador & Projeção'!$D$10
    # Tabela_Perfis -> 'Tabela de Apoio'!$B$5:$F$8
    # Tabela_Projecao -> 'Simulador & Projeção'!$B$17:$J$32
    
    named_ranges = [
        ("Aporte_Mensal", "'Simulador & Projeção'!$D$6"),
        ("Patrimonio_Inicial", "'Simulador & Projeção'!$D$7"),
        ("Taxa_Mensal", "'Simulador & Projeção'!$D$8"),
        ("Taxa_Anual", "'Simulador & Projeção'!$D$9"),
        ("Perfil_Investidor", "'Simulador & Projeção'!$D$10"),
        ("Tabela_Perfis", "'Tabela de Apoio'!$B$5:$F$8"),
        ("Tabela_Projecao", "'Simulador & Projeção'!$B$17:$J$32"),
        ("Cenarios_Anos", "'Simulador & Projeção'!$C$18:$C$32")
    ]
    
    for name, target in named_ranges:
        dn = DefinedName(name, attr_text=target)
        wb.defined_names.add(dn)

    # -------------------------------------------------------------
    # 5. VALIDAÇÃO DE DADOS (DROPDOWN DE PERFIL)
    # -------------------------------------------------------------
    dv = DataValidation(type="list", formula1='"Conservador,Moderado,Arrojado"', allow_blank=False)
    dv.error ='Escolha um perfil válido na lista: Conservador, Moderado ou Arrojado.'
    dv.errorTitle = 'Perfil Inválido'
    dv.prompt = 'Selecione o perfil do investidor para recalcular as alocações via PROCV.'
    dv.promptTitle = 'Perfil de Investidor'
    ws1.add_data_validation(dv)
    dv.add(ws1["D10"])

    # Salvar planilha
    output_filename = "Simulador_Investimento_FIIs.xlsx"
    wb.save(output_filename)
    print(f"Planilha salva com sucesso em: {os.path.abspath(output_filename)}")

if __name__ == "__main__":
    build_fii_simulator()
