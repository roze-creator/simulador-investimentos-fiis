import os
import matplotlib.pyplot as plt
import matplotlib.patches as patches
import numpy as np

os.makedirs("prints", exist_ok=True)

# Configurações globais de estilo
plt.rcParams['font.sans-serif'] = 'DejaVu Sans'
plt.rcParams['axes.edgecolor'] = '#CBD5E1'
plt.rcParams['axes.linewidth'] = 0.8

def fmt_br(val):
    return f"R$ {val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

# Cenários de anos
anos = [2, 3, 4, 5, 6, 7, 8, 9, 10, 12, 15, 18, 20, 25, 30]
aporte = 1000.0
taxa_m = 0.0085
patr_ini = 0.0

total_aportado = []
patr_acumulado = []
juros_acum = []
dividendos_m = []

for a in anos:
    nper = a * 12
    # FV formula antecipado: PMT * ((1+i)^n - 1)/i * (1+i)
    # ou postecipado: PMT * ((1+i)^n - 1)/i
    # openpyxl type 1 (antecipado):
    fv = aporte * (((1 + taxa_m)**nper - 1) / taxa_m) * (1 + taxa_m)
    tot_inv = patr_ini + aporte * nper
    juros = fv - tot_inv
    divs = fv * taxa_m
    
    total_aportado.append(tot_inv)
    patr_acumulado.append(fv)
    juros_acum.append(juros)
    dividendos_m.append(divs)

# ==============================================================================
# PRINT 1: DASHBOARD GERAL DO SIMULADOR (VISÃO CONSOLIDADA)
# ==============================================================================
fig = plt.figure(figsize=(14, 9), facecolor="#F8FAFC")
ax = fig.add_subplot(111)
ax.axis("off")

# Banner Superior
rect_banner = patches.FancyBboxPatch((0.02, 0.88), 0.96, 0.10, boxstyle="round,pad=0.01", facecolor="#1B365D", edgecolor="none")
ax.add_patch(rect_banner)
plt.text(0.50, 0.94, "SIMULADOR DE INVESTIMENTOS EM FUNDOS IMOBILIÁRIOS (FIIs)", fontsize=16, fontweight='bold', color="white", ha="center", va="center")
plt.text(0.50, 0.90, "Desafio de Projeto DIO | Modelagem com Função VF, PROCV, Intervalos Nomeados & Cenários (2 a 30 Anos)", fontsize=10, color="#CBD5E1", ha="center", va="center")

# Bloco 1: Parâmetros (Esquerda)
rect_p = patches.FancyBboxPatch((0.02, 0.52), 0.46, 0.33, boxstyle="round,pad=0.01", facecolor="white", edgecolor="#CBD5E1", linewidth=1.2)
ax.add_patch(rect_p)
rect_p_hdr = patches.FancyBboxPatch((0.02, 0.80), 0.46, 0.05, boxstyle="round,pad=0.005", facecolor="#2C5E8A", edgecolor="none")
ax.add_patch(rect_p_hdr)
plt.text(0.04, 0.825, "1. PARÂMETROS DO INVESTIDOR (INPUTS)", fontsize=11, fontweight='bold', color="white", va="center")

params_txt = [
    ("Aporte Mensal (R$):", "R$ 1.000,00", "#107C41"),
    ("Patrimônio Inicial (R$):", "R$ 0,00", "#334155"),
    ("Taxa Mensal Esperada (DY):", "0,85% a.m.", "#1B365D"),
    ("Taxa Anual Equivalente:", "10,69% a.a.", "#64748B"),
    ("Perfil do Investidor:", "Moderado (Dropdown)", "#2563EB")
]
for idx, (lbl, val, colr) in enumerate(params_txt):
    y = 0.76 - idx * 0.05
    plt.text(0.04, y, lbl, fontsize=10, fontweight='bold', color="#1E293B", va="center")
    plt.text(0.45, y, val, fontsize=10, fontweight='bold', color=colr, ha="right", va="center")

# Bloco 2: Alocação Dinâmica por Perfil com PROCV (Direita)
rect_a = patches.FancyBboxPatch((0.52, 0.52), 0.46, 0.33, boxstyle="round,pad=0.01", facecolor="white", edgecolor="#CBD5E1", linewidth=1.2)
ax.add_patch(rect_a)
rect_a_hdr = patches.FancyBboxPatch((0.52, 0.80), 0.46, 0.05, boxstyle="round,pad=0.005", facecolor="#2C5E8A", edgecolor="none")
ax.add_patch(rect_a_hdr)
plt.text(0.54, 0.825, "2. ALOCAÇÃO DO APORTE VIA PROCV (PERFIL MODERADO)", fontsize=11, fontweight='bold', color="white", va="center")

allocs = [
    ("FIIs de Tijolo (Logística / Shoppings)", "45,0%", "R$ 450,00", "=PROCV(Perfil; Tabela; 2; 0)"),
    ("FIIs de Papel (CRIs / Dívida Imob.)", "40,0%", "R$ 400,00", "=PROCV(Perfil; Tabela; 3; 0)"),
    ("FIIs Híbridos / FoFs / Fi-Infra", "15,0%", "R$ 150,00", "=PROCV(Perfil; Tabela; 4; 0)")
]
for idx, (cat, pct, v_alloc, form) in enumerate(allocs):
    y = 0.75 - idx * 0.065
    plt.text(0.54, y + 0.015, cat, fontsize=9.5, fontweight='bold', color="#1E293B")
    plt.text(0.54, y - 0.015, form, fontsize=8, color="#64748B", style="italic")
    plt.text(0.96, y + 0.015, f"{pct} ({v_alloc})", fontsize=10, fontweight='bold', color="#107C41", ha="right")

# Linha Total Alocado
plt.text(0.54, 0.55, "TOTAL ALOCADO:", fontsize=10, fontweight='bold', color="#1E293B")
plt.text(0.96, 0.55, "100,0%  (R$ 1.000,00)", fontsize=10, fontweight='bold', color="#107C41", ha="right")

# Bloco 3: Cards de Destaque / KPIs Executivos
kpi_data = [
    (0.02, 0.22, "PATRIMÔNIO EM 10 ANOS", fmt_br(patr_acumulado[8]), f"Renda: {fmt_br(dividendos_m[8])}/mês", "#1E293B", "#107C41"),
    (0.265, 0.22, "PATRIMÔNIO EM 20 ANOS", fmt_br(patr_acumulado[12]), f"Renda: {fmt_br(dividendos_m[12])}/mês", "#1E293B", "#107C41"),
    (0.51, 0.22, "PATRIMÔNIO EM 30 ANOS", fmt_br(patr_acumulado[14]), f"Renda: {fmt_br(dividendos_m[14])}/mês", "#107C41", "#047857"),
    (0.755, 0.22, "BOLA DE NEVE (MAGIC NUMBER)", "Alcançado no Ano 8 ❄️", "Renda supera o próprio aporte!", "#2563EB", "#1D4ED8")
]

for x, w, title, main_val, sub_val, main_c, sub_c in kpi_data:
    rect_kpi = patches.FancyBboxPatch((x, 0.33), 0.225, 0.15, boxstyle="round,pad=0.01", facecolor="white", edgecolor="#CBD5E1", linewidth=1.2)
    ax.add_patch(rect_kpi)
    rect_kpi_top = patches.FancyBboxPatch((x, 0.44), 0.225, 0.04, boxstyle="round,pad=0.005", facecolor="#F1F5F9", edgecolor="none")
    ax.add_patch(rect_kpi_top)
    plt.text(x + 0.112, 0.46, title, fontsize=8.5, fontweight='bold', color="#64748B", ha="center", va="center")
    plt.text(x + 0.112, 0.395, main_val, fontsize=12.5, fontweight='bold', color=main_c, ha="center", va="center")
    plt.text(x + 0.112, 0.35, sub_val, fontsize=8.5, fontweight='bold', color=sub_c, ha="center", va="center")

# Rodapé de resumo da tabela
rect_bot = patches.FancyBboxPatch((0.02, 0.03), 0.96, 0.27, boxstyle="round,pad=0.01", facecolor="white", edgecolor="#CBD5E1", linewidth=1.2)
ax.add_patch(rect_bot)
rect_bot_hdr = patches.FancyBboxPatch((0.02, 0.25), 0.96, 0.05, boxstyle="round,pad=0.005", facecolor="#1B365D", edgecolor="none")
ax.add_patch(rect_bot_hdr)
plt.text(0.04, 0.275, "3. SÍNTESE DA PROJEÇÃO DE CENÁRIOS COM A FUNÇÃO VF (2 A 30 ANOS)", fontsize=11, fontweight='bold', color="white", va="center")

table_summary_headers = ["Prazo", "Anos", "Meses", "Total Investido", "Patrimônio Acumulado (VF)", "Juros Compostos", "Dividendos Mensais", "Status Bola de Neve"]
xs = [0.04, 0.14, 0.21, 0.30, 0.45, 0.62, 0.77, 0.90]
for h_txt, x_pos in zip(table_summary_headers, xs):
    plt.text(x_pos, 0.22, h_txt, fontsize=8.5, fontweight='bold', color="#64748B", ha="left" if x_pos in (0.04, 0.90) else "center")

sample_indices = [0, 3, 6, 8, 12, 14] # 2 anos, 5 anos, 8 anos, 10 anos, 20 anos, 30 anos
fases = ["Curto Prazo", "Médio Prazo", "Bola de Neve ❄️", "Consolidação", "Maturidade", "Liberdade Plena"]

for row_i, (idx_s, f_name) in enumerate(zip(sample_indices, fases)):
    y_pos = 0.18 - row_i * 0.028
    a_val = anos[idx_s]
    m_val = a_val * 12
    inv_val = total_aportado[idx_s]
    pat_val = patr_acumulado[idx_s]
    jur_val = juros_acum[idx_s]
    div_val = dividendos_m[idx_s]
    st_val = "ATIVADO ❄️" if a_val >= 8 else "Em Acumulação"
    
    plt.text(xs[0], y_pos, f_name, fontsize=8, color="#1E293B", fontweight='bold' if a_val in (8, 30) else 'normal')
    plt.text(xs[1], y_pos, f"{a_val} anos", fontsize=8, color="#1E293B", ha="center")
    plt.text(xs[2], y_pos, f"{m_val} m", fontsize=8, color="#64748B", ha="center")
    plt.text(xs[3], y_pos, fmt_br(inv_val), fontsize=8, color="#334155", ha="center")
    plt.text(xs[4], y_pos, fmt_br(pat_val), fontsize=8.5, fontweight='bold', color="#107C41" if a_val >= 10 else "#1E293B", ha="center")
    plt.text(xs[5], y_pos, fmt_br(jur_val), fontsize=8, color="#2563EB", ha="center")
    plt.text(xs[6], y_pos, fmt_br(div_val), fontsize=8.5, fontweight='bold', color="#047857", ha="center")
    plt.text(xs[7], y_pos, st_val, fontsize=8, fontweight='bold', color="#15803D" if a_val >= 8 else "#64748B", ha="center")

plt.tight_layout()
plt.savefig("prints/01_dashboard_simulador.png", dpi=200, bbox_inches='tight')
plt.close()
print("Print 1 gerado: prints/01_dashboard_simulador.png")

# ==============================================================================
# PRINT 2: TABELA DE APOIO E ALOCAÇÃO POR PERFIL COM PROCV
# ==============================================================================
fig, (ax_tab, ax_pie) = plt.subplots(1, 2, figsize=(14, 7), facecolor="#F8FAFC", gridspec_kw={'width_ratios': [1.3, 1]})

ax_tab.axis("off")
rect_bg = patches.FancyBboxPatch((0.02, 0.05), 0.96, 0.90, boxstyle="round,pad=0.01", facecolor="white", edgecolor="#CBD5E1", linewidth=1.2)
ax_tab.add_patch(rect_bg)
rect_tab_hdr = patches.FancyBboxPatch((0.02, 0.85), 0.96, 0.10, boxstyle="round,pad=0.005", facecolor="#1B365D", edgecolor="none")
ax_tab.add_patch(rect_tab_hdr)

ax_tab.text(0.50, 0.91, "TABELA DE APOIO | MATRIZ DE ALOCAÇÃO POR PERFIL (PROCV)", fontsize=13, fontweight='bold', color="white", ha="center")
ax_tab.text(0.50, 0.87, "Base de dados consultada dinamicamente pela fórmula: =PROCV(Perfil_Investidor; Tabela_Perfis; col; FALSO)", fontsize=8.5, color="#E2E8F0", ha="center")

# Cabeçalhos da matriz
col_names = ["Perfil", "FIIs Tijolo", "FIIs Papel", "FIIs Híbridos", "DY Alvo", "Foco Estratégico"]
col_x = [0.06, 0.24, 0.40, 0.56, 0.72, 0.85]
for cn, cx in zip(col_names, col_x):
    ax_tab.text(cx, 0.79, cn, fontsize=9.5, fontweight='bold', color="#1B365D", ha="center" if cx != 0.85 else "left")

perfis_data = [
    ("Conservador", "60,0%", "30,0%", "10,0%", "0,75% a.m.", "Segurança e imóveis consolidados"),
    ("Moderado*", "45,0%", "40,0%", "15,0%", "0,85% a.m.", "Equilíbrio entre tijolo e CRIs"),
    ("Arrojado", "30,0%", "45,0%", "25,0%", "0,95% a.m.", "Maximização de yield via CRIs")
]

for idx, (p, t, pa, h, dy, foc) in enumerate(perfis_data):
    y = 0.70 - idx * 0.09
    is_mod = "*" in p
    if is_mod:
        rec_hl = patches.FancyBboxPatch((0.04, y - 0.03), 0.92, 0.075, boxstyle="round,pad=0.005", facecolor="#EFF6FF", edgecolor="#3B82F6", linewidth=1.5)
        ax_tab.add_patch(rec_hl)
    
    ax_tab.text(col_x[0], y, p, fontsize=10, fontweight='bold', color="#2563EB" if is_mod else "#1E293B", ha="center")
    ax_tab.text(col_x[1], y, t, fontsize=9.5, color="#1E293B", ha="center")
    ax_tab.text(col_x[2], y, pa, fontsize=9.5, color="#1E293B", ha="center")
    ax_tab.text(col_x[3], y, h, fontsize=9.5, color="#1E293B", ha="center")
    ax_tab.text(col_x[4], y, dy, fontsize=9.5, fontweight='bold', color="#107C41", ha="center")
    ax_tab.text(col_x[5], y, foc, fontsize=8.5, color="#64748B", style="italic")

# Box explicativo do PROCV
rec_box = patches.FancyBboxPatch((0.05, 0.10), 0.90, 0.28, boxstyle="round,pad=0.01", facecolor="#F8FAFC", edgecolor="#CBD5E1", linewidth=1)
ax_tab.add_patch(rec_box)
ax_tab.text(0.08, 0.33, "COMO O PROCV DIVIDE O APORTE AUTOMATICAMENTE:", fontsize=9.5, fontweight='bold', color="#1B365D")
ax_tab.text(0.08, 0.27, "• % Tijolo: =PROCV(Perfil_Investidor; 'Tabela de Apoio'!$B$5:$F$8; 2; FALSO)  → 45%", fontsize=8.5, color="#334155")
ax_tab.text(0.08, 0.22, "• % Papel: =PROCV(Perfil_Investidor; 'Tabela de Apoio'!$B$5:$F$8; 3; FALSO)  → 40%", fontsize=8.5, color="#334155")
ax_tab.text(0.08, 0.17, "• % Híbrido: =PROCV(Perfil_Investidor; 'Tabela de Apoio'!$B$5:$F$8; 4; FALSO) → 15%", fontsize=8.5, color="#334155")
ax_tab.text(0.08, 0.12, "• Aporte em R$: = Aporte_Mensal * % Alocação  (Ex: R$ 1.000 * 45% = R$ 450,00)", fontsize=8.5, fontweight='bold', color="#107C41")

# Gráfico de Rosca de Alocação (Perfil Selecionado: Moderado)
colors = ['#1B365D', '#2563EB', '#107C41']
sizes = [45, 40, 15]
labels = ['FIIs de Tijolo\n45% (R$ 450)', 'FIIs de Papel\n40% (R$ 400)', 'FIIs Híbridos\n15% (R$ 150)']

wedges, texts, autotexts = ax_pie.pie(
    sizes, labels=labels, autopct='%1.0f%%', startangle=140, colors=colors,
    textprops=dict(color="#1E293B", fontsize=10, fontweight='bold'),
    pctdistance=0.75, wedgeprops=dict(width=0.45, edgecolor='white', linewidth=2.5)
)
for at in autotexts:
    at.set_color('white')
    at.set_fontsize(11)
    at.set_fontweight('bold')

ax_pie.set_title("Distribuição do Aporte Mensal\n(Perfil Selecionado: Moderado)", fontsize=13, fontweight='bold', color="#1B365D", pad=20)

plt.tight_layout()
plt.savefig("prints/02_alocacao_perfil_procv.png", dpi=200, bbox_inches='tight')
plt.close()
print("Print 2 gerado: prints/02_alocacao_perfil_procv.png")

# ==============================================================================
# PRINT 3: TABELA DETALHADA DE CENÁRIOS DE 2 A 30 ANOS (FUNÇÃO VF)
# ==============================================================================
fig = plt.figure(figsize=(14, 8.5), facecolor="#F8FAFC")
ax = fig.add_subplot(111)
ax.axis("off")

# Banner
rect_b = patches.FancyBboxPatch((0.02, 0.90), 0.96, 0.08, boxstyle="round,pad=0.01", facecolor="#1B365D", edgecolor="none")
ax.add_patch(rect_b)
plt.text(0.50, 0.95, "PROJEÇÃO COMPLETA DE CENÁRIOS TEMPORAIS (2 A 30 ANOS) VIA FUNÇÃO VF", fontsize=15, fontweight='bold', color="white", ha="center")
plt.text(0.50, 0.92, "Fórmula Central: =VF(Taxa_Mensal; Meses; -Aporte_Mensal; -Patrimonio_Inicial; 1) | Aporte: R$ 1.000/mês | DY: 0,85% a.m.", fontsize=9, color="#CBD5E1", ha="center")

# Cabeçalhos da Tabela
headers_t3 = ["Cenário / Fase", "Anos", "Meses (NPER)", "Total Aportado", "Patrimônio Acumulado (VF)", "Juros Compostos", "% Juros", "Dividendos/Mês", "Status Bola de Neve"]
xs3 = [0.03, 0.17, 0.25, 0.35, 0.50, 0.65, 0.77, 0.86, 0.96]

rect_th = patches.FancyBboxPatch((0.02, 0.83), 0.96, 0.05, boxstyle="square,pad=0.0", facecolor="#2C5E8A", edgecolor="none")
ax.add_patch(rect_th)

for h_txt, x_pos in zip(headers_t3, xs3):
    plt.text(x_pos, 0.855, h_txt, fontsize=8.5, fontweight='bold', color="white", ha="left" if x_pos in (0.03,) else ("right" if x_pos in (0.96,) else "center"))

fases_todas = [
    ("Curto Prazo", 2),
    ("Curto Prazo", 3),
    ("Médio Prazo", 4),
    ("Médio Prazo", 5),
    ("Médio Prazo", 6),
    ("Médio Prazo", 7),
    ("Bola de Neve ❄️", 8),
    ("Médio Prazo", 9),
    ("Consolidação", 10),
    ("Consolidação", 12),
    ("Longo Prazo", 15),
    ("Longo Prazo", 18),
    ("Maturidade", 20),
    ("Independência", 25),
    ("Liberdade Plena", 30)
]

for row_idx, (fase, a_num) in enumerate(fases_todas):
    y = 0.79 - row_idx * 0.048
    i_sc = anos.index(a_num)
    m_num = a_num * 12
    ap_val = total_aportado[i_sc]
    pat_val = patr_acumulado[i_sc]
    jur_val = juros_acum[i_sc]
    pct_j = (jur_val / pat_val) * 100
    div_val = dividendos_m[i_sc]
    is_snow = a_num >= 8
    
    # Zebra
    bg_c = "#EFF6FF" if a_num in (8, 10, 20, 30) else ("#F8FAFC" if row_idx % 2 == 0 else "white")
    rec_r = patches.Rectangle((0.02, y - 0.015), 0.96, 0.045, facecolor=bg_c, edgecolor="#E2E8F0", linewidth=0.6)
    ax.add_patch(rec_r)
    
    bold_flag = 'bold' if a_num in (8, 10, 20, 30) else 'normal'
    plt.text(xs3[0], y, fase, fontsize=8.5, color="#1E293B", fontweight=bold_flag)
    plt.text(xs3[1], y, f"{a_num} anos", fontsize=8.5, color="#1E293B", ha="center")
    plt.text(xs3[2], y, f"{m_num}", fontsize=8.5, color="#64748B", ha="center")
    plt.text(xs3[3], y, fmt_br(ap_val), fontsize=8.5, color="#334155", ha="center")
    plt.text(xs3[4], y, fmt_br(pat_val), fontsize=8.5, color="#107C41" if a_num >= 10 else "#1E293B", fontweight='bold', ha="center")
    plt.text(xs3[5], y, fmt_br(jur_val), fontsize=8.5, color="#2563EB", ha="center")
    plt.text(xs3[6], y, f"{pct_j:.1f}%", fontsize=8.5, color="#D97706" if pct_j > 50 else "#64748B", fontweight='bold' if pct_j > 50 else 'normal', ha="center")
    plt.text(xs3[7], y, fmt_br(div_val), fontsize=8.5, color="#047857", fontweight='bold', ha="center")
    plt.text(xs3[8], y, "ATIVADO ❄️" if is_snow else "Acumulando", fontsize=8.5, color="#15803D" if is_snow else "#94A3B8", fontweight='bold', ha="right")

# Rodapé explicativo
plt.text(0.50, 0.03, "Nota: Observe a virada exponencial — A partir do 8º ano (Magic Number), a renda mensal supera o aporte e, no 30º ano, 80,7% do patrimônio é fruto de juros compostos!", fontsize=8.5, color="#475569", ha="center", style="italic")

plt.tight_layout()
plt.savefig("prints/03_projecao_cenarios_vf.png", dpi=200, bbox_inches='tight')
plt.close()
print("Print 3 gerado: prints/03_projecao_cenarios_vf.png")

# ==============================================================================
# PRINT 4: GRÁFICO DE EVOLUÇÃO PATRIMONIAL (JUROS COMPOSTOS X TOTAL APORTADO)
# ==============================================================================
fig, (ax_growth, ax_div) = plt.subplots(1, 2, figsize=(14, 6.5), facecolor="#F8FAFC")

# Gráfico 1: Evolução Patrimonial (Total Aportado vs Juros)
ax_growth.plot(anos, np.array(total_aportado)/1000, label="Total Aportado do Bolso (R$ mil)", color="#64748B", linewidth=2.5, linestyle="--")
ax_growth.plot(anos, np.array(patr_acumulado)/1000, label="Patrimônio Acumulado Total (VF)", color="#107C41", linewidth=3.2)
ax_growth.fill_between(anos, np.array(total_aportado)/1000, np.array(patr_acumulado)/1000, color="#22C55E", alpha=0.18, label="Juros Compostos Reinvestidos")

ax_growth.set_title("Evolução Patrimonial: Aportes vs. Juros Compostos", fontsize=12, fontweight='bold', color="#1B365D", pad=12)
ax_growth.set_xlabel("Prazo do Investimento (Anos)", fontsize=10, fontweight='bold', color="#334155")
ax_growth.set_ylabel("Valor Acumulado (R$ Milhares)", fontsize=10, fontweight='bold', color="#334155")
ax_growth.grid(True, linestyle=":", alpha=0.6)
ax_growth.legend(loc="upper left", frameon=True, facecolor="white")

# Ponto de destaque no Ano 30
ax_growth.scatter([30], [patr_acumulado[-1]/1000], color="#107C41", s=80, zorder=5)
ax_growth.annotate(f"30 Anos:\n{fmt_br(patr_acumulado[-1])}", (30, patr_acumulado[-1]/1000), textcoords="offset points", xytext=(-80,-15),
                   bbox=dict(boxstyle="round,pad=0.3", fc="#ECFDF5", ec="#107C41", lw=1.2), fontsize=8.5, fontweight='bold', color="#047857")

# Gráfico 2: Renda Passiva Mensal (Dividendos) e Efeito Bola de Neve
ax_div.plot(anos, dividendos_m, color="#2563EB", linewidth=3, marker='o', markersize=5, label="Dividendos Mensais (R$)")
ax_div.axhline(aporte, color="#EF4444", linestyle="--", linewidth=2, label="Aporte Mensal (R$ 1.000,00)")

# Marcação do Efeito Bola de Neve no Ano 8
ax_div.scatter([8], [dividendos_m[6]], color="#D97706", s=100, zorder=5)
ax_div.annotate(f"Efeito Bola de Neve!\nAno 8: {fmt_br(dividendos_m[6])}/mês", (8, dividendos_m[6]), textcoords="offset points", xytext=(-40, 25),
                bbox=dict(boxstyle="round,pad=0.3", fc="#FEF3C7", ec="#D97706", lw=1.2), fontsize=8.5, fontweight='bold', color="#92400E",
                arrowprops=dict(arrowstyle="->", color="#D97706", lw=1.2))

ax_div.set_title("Crescimento dos Dividendos & Efeito Bola de Neve", fontsize=12, fontweight='bold', color="#1B365D", pad=12)
ax_div.set_xlabel("Prazo do Investimento (Anos)", fontsize=10, fontweight='bold', color="#334155")
ax_div.set_ylabel("Renda Mensal de Dividendos (R$)", fontsize=10, fontweight='bold', color="#334155")
ax_div.grid(True, linestyle=":", alpha=0.6)
ax_div.legend(loc="upper left", frameon=True, facecolor="white")

plt.tight_layout()
plt.savefig("prints/04_grafico_evolucao_patrimonial.png", dpi=200, bbox_inches='tight')
plt.close()
print("Print 4 gerado: prints/04_grafico_evolucao_patrimonial.png")
