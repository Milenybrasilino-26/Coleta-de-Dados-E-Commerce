import os
import openpyxl
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
from datetime import datetime

# ==============================================================================
# CONFIGURAÇÃO DOS CAMINHOS (Substitua pelo seu caminho real)
# ==============================================================================
caminho_dados_originais = r"Y:\ADMINISTRATIVO\Vanessa - Vendas do E-Commerce\Coleta de Dados.xlsx"
caminho_novo_dashboard  = r"Y:\ADMINISTRATIVO\Vanessa - Vendas do E-Commerce\Dashboard_Analise.xlsx"

def criar_dashboard():
    print("Lendo dados do arquivo original...")
    
    if not os.path.exists(caminho_dados_originais):
        print(f"Erro: Arquivo original não encontrado em: {caminho_dados_originais}")
        return

    # 1. Carrega as duas abas para processamento das métricas
    df_vendas = pd.read_excel(caminho_dados_originais, sheet_name="Vendas")
    df_devol = pd.read_excel(caminho_dados_originais, sheet_name="Devoluções")

    # Normaliza o nome das colunas para evitar erros de maiúsculas/espaços
    df_vendas.columns = [str(c).upper().strip() for c in df_vendas.columns]
    df_devol.columns = [str(c).upper().strip() for c in df_devol.columns]

    # Ordem cronológica padrão para o Dashboard
    ordem_meses = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", 
                   "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]

    # ==============================================================================
    # PROCESSAMENTO DE MÉTRICAS COM PANDAS
    # ==============================================================================
    print("Processando indicadores...")

    # Métrica 1: Vendas vs Devoluções por Mês e Plataforma
    vendas_mes_plat = df_vendas.groupby(["MÊS", "PLATAFORMA"])["VALOR"].sum().reset_index()
    devol_mes_plat = df_devol.groupby(["MÊS", "PLATAFORMA"])["VALOR"].sum().reset_index()

    resumo_geral = pd.merge(vendas_mes_plat, devol_mes_plat, on=["MÊS", "PLATAFORMA"], how="outer", suffixes=("_VENDAS", "_DEVOL")).fillna(0)
    resumo_geral["MÊS"] = pd.Categorical(resumo_geral["MÊS"], categories=ordem_meses, ordered=True)
    resumo_geral = resumo_geral.sort_values(["MÊS", "PLATAFORMA"]).reset_index(drop=True)

    # Métrica 2: Vendas específicas do item Mangueiras
    col_modelo = "MODELO" if "MODELO" in df_vendas.columns else "PRODUTO"
    df_mangueiras = df_vendas[df_vendas[col_modelo].str.contains("mangueira", case=False, na=False)]
    vendas_mangueiras = df_mangueiras.groupby(["MÊS", "PLATAFORMA"])["VALOR"].sum().reset_index()
    vendas_mangueiras["MÊS"] = pd.Categorical(vendas_mangueiras["MÊS"], categories=ordem_meses, ordered=True)
    vendas_mangueiras = vendas_mangueiras.sort_values(["MÊS", "PLATAFORMA"]).reset_index(drop=True)

    # Métrica 3: Estados (UF) que mais vendem
    col_estado = "ESTADO" if "ESTADO" in df_vendas.columns else "UF"
    top_estados = df_vendas.groupby(col_estado)["VALOR"].sum().reset_index().sort_values(by="VALOR", ascending=False).head(10)

    # Métrica 4: Semanas que mais venderam
    top_semanas = df_vendas.groupby("SEMANA")["VALOR"].sum().reset_index().sort_values(by="VALOR", ascending=False).head(10)

    # ==============================================================================
    # MONTAGEM VISUAL DA PLANILHA DO DASHBOARD (OPENPYXL)
    # ==============================================================================
    print("Montando o arquivo visual do Dashboard...")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Dashboard Geral"
    ws.views.sheetView[0].showGridLines = True

    # Estilos de Fontes e Cores (Tema Azul Corporativo Escuro)
    f_titulo = Font(name="Calibri", size=16, bold=True, color="FFFFFF")
    f_secao = Font(name="Calibri", size=12, bold=True, color="1F497D")
    f_cabecalho = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    f_dados = Font(name="Calibri", size=10)
    f_total = Font(name="Calibri", size=10, bold=True)

    fill_titulo = PatternFill(fill_type="solid", start_color="1F497D", end_color="1F497D")
    fill_cabecalho = PatternFill(fill_type="solid", start_color="366092", end_color="366092")
    fill_totais = PatternFill(fill_type="solid", start_color="DCE6F1", end_color="DCE6F1")
    borda_fina = Border(
        left=Side(style="thin", color="D9D9D9"), right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"), bottom=Side(style="thin", color="D9D9D9")
    )

    # Banner Principal do Topo
    ws.merge_cells("B2:P2")
    ws["B2"] = "PAINEL COMPLETO DE DESEMPENHO E ESTATÍSTICAS COMERCIAIS"
    ws["B2"].font = f_titulo
    ws["B2"].fill = fill_titulo
    ws["B2"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 40

    # --- TABELA 1: COMPARATIVO MENSAL VENDAS VS DEVOLUÇÕES ---
    ws["B4"] = "Comparativo Mensal de Canais e Devoluções"
    ws["B4"].font = f_secao

    headers_t1 = ["Mês", "Plataforma", "Faturamento Bruto", "Total Devolvido"]
    for i, h in enumerate(headers_t1, start=2):
        c = ws.cell(row=5, column=i, value=h)
        c.font = f_cabecalho
        c.fill = fill_cabecalho
        c.alignment = Alignment(horizontal="center")
        c.border = borda_fina

    r_atual = 6
    for idx, row in resumo_geral.iterrows():
        ws.cell(row=r_atual, column=2, value=str(row["MÊS"])).alignment = Alignment(horizontal="center")
        ws.cell(row=r_atual, column=3, value=row["PLATAFORMA"]).alignment = Alignment(horizontal="left")

        c_v = ws.cell(row=r_atual, column=4, value=row["VALOR_VENDAS"])
        c_v.number_format = "R$ #,##0.00"
        c_v.alignment = Alignment(horizontal="right")

        c_d = ws.cell(row=r_atual, column=5, value=row["VALOR_DEVOL"])
        c_d.number_format = "R$ #,##0.00"
        c_d.alignment = Alignment(horizontal="right")

        for col in range(2, 6):
            ws.cell(row=r_atual, column=col).font = f_dados
            ws.cell(row=r_atual, column=col).border = borda_fina
        r_atual += 1

    # Totais da Tabela 1
    ws.cell(row=r_atual, column=2, value="TOTAL GERAL").font = f_total
    ws.cell(row=r_atual, column=2).border = borda_fina
    ws.cell(row=r_atual, column=3).border = borda_fina
    for col in range(4, 6):
        letra = get_column_letter(col)
        tot_c = ws.cell(row=r_atual, column=col, value=f"=SUM({letra}6:{letra}{r_atual-1})")
        tot_c.number_format = "R$ #,##0.00"
        tot_c.font = f_total
        tot_c.fill = fill_totais
        tot_c.border = borda_fina

    # --- TABELA 2: APENAS VENDAS DE MANGUEIRAS ---
    ws["G4"] = "Análise Focada: Vendas de Mangueiras"
    ws["G4"].font = f_secao

    headers_t2 = ["Mês", "Plataforma", "Total Mangueiras"]
    for i, h in enumerate(headers_t2, start=7):
        c = ws.cell(row=5, column=i, value=h)
        c.font = f_cabecalho
        c.fill = fill_cabecalho
        c.alignment = Alignment(horizontal="center")
        c.border = borda_fina

    r_m = 6
    for idx, row in vendas_mangueiras.iterrows():
        ws.cell(row=r_m, column=7, value=str(row["MÊS"])).alignment = Alignment(horizontal="center")
        ws.cell(row=r_m, column=8, value=row["PLATAFORMA"]).alignment = Alignment(horizontal="left")

        c_v = ws.cell(row=r_m, column=9, value=row["VALOR"])
        c_v.number_format = "R$ #,##0.00"
        c_v.alignment = Alignment(horizontal="right")

        for col in range(7, 10):
            ws.cell(row=r_m, column=col).font = f_dados
            ws.cell(row=r_m, column=col).border = borda_fina
        r_m += 1

    # Totais da Tabela 2
    ws.cell(row=r_m, column=7, value="TOTAL MANGUEIRAS").font = f_total
    ws.cell(row=r_m, column=7).border = borda_fina
    ws.cell(row=r_m, column=8).border = borda_fina
    tot_m = ws.cell(row=r_m, column=9, value=f"=SUM(I6:I{r_m-1})")
    tot_m.number_format = "R$ #,##0.00"
    tot_m.font = f_total
    tot_m.fill = fill_totais
    tot_m.border = borda_fina

    # --- TABELA 3: TOP 10 ESTADOS ---
    ws["K4"] = "Top 10 Estados (Faturamento)"
    ws["K4"].font = f_secao

    headers_t3 = ["UF", "Faturamento"]
    for i, h in enumerate(headers_t3, start=11):
        c = ws.cell(row=5, column=i, value=h)
        c.font = f_cabecalho
        c.fill = fill_cabecalho
        c.alignment = Alignment(horizontal="center")
        c.border = borda_fina

    r_e = 6
    for idx, row in top_estados.iterrows():
        ws.cell(row=r_e, column=11, value=row[col_estado]).alignment = Alignment(horizontal="center")
        c_v = ws.cell(row=r_e, column=12, value=row["VALOR"])
        c_v.number_format = "R$ #,##0.00"
        c_v.alignment = Alignment(horizontal="right")

        ws.cell(row=r_e, column=11).font = f_dados
        ws.cell(row=r_e, column=11).border = borda_fina
        ws.cell(row=r_e, column=12).font = f_dados
        ws.cell(row=r_e, column=12).border = borda_fina
        r_e += 1

    # --- TABELA 4: TOP 10 SEMANAS ---
    ws["N4"] = "Top 10 Semanas Campeãs"
    ws["N4"].font = f_secao

    headers_t4 = ["Semana", "Faturamento"]
    for i, h in enumerate(headers_t4, start=14):
        c = ws.cell(row=5, column=i, value=h)
        c.font = f_cabecalho
        c.fill = fill_cabecalho
        c.alignment = Alignment(horizontal="center")
        c.border = borda_fina

    r_s = 6
    for idx, row in top_semanas.iterrows():
        ws.cell(row=r_s, column=14, value=row["SEMANA"]).alignment = Alignment(horizontal="center")
        c_v = ws.cell(row=r_s, column=15, value=row["VALOR"])
        c_v.number_format = "R$ #,##0.00"
        c_v.alignment = Alignment(horizontal="right")

        ws.cell(row=r_s, column=14).font = f_dados
        ws.cell(row=r_s, column=14).border = borda_fina
        ws.cell(row=r_s, column=15).font = f_dados
        ws.cell(row=r_s, column=15).border = borda_fina
        r_s += 1

    # ADICIONANDO O GRÁFICO DE BARRAS NATIVO 
    chart = BarChart()
    chart.type = "col"
    chart.style = 10
    chart.title = "Comparativo: Fatimento Bruto vs Devoluções"
    chart.y_axis.title = "Valores (R$)"
    chart.x_axis.title = "Meses/Canais"

    # Referências para colunas "Faturamento Bruto" e "Total Devolvido" da primeira tabela
    data_ref = Reference(ws, min_col=4, min_row=5, max_col=5, max_row = r_atual-1)
    cats_ref = Reference(ws, min_col=2, min_row=6, max_row= r_atual-1)

    chart.add_data(data_ref, titles_from_data = True)
    chart.set_categories(cats_ref)
    chart.height = 14
    chart.width = 25

    # Posiciona o gráfico debaixo do bloco de tabelas
    linha_grafico = max(r_atual, r_m, r_e, r_s) + 3
    ws.add_chart(chart, f"B{linha_grafico}"
    )

#Ajusta as larguras de coluna automaticamente
    for col in ws.columns:
        maior = max(len(str(celula.value or '')) for celula in col)
        letra = get_column_letter(col[0].column)
        if letra != "A":  # Corrigido: adicionado os dois pontos
            ws.column_dimensions[letra].width = max(maior + 3, 13)  # Corrigido: identado para dentro do IF

    # Cria a pasta de destino caso não exista e salva o dashboard
    pasta_destino = os.path.dirname(caminho_novo_dashboard)
    if pasta_destino and not os.path.exists(pasta_destino):
        os.makedirs(pasta_destino)

    wb.save(caminho_novo_dashboard)
    print(f"Sucesso! Planilha de teste gerada em: {caminho_novo_dashboard}")

# Execução padrão do arquivo (Corrigido os dunders __)
if __name__ == "__main__":
    criar_dashboard()