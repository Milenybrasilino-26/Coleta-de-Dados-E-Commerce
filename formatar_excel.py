import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation



def formatar_planilha_vendas(caminho):

    print(f"ESTILIZANDO O ARQUIVO")

    # Abrir planilha
    wb = openpyxl.load_workbook(caminho)
    if "Vendas" not in wb.sheetnames:
        print("Aba Vendas não encontrada")
        return

    ws = wb["Vendas"]

    # ===== ESTILOS =====

    fonte_cabecalho = Font(
        name="Calibri",
        size=11,
        bold=True,
        color="FFFFFF"
    )

    preenchimento_cabecalho = PatternFill(
        fill_type="solid",
        start_color="4F81BD",
        end_color="4F81BD"
    )

    fonte_dados = Font(
        name="Calibri",
        size=11
    )

    alinhamento_centro = Alignment(
        horizontal="center",
        vertical="center"
    )

    alinhamento_esquerda = Alignment(
        horizontal="left",
        vertical="center"
    )

    alinhamento_direita = Alignment(
        horizontal="right",
        vertical="center"
    )

    borda = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin")
    )

    # ===== CABEÇALHO =====

    for celula in ws[1]:
        celula.font = fonte_cabecalho
        celula.fill = preenchimento_cabecalho
        celula.alignment = alinhamento_centro
        celula.border = borda

    # ===== DADOS =====

    for linha in ws.iter_rows(min_row=2):

        for celula in linha:

            celula.font = fonte_dados
            celula.border = borda

    # ===== FORMATAR COLUNAS =====

    for coluna in ws.columns:

        maior = 0
        letra = get_column_letter(coluna[0].column)

        for celula in coluna:

            try:
                tamanho = len(str(celula.value))

                if tamanho > maior:
                    maior = tamanho

            except:
                pass

        ws.column_dimensions[letra].width = maior + 4

    # ===== FORMATOS ESPECÍFICOS =====

    for linha in range(2, ws.max_row + 1):

        # PLATAFORMA    
        ws[f"A{linha}"].alignment = alinhamento_esquerda

        # MES    
        ws[f"B{linha}"].alignment = alinhamento_centro

        # DATA
        ws[f"C{linha}"].alignment = alinhamento_centro

        # SEMANA
        ws[f"D{linha}"].alignment = alinhamento_centro

        # CLIENTE
        ws[f"E{linha}"].alignment = alinhamento_esquerda

        # CPF/CNPJ
        celula = ws[f"F{linha}"]
        celula.number_format = "@"
        celula.alignment = alinhamento_centro

        # NOTA
        ws[f"G{linha}"].alignment = alinhamento_centro
        
        # ESTADO
        ws[f"H{linha}"].alignment = alinhamento_centro

        # QTD
        celula = ws[f"I{linha}"]
        celula.alignment = alinhamento_centro
        celula.number_format = "#,##0"

        # PRODUTO
        ws[f"K{linha}"].alignment = alinhamento_esquerda

        # METRAGEM
        celula = ws[f"L{linha}"]
        celula.alignment = alinhamento_centro

        # MODELO
        ws[f"M{linha}"].alignment = alinhamento_centro

        # METROS
        celula = ws[f"N{linha}"]
        celula.alignment = alinhamento_centro

        # VALOR
        celula = ws[f"O{linha}"]
        celula.number_format = "R$ #,##0.00"
        celula.alignment = alinhamento_centro

        # STATUS
        ws[f"P{linha}"].alignment = alinhamento_centro


    # ===== FILTROS =====

    ws.auto_filter.ref = ws.dimensions

    # ===== CONGELAR CABEÇALHO =====

    ws.freeze_panes = "A2"

   # ===== COR =====

    preenchimento_erro = PatternFill(
    fill_type="solid",
    start_color="FFC7CE",
    end_color="FFC7CE"
    )

   # ===== FORMATAÇÃO CONDICIONAL =====

    regra_metragem = FormulaRule(
        formula=['ISBLANK(L2)'],
        fill=preenchimento_erro
    )

    ws.conditional_formatting.add(
        f'L2:L{ws.max_row}',
        regra_metragem
    )
    # ===== LISTA SUSPENSA STATUS =====
    status_dropdown = DataValidation(
        type="list",
        formula1='"Vendido,Devolvido,Cancelado"',
        allow_blank=True
    )
    ws.add_data_validation(status_dropdown)
    status_dropdown.add("P2:P50000")

    for linha in range(2, ws.max_row + 1):
        status_dropdown.add(ws[f"P{linha}"])

    # ==== COR ====
    
    verde = PatternFill(
    start_color="C6EFCE",
    end_color="C6EFCE",
    fill_type="solid"
    )

    vermelho = PatternFill(
        start_color="FFC7CE",
        end_color="FFC7CE",
        fill_type="solid"
    )

    amarelo = PatternFill(
        start_color="FFF5BA",
        end_color="FFF5BA",
        fill_type="solid"
    )

    ws.conditional_formatting.add(
        f"O2:O{ws.max_row}",
        FormulaRule(
            formula=['$P2="Vendido"'],
            fill=verde
        )
    )

    ws.conditional_formatting.add(
        f"O2:O{ws.max_row}",
        FormulaRule(
            formula=['$P2="Devolvido"'],
            fill=vermelho
        )
    )

    ws.conditional_formatting.add(
        f"O2:O{ws.max_row}",
        FormulaRule(
            formula=['$P2="Cancelado"'],
            fill=amarelo
        )
    )
    # ===== SALVAR =====

    wb.save(caminho)

    print("Planilha formatada com sucesso!")

def formatar_planilha_devolucao(caminho):
    print("FORMATANDO PLANILHA DE DEVOLUÇÃO!")

    wb = openpyxl.load_workbook(caminho)
    if "Devoluções" not in wb.sheetnames:
        print('Aviso: aba não encontrada!')
        return
    
    ws = wb["Devoluções"]

    # ESTILOS
    fonte_cabecalho = Font(
        name="Calibri",
        size=11,
        bold=True,
        color="FFFFFF"
    )

    # Preenchimento Vinho/Bordô para diferenciar visualmente de Vendas (Azul)
    preenchimento_cabecalho = PatternFill(
        fill_type="solid",
        start_color="9E2A2B",
        end_color="9E2A2B"
    )

    fonte_dados = Font(
        name="Calibri",
        size=11
    )

    alinhamento_centro = Alignment(
        horizontal="center",
        vertical="center"
    )

    alinhamento_esquerda = Alignment(
        horizontal="left",  # Corrigido para esquerda real
        vertical="center"
    )

    borda = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin")
    )
    
    # CABEÇALHO
    for celula in ws[1]:
        celula.font = fonte_cabecalho
        celula.fill = preenchimento_cabecalho  # Adicionado o preenchimento colorido
        celula.alignment = alinhamento_centro
        celula.border = borda

    # DADOS
    for linha in ws.iter_rows(min_row=2):
        for celula in linha:
            celula.font = fonte_dados
            celula.border = borda

    # FORMATAR LARGURA (Corrigido o erro de sintaxe 'column')
    for coluna in ws.columns:
        maior = 0
        letra = get_column_letter(coluna[0].column)
        for celula in coluna:  # Corrigido de 'column' para 'coluna'
            try:
                tamanho = len(str(celula.value))
                if tamanho > maior:
                    maior = tamanho
            except:
                pass
        ws.column_dimensions[letra].width = maior + 4

    # ===== MAPPING DINÂMICO DE FORMATOS ESPECÍFICOS =====
    # (Removido de dentro do laço de colunas para rodar apenas uma vez com recuo correto)
    colunas_index = {ws.cell(row=1, column=col).value: col for col in range(1, ws.max_column + 1)}

    for linha in range(2, ws.max_row + 1):
        for nome_coluna, col_idx in colunas_index.items():
            if not nome_coluna:
                continue
                
            celula = ws.cell(row=linha, column=col_idx)
            nome_limpo = str(nome_coluna).upper().strip()

            # Alinhamentos à esquerda
            if "CLIENTE" in nome_limpo or "NOME" in nome_limpo:
                celula.alignment = alinhamento_esquerda
                
            # Formatos numéricos e alinhamentos ao centro
            elif "VALOR" in nome_limpo or "PREÇO" in nome_limpo:
                celula.number_format = "R$ #,##0.00"
                celula.alignment = alinhamento_centro
                
            elif "QTD" in nome_limpo or "QUANTIDADE" in nome_limpo:
                celula.number_format = "#,##0"
                celula.alignment = alinhamento_centro
                
            elif "CPF" in nome_limpo or "CNPJ" in nome_limpo:
                celula.number_format = "@"
                celula.alignment = alinhamento_centro
            
            elif "SKU" in nome_limpo or "CODIGO" in nome_limpo:
                celula.number_format = "@"
                celula.alignment = alinhamento_esquerda

            if "PRODUTO" in nome_limpo or "DESCRIÇÃO" in nome_limpo:
                celula.alignment = alinhamento_esquerda
                
            # Padrão para códigos, notas, chaves, meses e datas
            else:
                celula.number_format = "@" 
                celula.alignment = alinhamento_centro

    # ===== FILTROS =====
    ws.auto_filter.ref = ws.dimensions

    # ===== CONGELAR CABEÇALHO =====
    ws.freeze_panes = "A2"
    
    # ===== SALVAR =====
    wb.save(caminho)
    print("FORMATAÇÃO CONCLUIDA COM SUCESSO!")
