import pandas as pd

# 1. Configurações iniciais - Ajuste os nomes dos arquivos e colunas aqui
ARQUIVO_SISTEMA = r"y:\ADMINISTRATIVO\Mileny\E-Commerce\Notas Emitidas no Sistema\relatorio_sistema.xlsx"
ARQUIVO_CODIGO = r"y:\ADMINISTRATIVO\Mileny\E-Commerce\Notas Emitidas no Sistema\relatorio_codigo.xlsx"

COLUNA_NF = "Nota Fiscal"
COLUNA_VALOR = "Total NF"

def validar_planilhas():
    try:
        df_sistema = pd.read_excel(ARQUIVO_SISTEMA)
        df_codigo = pd.read_excel(ARQUIVO_CODIGO)
    except Exception as e:
        print(f"Erro ao ler os arquivos: {e}")
        return

    # Padronizar os dados das notas fiscais para evitar erros de texto
    df_sistema[COLUNA_NF] = df_sistema[COLUNA_NF].astype(str).str.strip()
    df_codigo[COLUNA_NF] = df_codigo[COLUNA_NF].astype(str).str.strip()

    # Agrupa por NF e SOMA os valores dos itens
    df_codigo_somado = df_codigo.groupby(COLUNA_NF)[COLUNA_VALOR].sum().reset_index()
    df_sistema_somado = df_sistema.groupby(COLUNA_NF)[COLUNA_VALOR].sum().reset_index()

    # Dicionários para buscas rápidas {NF: Valor} nos dois sentidos
    dados_codigo = dict(zip(df_codigo_somado[COLUNA_NF], df_codigo_somado[COLUNA_VALOR]))
    dados_sistema = dict(zip(df_sistema_somado[COLUNA_NF], df_sistema_somado[COLUNA_VALOR]))

    # Listas para o resultado final
    deram_certo = []
    ausentes = []
    erradas = []
    excesso_codigo = []  # <--- Nova lista para notas exclusivas do seu código

    # 1. Comparação principal: usando o Sistema como base
    for _, linha in df_sistema_somado.iterrows():
        nf = linha[COLUNA_NF]
        valor_sistema = float(linha[COLUNA_VALOR])

        if nf not in dados_codigo:
            ausentes.append(nf)
        else:
            valor_codigo = float(dados_codigo[nf])
            if abs(round(valor_sistema, 2) - round(valor_codigo, 2)) <= 0.02:
                deram_certo.append(nf)
            else:
                erradas.append(f"NF: {nf} (Sistema: R${valor_sistema:.2f} vs Código: R${valor_codigo:.2f})")

    # 2. Verificação Reversa: notas que estão no Código mas não estão no Sistema
    for _, linha in df_codigo_somado.iterrows():
        nf = linha[COLUNA_NF]
        if nf not in dados_sistema:
            excesso_codigo.append(nf)

    # --- EXIBIÇÃO DO RESULTADO DIRETO ---
    print("\n" + "="*50)
    print("       RESULTADO DA VALIDAÇÃO COMPLETA")
    print("="*50)

    print(f"\n✅ DERAM CERTO ({len(deram_certo)}):")
    for nf in deram_certo:
        print(f"  • NF {nf} OK")

    print(f"\n❌ NÃO EXISTEM NO RELATÓRIO DE CÓDIGO (Falta no Código) ({len(ausentes)}):")
    for nf in ausentes:
        print(f"  • NF {nf} ausente no relatório do código")

    print(f"\n🔍 EXISTEM APENAS NO CÓDIGO (Sobrando no Código) ({len(excesso_codigo)}):")
    if excesso_codigo:
        for nf in excesso_codigo:
            print(f"  • NF {nf} está no código mas NÃO está no sistema")
    else:
        print("  • Nenhuma nota sobrando no código!")

    print(f"\n⚠️ VALOR ERRADO ({len(erradas)}):")
    for msg in erradas:
        print(f"  • {msg}")
    print("="*50)

if __name__ == "__main__":
    validar_planilhas()
