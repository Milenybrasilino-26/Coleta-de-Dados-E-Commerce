from formatar_excel import formatar_planilha_vendas, formatar_planilha_devolucao
import os
import xml.etree.ElementTree as ET
import pandas as pd
from datetime import datetime
import re
import time


print(" === MENU DE PROCESSAMENTO ===")
time.sleep(0.2)
print("1 - Criar uma nova planilha")
time.sleep(0.1)
print("2 - Atualizar histórico existente")
time.sleep(0.2)
opcao = input("Digite uma opcao: (1/2) ").strip()


notas_vendas_existentes = set()
notas_devolucoes_existentes = set()

df_base = pd.DataFrame()
df_devol_base = pd.DataFrame()

colunas = ["PLATAFORMA", "MÊS", "DATA", "SEMANA", "CLIENTE", "CPF/CNPJ", "NOTA", "ESTADO", "QTD", "SKU", "METRAGEM", "MODELO", "METROS", "VALOR", "STATUS"]

if opcao == "1":
    nome_arquivo = input("Digite o nome do arquivo: ")
    caminho = f"y:\\ADMINISTRATIVO\\Vanessa - Vendas do E-Commerce\\{nome_arquivo}.xlsx"
    df_base = pd.DataFrame(columns=colunas)
    print("Modo selecionado: Criando uma nova planilha do zero na memória...")
    
elif opcao == "2":
    pasta_historico = r"y:\ADMINISTRATIVO\Vanessa - Vendas do E-Commerce"
    
    # 1. Tenta listar os arquivos .xlsx da pasta de rede
    try:
        arquivos = [f for f in os.listdir(pasta_historico) if f.endswith('.xlsx')]
    except FileNotFoundError:
        print(f"Erro: O caminho '{pasta_historico}' não foi encontrado.")
        arquivos = []

    # 2. Se a pasta sumiu ou está vazia, cai direto na sua lógica de segurança alternativa
    if not arquivos:
        print("Nenhuma planilha encontrada na pasta de rede.")
        escolha_valida = False
    else:
        # Exibe os arquivos numerados para o usuário escolher
        print("\nSelecione a planilha que deseja atualizar:")
        for i, arquivo in enumerate(arquivos, start=1):
            print(f"{i} - {arquivo}")
        
        try:
            escolha = int(input("\nDigite o número da opção desejada: "))
            if 1 <= escolha <= len(arquivos):
                nome_arquivo_selecionado = arquivos[escolha - 1]
                caminho = os.path.join(pasta_historico, nome_arquivo_selecionado)
                escolha_valida = True
            else:
                print("Opção inválida! Número fora da lista.")
                escolha_valida = False
        except ValueError:
            print("Por favor, digite um número inteiro válido.")
            escolha_valida = False

    # 3. Se ele escolheu um arquivo válido na lista, roda a leitura principal
    if escolha_valida:
        try:
            df_base = pd.read_excel(caminho, sheet_name="Vendas")
        except: 
            df_base = pd.DataFrame()
        try:
            df_devol_base = pd.read_excel(caminho, sheet_name="Devoluções")
        except:
            df_devol_base = pd.DataFrame()
            
        print("\nModo selecionado: Atualizando histórico existente...")

        # Mapeia as notas antigas para ativar a trava de segurança
        if 'NOTA' in df_base.columns:
            notas_vendas_existentes = set(df_base['NOTA'].dropna().astype(str))
        if 'NOTA' in df_devol_base.columns:
            notas_devolucoes_existentes = set(df_devol_base['NOTA'].dropna().astype(str).unique())    
            
    # 4. Caso a escolha tenha falhado ou a pasta estivesse indisponível (Substitui o seu antigo 'else')
    else:
        print("\nAviso: Não foi possível carregar o arquivo da rede.")
        alteracao_local = input("Existe outro local para esse arquivo? (s/n): ").lower().strip()
        
        if alteracao_local in ["s", "sim"]:
            pasta_alternativa = input("Cole o caminho da pasta onde quer buscar (ex: C:/SuaPasta): ")
            nome_arquivo = input("Digite o nome do arquivo (sem .xlsx): ").strip()
            caminho = os.path.join(pasta_alternativa, f"{nome_arquivo}.xlsx")
            
            if os.path.exists(caminho):
                df_base = pd.read_excel(caminho)
                print("Modo selecionado: Atualizando histórico existente localmente...")
            else:
                print("Arquivo não encontrado no novo local. Criando novo do zero...")
                df_base = pd.DataFrame(columns=colunas)
        else:
            nome_arquivo = input("Digite o nome do arquivo para o OneDrive (sem .xlsx): ").strip()
            caminho = f"C:\\Users\\Windows\\OneDrive\\Documentos\\Technigas\\TECH\\DADOS\\{nome_arquivo}.xlsx"
            print(f"Caminho gerado: {caminho}")
            df_base = pd.DataFrame(columns=colunas)
            print("Modo selecionado: Criando uma nova planilha do zero na memória...")

pasta_raiz = r"Y:\ADMINISTRATIVO\\Vanessa - Vendas do E-Commerce\\Vendas 2026"
print("\nIniciando a leitura direta dos XMLs das pastas...")

dados_venda = []
dados_devolucoes = []


ns = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}

# Varrendo as pastas de XML
for raiz, pastas, arquivos in os.walk(pasta_raiz):
    for arquivo in arquivos:
        if arquivo.endswith('.xml'):
            caminho_completo_xml = os.path.join(raiz, arquivo)
            

            # Normaliza as barras para o Windows
            caminho_normalizado = os.path.normpath(raiz) 
            partes_do_caminho = caminho_normalizado.split(os.sep)

            # Estrutura: Plataforma/Pedidos/Mês
            if len(partes_do_caminho) >= 3 and partes_do_caminho[-2] == "Pedidos":
                mes = partes_do_caminho[-1]
                tipo = "Pedidos"
                plataforma = partes_do_caminho[-3]
                
                if plataforma.lower() == "full":
                    plataforma = "Mercado Livre Full"

            # Estrutura flexível para Devoluções (Alcança caminhos normais e subpastas como o Full)
            elif len(partes_do_caminho) >= 2 and any(p.lower().startswith("devolu") for p in partes_do_caminho):
                tipo = "Devoluções"
                
                if any("full" in p.lower() for p in partes_do_caminho):
                    plataforma = "Mercado Livre Full"
                else:
                    pasta_atual = partes_do_caminho[-1].lower()
                    if "cancel" in pasta_atual or partes_do_caminho[-2].lower() == "full":
                        plataforma = partes_do_caminho[-3] if len(partes_do_caminho) >= 3 else partes_do_caminho[-2]
                    else:
                        plataforma = partes_do_caminho[-2]
                
                # Define o mês atual como padrão (depois será corrigido pela data do XML)
                mes = datetime.now().strftime("%B")

            else:
                continue

            try:
                tree = ET.parse(caminho_completo_xml) 
                root = tree.getroot() 

                # 1. Dados Gerais da Nota
                ide = root.find('.//nfe:ide', ns)
                if ide is None:
                    continue


                tp_nf = ide.find('nfe:tpNF', ns).text if ide.find('nfe:tpNF', ns) is not None else ""


                numero_nota = ide.find('nfe:nNF', ns).text if ide is not None else ""

                # TRAVA DE SEGURANÇA DIÁRIA: Se a nota já existir no Excel antigo, pula o XML imediatamente
                if tipo == "Pedidos":
                    if str(numero_nota) in notas_vendas_existentes:
                        continue
                if tipo == "Devoluções":
                    if str(numero_nota) in notas_devolucoes_existentes:
                        continue

                # Tratamento da data e semana contínua do ano
                dh_emi_element = ide.find('nfe:dhEmi', ns)
                dh_emi_text = dh_emi_element.text[:10] if dh_emi_element is not None and dh_emi_element.text else ""
                data_final = ""
                semana_texto = ""

                if dh_emi_text:
                    try:
                        data_obj = datetime.strptime(dh_emi_text, "%Y-%m-%d")
                        data_final = data_obj.strftime("%d/%m/%Y")

                        dia_do_ano = data_obj.timetuple().tm_yday
                        dia_da_semana_ano_novo = datetime(data_obj.year, 1, 1).weekday()
                        num_semana_ano = (dia_do_ano + dia_da_semana_ano_novo - 1) // 7 + 1
                        
                        if num_semana_ano == 0:
                            num_semana_ano = 52
                            
                        semana_texto = f"Semana {num_semana_ano}"
                    except Exception:
                        data_final = dh_emi_text
                        semana_texto = ""

                # 2. Dados do Cliente
                dest = root.find('.//nfe:dest', ns)
                cliente = dest.find('nfe:xNome', ns).text if dest is not None and dest.find('nfe:xNome', ns) is not None else ""
                doc_cliente = ""
                if dest is not None:
                    cnpj_el = dest.find('nfe:CNPJ', ns)
                    cpf_el = dest.find('nfe:CPF', ns)
                    if cnpj_el is not None:
                        doc_cliente = cnpj_el.text
                        doc_cliente = (f"{doc_cliente[0:2]}.{doc_cliente[2:5]}.{doc_cliente[5:8]}/{doc_cliente[8:12]}-{doc_cliente[12:14]}")
                    elif cpf_el is not None:
                        doc_cliente = "CPF"

                ender_dest = root.find('.//nfe:dest/nfe:enderDest', ns)
                estado = ender_dest.find('nfe:UF', ns).text if ender_dest is not None and ender_dest.find('nfe:UF', ns) is not None else ""

                # 3. Itens da Nota
                detalhes_produtos = root.findall('.//nfe:det', ns)
                # Valor total da nota
                vnf_tag = root.find('.//nfe:ICMSTot/nfe:vNF', ns)

                valor_nf = (
                    float(vnf_tag.text)
                    if vnf_tag is not None
                    else 0
                )

                # Soma dos valores dos produtos
                soma_produtos = 0

                for item in detalhes_produtos:

                    prod_item = item.find('nfe:prod', ns)

                    if prod_item is None:
                        continue

                    vprod = prod_item.find('nfe:vProd', ns)

                    if vprod is not None:
                        soma_produtos += float(vprod.text)
                for prod_det in detalhes_produtos:
                    prod = prod_det.find('nfe:prod', ns)
                    if prod is None:
                        continue

                    #descricao_xml
                    descricao_xml = prod.find('nfe:xProd', ns).text if prod.find ('nfe:xProd', ns) is not None else ""
                    descricao_xml_min = str(descricao_xml).lower()

                    cod_xml = prod.find('nfe:cProd', ns).text if prod.find ('nfe:cProd', ns) is not None else ""
                    cod_xml_min = str(cod_xml).lower()

                    # EXTRAÇÃO DIRETA DO SKU DO XML (<cProd>)
                    sku_xml = prod.find('nfe:cProd', ns).text if prod.find('nfe:cProd', ns) is not None else ""
                    
                    qtd = float(prod.find('nfe:qCom', ns).text) if prod.find('nfe:qCom', ns) is not None else 0.0
                    
                    vprod = prod.find('nfe:vProd', ns)

                    valor_produto = (
                        float(vprod.text)
                        if vprod is not None
                        else 0
                    )

                    if soma_produtos > 0:

                        valor_total_item = round(
                            valor_nf * (valor_produto / soma_produtos),
                            2
                        )

                    else:
                        valor_total_item = 0

                    # Extrair o modelo do item
                    codigo_mang = ["mahy", "7118", "npod", "mangu", "05ha", "08la", 
                                    "5090", "5088", "lava", "pt500", "pt300", "r1at", "r2at", "maq"]

                    e_mang = False

                    # Teste de cada palavra da lista no SKU
                    for prefixo in codigo_mang:
                        if prefixo in cod_xml_min:
                            e_mang = True
                            break

                    #Definir o resultado final baseado no teste anterior
                    if e_mang:
                        modelo_final = "MANG"
                    else:
                        modelo_final = "OUTROS"


                    # Extração automática da Metragem baseada no texto do SKU bruto
                    metragem_extraida = ""
                    total_metros = ""
                    
                    padrao_metros = re.search(r'(\d+[\.,]?\d*)\s*(?:m|mt|mts|metro|metros)\b', descricao_xml_min)

                    if padrao_metros:
                        texto_numero = padrao_metros.group(1).replace(",", ".")
                        try:
                            metragem_extraida = float(texto_numero)
                        except ValueError:
                            metragem_extraida = ""
                    
                    if modelo_final == "OUTROS" and metragem_extraida == "":
                        metragem_extraida = qtd

                    if isinstance(metragem_extraida, (int, float)):
                        total_metros = round(qtd * metragem_extraida, 2)
                    else:
                        total_metros = qtd


                    # Monta a linha com o SKU bruto direto do XML
                    linha = {
                        "PLATAFORMA": plataforma,
                        "MÊS": mes.capitalize(),
                        "DATA": data_final,
                        "SEMANA": semana_texto,
                        "CLIENTE": cliente,
                        "CPF/CNPJ":doc_cliente,
                        "NOTA": int(numero_nota) if numero_nota.isdigit() else numero_nota,
                        "NOTA REFERÊNCIA": "0",
                        "ESTADO": estado,
                        "QTD": qtd,
                        "SKU": sku_xml,        
                        "DESCRIÇÃO": descricao_xml, 
                        "METRAGEM": metragem_extraida,   
                        "MODELO": modelo_final,               
                        "METROS": total_metros, 
                        "VALOR": valor_total_item,                        
                        "STATUS": ("Vendido") if tipo == "Pedidos" else "Devolvido",
                    }
                    if tipo == "Pedidos":
                        dados_venda.append(linha)
                    elif tipo == "Devoluções":
                        dados_devolucoes.append(linha)
            except Exception as e:
                print(f"Erro ao ler o arquivo {arquivo}: {e}")


# 4. SALVAMENTO COMPLETO INTEGRADO NO EXCEL
if dados_venda or dados_devolucoes:

    df_novos = pd.DataFrame(dados_venda)
    df_devol_novos = pd.DataFrame(dados_devolucoes)

    # Inicializa os DataFrames finais garantindo que herdem a estrutura caso a base esteja vazia
    df_final = df_novos if df_base.empty else pd.concat([df_base, df_novos], ignore_index=True)
    df_devol_final = df_devol_novos if df_devol_base.empty else pd.concat([df_devol_base, df_devol_novos], ignore_index=True)


    # ==================================================================
    # EXTRAÇÃO ULTRA ROBUSTA (IGNORANDO NAMESPACE DO XML)
    # ==================================================================
    nota_referenciada = "0"  # Valor padrão inicial
    texto_complementar = ""

    # Busca qualquer tag 'infCpl' no arquivo inteiro, não importa o prefixo/namespace
    for tag in root.iter():
        if tag.tag.endswith('infCpl'):
            if tag.text:
                texto_complementar = str(tag.text)
                break  # Achou o texto, pode parar de procurar no XML

    if texto_complementar:
        # Limpa quebras de linha que possam quebrar a leitura da regex
        texto_limpo = texto_complementar.replace('<br />', ' ').replace('<br>', ' ').replace('\n', ' ')
        
        # Mostra no terminal para conferência
        if tp_nf == 0:
            print(f"Nota {numero_nota} | Texto localizado: {texto_limpo[:15]}...")

        # Expressão regular ultra flexível: aceita com ou sem traço, maiúscula ou minúscula
        # Procura a palavra "referenciada" e pega os números logo à frente dela
        busca_nota_ref = re.search(r'referenciada[^\d]*(\d+)', texto_limpo, re.IGNORECASE)
        
        if busca_nota_ref:
            nota_referenciada = busca_nota_ref.group(1)
            print(f"   -> SUCESSO: Capturada Nota Origem: {nota_referenciada}")
    # ==================================================================


    #REGRA EXCLUSIVA: Remove a coluna "Nota referência" apenas da aba de Vendas
    colunas_para_remover_vendas = ["NOTA REFERÊNCIA"]
    if not df_final.empty:
        df_final = df_final.drop(columns=colunas_para_remover_vendas, errors="ignore")

    # ORDEM
    # DEVOLUÇÃO
    if not df_devol_final.empty and "DATA" in df_devol_final.columns:
        df_devol_final["DATA"] = pd.to_datetime(
            df_devol_final["DATA"],
            format="%d/%m/%Y",
            errors="coerce"
        )

        df_devol_final = df_devol_final.sort_values(
            by="DATA",
            ascending=True
        )
        meses = {
            1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril",
            5: "Maio", 6: "Junho", 7: "Julho", 8: "Agosto",
            9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro"
        }
        #Extrai o número do mês da coluna DATA e o .map() troca pelo nome em português
        df_devol_final["MÊS"] = df_devol_final["DATA"].dt.month.map(meses)

        # Preenche com um valor padrão caso a data seja inválida/vazia
        df_devol_final["MÊS"] = df_devol_final["MÊS"].fillna("Não identificado")


        df_devol_final["DATA"] = (
            df_devol_final["DATA"]
            .dt.strftime("%d/%m/%Y")
        )


    # Converter para data
    if not df_final.empty and "DATA" in df_final.columns:
        df_final["DATA"] = pd.to_datetime(
            df_final["DATA"],
            format="%d/%m/%Y",
            errors="coerce"
        )


    # Ordenar da mais antiga para a mais recente
    df_final = df_final.sort_values(
        by="DATA",
        ascending=True
    )

    # Voltar para formato brasileiro
    df_final["DATA"] = df_final["DATA"].dt.strftime("%d/%m/%Y")


    # PRIMEIRO SALVA
    print(f"Linhas vendas: {len(df_final)}")
    print(f"Linhas devoluções: {len(df_devol_final)}")

    pasta_destino = os.path.dirname(caminho)
    if pasta_destino and not os.path.exists(pasta_destino):
        os.makedirs(pasta_destino)

    with pd.ExcelWriter(
        caminho,
        engine="openpyxl"
    ) as writer:

        df_final.to_excel(
            writer,
            sheet_name="Vendas",
            index=False
        )

        df_devol_final.to_excel(
            writer,
            sheet_name="Devoluções",
            index=False
        )


    # DEPOIS FORMATA
    try:
        formatar_planilha_vendas(caminho)
        formatar_planilha_devolucao(caminho)
    except Exception as e:
        print(f"Aviso: O arquivo foi salvo, mas a formatação visual falhou: {e}")


    print(f"\nSucesso!")
    print(f"\nPlanilha de venda atualizada. {len(dados_venda)} novas linhas de itens foram adicionadas.")
    print(f"\nPlanilha de Devolução atualizada. {len(dados_devolucoes)} novas linhas de itens foram adicionadas.")
else:
    print("\nNenhum arquivo XML inédito foi localizado nas pastas.")
    formatar_planilha_vendas(caminho)
    formatar_planilha_devolucao(caminho)


if 'caminho' in locals() and os.path.exists(caminho):
    abrir = input("\nDeseja abrir o arquivo gerado/atualizado? (s/n): ").lower().strip()
    
    if abrir in ["s", "sim"]:
        print("Abrindo o arquivo no Excel...")
        os.startfile(caminho)
    else:
        time.sleep(0.1)
        print("Processo finalizado com sucesso!")
        time.sleep(0.5)
else:

    print("\nProcesso finalizado. Nenhum arquivo foi gerado ou alterado para abrir.")   

