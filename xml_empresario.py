import os
import re
import xml.etree.ElementTree as ET
from datetime import datetime
import pandas as pd
from pathlib import Path

# ==============================================================================
# CONFIGURAÇÃO DOS CAMINHOS PARA O TESTE ISOLADO
# ==============================================================================
pasta_sistema_externo = r"y:\ADMINISTRATIVO\Mileny\E-Commerce\Notas Emitidas no Sistema\NFs"
caminho_planilha_teste = r"y:\ADMINISTRATIVO\Mileny\E-Commerce\Notas Emitidas no Sistema\Teste_Coleta_Sistema_Externo.xlsx"

print("=== INICIANDO ARQUIVO DE TESTE ISOLADO - SISTEMA EXTERNO ===")

# Definição de colunas específicas para cada aba (Nota Referência exclusiva nas Devoluções/Canceladas)
colunas_pedidos = [
    "PLATAFORMA", "MÊS", "DATA", "SEMANA", "CLIENTE", "CPF/CNPJ", "NOTA", 
    "ESTADO", "QTD", "SKU", "DESCRIÇÃO", "VALOR", "STATUS", "MOTIVO", "TIPO OPERAÇÃO"
]

colunas_cancelados_devolucoes = [
    "PLATAFORMA", "MÊS", "DATA", "SEMANA", "CLIENTE", "CPF/CNPJ", "NOTA", 
    "NOTA REFERÊNCIA", "ESTADO", "QTD", "SKU", "DESCRIÇÃO", "VALOR", "STATUS", "MOTIVO", "TIPO OPERAÇÃO"
]

# ==============================================================================
# PASSAGEM 1: MAPEAR EVENTOS DE CANCELAMENTO ANTES DE PROCESSAR AS NOTAS
# ==============================================================================
notas_canceladas_no_lote = {}
ns = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}

if not os.path.exists(pasta_sistema_externo):
    print(f"[ERRO] A pasta de XMLs configurada não existe: {pasta_sistema_externo}")
    exit()

print("Passo 1: Mapeando eventos de cancelamento homologados (Tags SEFAZ)...")
for raiz, pastas, arquivos in os.walk(pasta_sistema_externo):
    for arquivo in arquivos:
        if arquivo.endswith('.xml'):
            try:
                tree = ET.parse(os.path.join(raiz, arquivo))
                root = tree.getroot()
                
                desc_evento = root.find('.//nfe:descEvento', ns)
                desc_texto = desc_evento.text.lower() if desc_evento is not None and desc_evento.text else ""
                
                if "cancel" in root.tag.lower() or "cancelamento" in desc_texto:
                    ch_nfe = root.find('.//nfe:chNFe', ns)
                    justificativa = root.find('.//nfe:xJust', ns)
                    
                    if ch_nfe is not None and len(ch_nfe.text) == 44:
                        num_nota_canc = ch_nfe.text[25:34]
                        num_nota_inteiro = int(num_nota_canc) if num_nota_canc.isdigit() else num_nota_canc
                        motivo_texto = justificativa.text if justificativa is not None else "Cancelado via evento SEFAZ"
                        
                        notas_canceladas_no_lote[str(num_nota_inteiro)] = motivo_texto
            except: pass

print(f"-> Notas canceladas identificadas no lote: {len(notas_canceladas_no_lote)}")

# ==============================================================================
# PASSAGEM 2: PROCESSAMENTO DAS NOTAS E ITENS
# ==============================================================================
dados_venda = []
dados_devolucoes = []

print("\nPasso 2: Processando os itens internos de cada XML...")
for raiz, pastas, arquivos in os.walk(pasta_sistema_externo):
    for arquivo in arquivos:
        if not arquivo.endswith('.xml'):
            continue

        try:
            tree = ET.parse(os.path.join(raiz, arquivo))
            root = tree.getroot()

            ide = root.find('.//nfe:ide', ns)
            if ide is None: 
                continue

            tp_nf = ide.find('nfe:tpNF', ns).text if ide.find('nfe:tpNF', ns) is not None else "1"
            nat_op = ide.find('nfe:natOp', ns).text if ide.find('nfe:natOp', ns) is not None else ""
            nat_op_min = str(nat_op).lower()

            numero_nota = ide.find('nfe:nNF', ns).text if ide.find('nfe:nNF', ns) is not None else ""
            num_nota_chave = int(numero_nota) if numero_nota.isdigit() else numero_nota

            esta_cancelada = str(num_nota_chave) in notas_canceladas_no_lote or "cancel" in nat_op_min
            eh_devolucao_fisica = tp_nf == "0" or any(termo in nat_op_min for termo in ["devolu", "estorno", "retorno"])

            # Busca dinâmica da Nota Fiscal referenciada (se houver a tag NFref)

            nota_referencia_xml = ""
            nf_ref_element = root.find('.//nfe:NFref/nfe:refNFe', ns)
            
            # Só tenta ler o texto se o elemento realmente existir no XML
            if nf_ref_element is not None and nf_ref_element.text and len(nf_ref_element.text) == 44:
                num_ref_extraido = nf_ref_element.text[25:34]
                nota_referencia_xml = int(num_ref_extraido) if num_ref_extraido.isdigit() else num_ref_extraido

            # 2. Definição dos Status (Sem repetição de código e sem risco de quebrar)
            if esta_cancelada:
                status_final = "Cancelada"
                tipo_operacao = "Cancelamento"
                motivo_final = notas_canceladas_no_lote.get(str(num_nota_chave), "Cancelado no faturamento")
                nota_referencia_xml = "CANEVE" 

                
            elif eh_devolucao_fisica:
                status_final = "Devolvido"
                tipo_operacao = "Devolução"
                motivo_final = "Devolvida / Retorno"
                
            else:
                status_final = "Vendido"
                tipo_operacao = "Venda"
                motivo_final = "Faturamento normal"
                nota_referencia_xml = "" 
                

            #PLATAFORMA
            plataforma_final = "Não identificada"
            
            # Extrai o texto da tag de informações complementares do XML atual
            inf_cpl_atual = root.find('.//nfe:infAdic/nfe:infCpl', ns)
            texto_cpl_atual = str(inf_cpl_atual.text).lower() if inf_cpl_atual is not None and inf_cpl_atual.text else ""
            
            num_nota_string = str(num_nota_chave).strip()
            excecoes_site = ("34469", "34519", "34182", "34880", "35006", "34980", "34981")

            if num_nota_string in excecoes_site:
                plataforma_final = "Site"

            elif not eh_devolucao_fisica and not esta_cancelada:
                if "mercado livre" in texto_cpl_atual or "mercad livre" in texto_cpl_atual :
                    plataforma_final = "Mercado Livre"
                    if "full" in texto_cpl_atual or "Venda full" in texto_cpl_atual: 
                        plataforma_final = "Mercado Livre Full"
                elif "shopee" in texto_cpl_atual or "Venda shopee" in texto_cpl_atual:
                    plataforma_final = "Shopee"
                elif "magalu" in texto_cpl_atual or "magazine" in texto_cpl_atual or "Venda magalu" in texto_cpl_atual:
                    plataforma_final = "Magalu"
                elif "site" in texto_cpl_atual or 'Venda Site' in texto_cpl_atual:
                    plataforma_final = "Site"
                else:
                    continue

            # CENÁRIO 2: É UMA NOTA DE DEVOLUÇÃO OU CANCELADA
            else:
                # Pegamos o texto original da tag (sem converter para minúsculo) para fazer o fatiamento correto
                texto_original = inf_cpl_atual.text if inf_cpl_atual is not None and inf_cpl_atual.text else ""
                num_nota_busca = texto_original[42:].strip() if texto_original else ""
                
                # Só entra na pasta se o fatiamento realmente extraiu um número
                if num_nota_busca in excecoes_site:
                    plataforma_final = "Site"
                elif num_nota_busca:
                    pasta_obj = Path(pasta_sistema_externo)
                    
                    # Varre a pasta procurando o XML que contém esse número no nome
                    for nome_arq in pasta_obj.iterdir():
                        if nome_arq.is_file() and nome_arq.suffix.lower() == ".xml":
                            if num_nota_busca in nome_arq.name:
                                try:
                                    # Abre o XML de venda localizado na pasta
                                    tree_externo = ET.parse(nome_arq)
                                    root_externo = tree_externo.getroot()
                                    
                                    # Procura a plataforma dentro do arquivo de venda que encontramos
                                    inf_cpl_ext = root_externo.find('.//nfe:infAdic/nfe:infCpl', ns)
                                    texto_cpl_venda = str(inf_cpl_ext.text).lower() if inf_cpl_ext is not None and inf_cpl_ext.text else ""
                                    
                                    if "mercado livre" in texto_cpl_venda:
                                        plataforma_final = "Mercado Livre"
                                        if "full" in texto_cpl_venda: plataforma_final = "Mercado Livre Full"
                                    elif "shopee" in texto_cpl_venda:
                                        plataforma_final = "Shopee"
                                    elif "magalu" in texto_cpl_venda or "magazine" in texto_cpl_venda:
                                        plataforma_final = "Magalu"
                                    elif "site" in texto_cpl_venda:
                                        plataforma_final = "SITE"
                                    
                                    break # Encontrou o arquivo de venda? Para o laço for imediatamente.
                                except:
                                    pass

            # Tratamento de Datas
            dh_emi_element = ide.find('nfe:dhEmi', ns)
            dh_emi_text = dh_emi_element.text[:10] if dh_emi_element is not None and dh_emi_element.text else ""
            data_final, mes_extenso, semana_texto = "", "Não identificado", ""

            if dh_emi_text:
                try:
                    data_obj = datetime.strptime(dh_emi_text, "%Y-%m-%d")
                    data_final = data_obj.strftime("%d/%m/%Y")
                    meses_map = {1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril", 5: "Maio", 6: "Junho", 7: "Julho", 8: "Agosto", 9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro"}
                    mes_extenso = meses_map[data_obj.month]
                    dia_do_ano = data_obj.timetuple().tm_yday
                    dia_da_semana_ano_novo = datetime(data_obj.year, 1, 1).weekday()
                    num_semana_ano = (dia_do_ano + dia_da_semana_ano_novo - 1) // 7 + 1
                    semana_texto = f"Semana {52 if num_semana_ano == 0 else num_semana_ano}"
                except: 
                    data_final = dh_emi_text

            dest = root.find('.//nfe:dest', ns)
            cliente = dest.find('nfe:xNome', ns).text if dest is not None and dest.find('nfe:xNome', ns) is not None else ""
            doc_cliente = ""
            if dest is not None:
                cnpj_el = dest.find('nfe:CNPJ', ns)
                cpf_el = dest.find('nfe:CPF', ns)
                doc_cliente = cnpj_el.text if cnpj_el is not None else (cpf_el.text if cpf_el is not None else "")

            ender_dest = root.find('.//nfe:dest/nfe:enderDest', ns)
            estado = ender_dest.find('nfe:UF', ns).text if ender_dest is not None and ender_dest.find('nfe:UF', ns) is not None else ""

            # ==============================================================================
            # PROCESSAMENTO DE ITENS (CORRIGIDO E UNIFICADO)
            # ==============================================================================
            detalhes_produtos = root.findall('.//nfe:det', ns)

            for prod_det in detalhes_produtos:
                prod = prod_det.find('nfe:prod', ns)
                if prod is None: 
                    continue

                # Extrai os dados individuais e corretos do produto atual da linha
                descricao_xml = prod.find('nfe:xProd', ns).text if prod.find('nfe:xProd', ns) is not None else ""
                sku_xml = prod.find('nfe:cProd', ns).text if prod.find('nfe:cProd', ns) is not None else ""
                qtd = float(prod.find('nfe:qCom', ns).text) if prod.find('nfe:qCom', ns) is not None else 0.0

                # --- CÁLCULO DO VALOR LÍQUIDO COMERCIAL ---
                vprod = prod.find('nfe:vProd', ns)
                valor_bruto_item = float(vprod.text) if vprod is not None and vprod.text else 0.0

                v_frete = float(prod.find('nfe:vFrete', ns).text) if prod.find('nfe:vFrete', ns) is not None and prod.find('nfe:vFrete', ns).text else 0.0
                v_desc  = float(prod.find('nfe:vDesc', ns).text)  if prod.find('nfe:vDesc', ns) is not None and prod.find('nfe:vDesc', ns).text else 0.0
                v_outro = float(prod.find('nfe:vOutro', ns).text) if prod.find('nfe:vOutro', ns) is not None and prod.find('nfe:vOutro', ns).text else 0.0

                # --- EXTRAÇÃO DE IMPOSTOS ADICIONAIS (IPI E ICMS ST) ---
                v_ipi = 0.0
                v_icms_st = 0.0
                imposto = prod_det.find('nfe:"imposto"', ns)
                
                if imposto is not None:
                    # 1. Busca o valor do IPI se houver na nota
                    ipi_tag = imposto.find('.//nfe:IPI/nfe:IPITrib/nfe:vIPI', ns)
                    if ipi_tag is not None and ipi_tag.text:
                        v_ipi = float(ipi_tag.text)
                    
                    # 2. Busca o valor do ICMS ST (pode estar em várias sub-tags de ICMS, ex: ICMS10, ICMS30, ICMS70, ICMS90)
                    icms_st_tag = imposto.find('.//nfe:ICMS//nfe:vICMSST', ns)
                    if icms_st_tag is not None and icms_st_tag.text:
                        v_icms_st = float(icms_st_tag.text)

                # --- CÁLCULO FINAL EXATO (IGUAL AO TOTAL DA NF DO SISTEMA) ---
                # Soma tudo o que agrega valor ao item e subtrai os descontos
                valor_total_item = valor_bruto_item + v_frete + v_outro + v_ipi + v_icms_st - v_desc

                # Arredonda o valor final para o Excel
                valor_total_item = round(valor_total_item, 2)

                # Organiza a separação das colunas de acordo com o destino
                if status_final in ["Cancelada", "Devolvido"]:
                    linha_devolucao = {
                        "PLATAFORMA": plataforma_final, "MÊS": mes_extenso, "DATA": data_final, "SEMANA": semana_texto,
                        "CLIENTE": cliente, "CPF/CNPJ": doc_cliente, "NOTA": num_nota_chave, 
                        "NOTA REFERÊNCIA": nota_referencia_xml,  # Coluna exclusiva aqui
                        "ESTADO": estado, "QTD": qtd, "SKU": sku_xml, "DESCRIÇÃO": descricao_xml, "VALOR": valor_total_item,
                        "STATUS": status_final, "MOTIVO": motivo_final
                    }
                    dados_devolucoes.append(linha_devolucao)
                else:
                    linha_venda = {
                        "PLATAFORMA": plataforma_final, "MÊS": mes_extenso, "DATA": data_final, "SEMANA": semana_texto,
                        "CLIENTE": cliente, "CPF/CNPJ": doc_cliente, "NOTA": num_nota_chave, # Sem nota referência aqui
                        "ESTADO": estado, "QTD": qtd, "SKU": sku_xml, "DESCRIÇÃO": descricao_xml, "VALOR": valor_total_item,
                        "STATUS": status_final, "MOTIVO": motivo_final
                    }
                    dados_venda.append(linha_venda)

        except Exception as e:
            print(f"Erro ao processar arquivo {arquivo}: {e}")

# ==============================================================================
# PASSAGEM 3: GERAÇÃO DA PLANILHA EXCEL FINAL
# ==============================================================================
print("\nPasso 3: Gravando dados no arquivo Excel...")

df_vendas = pd.DataFrame(dados_venda, columns=colunas_pedidos)
df_devolucoes = pd.DataFrame(dados_devolucoes, columns=colunas_cancelados_devolucoes)

with pd.ExcelWriter(caminho_planilha_teste, engine='openpyxl') as writer:
    df_vendas.to_excel(writer, sheet_name='Pedidos', index=False)
    df_devolucoes.to_excel(writer, sheet_name='Cancelados e Devolucoes', index=False)
print("NOTAS ADICIONADAS MANUAL, VIA CÓDIGO!")
print(excecoes_site)


print(f"=== PROCESSO CONCLUÍDO COM SUCESSO! ===")
