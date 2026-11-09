import os
import xml.etree.ElementTree as ET
import pandas as pd
from datetime import datetime
import re
import time
from formatar_excel import formatar_planilha_vendas, formatar_planilha_devolucao


class processadorXML:
    def __init__(self, caminho_xml, pasta_planilha):
        # Configurações iniciais e namespaces
        self.ns = {'nfe': 'http://portalfiscal.inf.br'}
        self.caminho_xml = camino_xml  
        self.pasta_planilha = pasta_planilha

        # Sets para controle de duplicados
        self.notas_vendas_existentes = set()
        self.notas_devolucoes_existentes = set()

        # Armazenamento dos dados finais
        self.dados_venda = []
        self.dados_devolucoes = []

        # Definição dos cabeçalhos das colunas
        self.colunas_vendas = [
            "PLATAFORMA", "MES", "DATA EMISSAO", "SEMANA", "CLIENTE", "ID_CLIENTE", "NOTA", "ESTADO", 
            "QTD", "ID_PRODUTO", "METRAGEM", "MODELO", "METROS TOTAL", "VALOR"
        ]

        self.coluna_devolucoes = [
            "PLATAFORMA", "MES", "SEMANA", "ID_PRODUTO", "NOME_PRODUTO", "QTD", "ID_CLIENTE",
            "NF VENDA", "MOTIVO", "NF DEVOLUCAO", "ESTADO", "DATA DEVOLUÇÃO", "VALOR", "DIAS EM USO"
        ]
        
    def varredura_pastas(self):
        # 1. Indicação do local raiz dos XMLs
        pasta_raiz = r"y:\ADMINISTRATIVO\\Vanessa - Vendas do E-Commerce\\Vendas 2026"
        
        # 2. Mapeamento recursivo procurando arquivos XML
        for raiz, pastas, arquivos in os.walk(pasta_raiz):
            for arquivo in arquivos:
                if arquivo.endswith('.xml'):
                    
                    # Definição e normalização dos caminhos de pasta
                    self.caminho_completo_xml = os.path.join(raiz, arquivo)
                    self.path_normalize = os.path.normpath(raiz)
                    self.partes_do_path = self.path_normalize.split(os.sep)

                    # Triagem por estrutura de pastas (Pedidos vs Devoluções)
                    if len(self.partes_do_path) >= 3 and self.partes_do_path[-2] == "Pedidos":
                        self.mes = self.partes_do_path[-1]
                        self.tipo = "Pedidos"
                        self.plataforma = self.partes_do_path[-3]

                        if self.plataforma.lower() == "full":
                            self.plataforma = "Mercado Livre Full"

                    elif len(self.partes_do_path) >= 2 and any(p.lower().startswith("devolu") for p in self.partes_do_path):
                        self.tipo = "Devoluções"

                        if any("full" in p.lower() for p in self.partes_do_path):
                            self.plataforma = "Mercado Livre Full"
                        else:
                            self.past_atual = self.partes_do_path[-1].lower()
                            if "cancel" in self.past_atual or self.partes_do_path[-2].lower() == "full":
                                self.plataforma = self.partes_do_path[-3] if len(self.partes_do_path) >= 3 else self.partes_do_path[-2]
                            else:
                                self.plataforma = self.partes_do_path[-2]
                        
                        # Define o mês atual em português por extenso
                        self.mes = datetime.now().strftime("%B")
                    else:
                        continue

                    # Processamento interno do XML (Sem funções aninhadas com def)
                    try:
                        self.tree = ET.parse(self.caminho_completo_xml)
                        self.root = self.tree.getroot()

                        # Coleta dos dados gerais do cabeçalho
                        self.ide = self.root.find('.//nfe:ide', self.ns)
                        if self.ide is None:
                            continue

                        self.tp_nf = self.ide.find('nfe:tpNF', self.ns).text if self.ide.find('nfe:tpNF', self.ns) is not None else ""
                        self.num_nota = self.ide.find('nfe:nNF', self.ns).text if self.ide.find('nfe:nNF', self.ns) is not None else ""

                        # Validação de segurança contra registros duplicados
                        if self.tipo == "Pedidos":
                            if str(self.num_nota) in self.notas_vendas_existentes:
                                continue
                        elif self.tipo == "Devoluções":
                            if str(self.num_nota) in self.notas_devolucoes_existentes:
                                continue
                                
                        # Tratamento da data de emissão e cálculo da semana
                        self.dh_emi_element = self.ide.find('nfe:dhEmi', self.ns)
                        self.dh_emi_text = self.dh_emi_element.text[:10] if self.dh_emi_element is not None and self.dh_emi_element.text else ""
                        self.data_final = ""
                        self.semana_texto = ""

                        if self.dh_emi_text:
                            try:
                                self.data_obj = datetime.strptime(self.dh_emi_text, "%Y-%m-%d")
                                self.data_final = self.data_obj.strftime("%d/%m/%Y")

                                self.dia_do_ano = self.data_obj.timetuple().tm_yday
                                self.dia_da_semana_ano_novo = datetime(self.data_obj.year, 1, 1).weekday()
                                self.num_semana_ano = (self.dia_do_ano + self.dia_da_semana_ano_novo - 1) // 7 + 1
                                
                                if self.num_semana_ano == 0:
                                    self.num_semana_ano = 52
                                    
                                self.semana_texto = f"Semana {self.num_semana_ano}"
                            except Exception:
                                self.data_final = self.dh_emi_text
                                self.semana_texto = ""

                        # Leitura dos dados do cliente (<dest>)
                        self.dest = self.root.find('.//nfe:dest', self.ns)
                        self.cliente_nome = ""
                        self.id_cliente = ""
                        self.estado = ""
                        
                        if self.dest is not None:
                            # Nome do cliente
                            self.xnome_el = self.dest.find('nfe:xNome', self.ns)
                            self.cliente_nome = self.xnome_el.text if self.xnome_el is not None else ""
                            
                            # Lógica para CPF / CNPJ
                            self.cpf_el = self.dest.find('nfe:CPF', self.ns)
                            self.cnpj_el = self.dest.find('nfe:CNPJ', self.ns)
                            if self.cpf_el is not None and self.cpf_el.text:
                                self.id_cliente = "CPF"
                            elif self.cnpj_el is not None and self.cnpj_el.text:
                                self.id_cliente = self.cnpj_el.text
                            
                            # Estado (UF) do endereço
                            self.ender_dest = self.dest.find('nfe:enderDest', self.ns)
                            if self.ender_dest is not None:
                                self.uf_el = self.ender_dest.find('nfe:UF', self.ns)
                                self.estado = self.uf_el.text if self.uf_el is not None else ""

                        # Informações complementares da nota (<infCpl>)
                        self.inf_cpl_texto = ""
                        self.inf_cpl_el = self.root.find('.//nfe:infAdic/nfe:infCpl', self.ns)
                        if self.inf_cpl_el is not None and self.inf_cpl_el.text:
                            self.inf_cpl_texto = self.inf_cpl_el.text

                        # Captura dos elementos para cálculo do valor
                        self.detalhes_produtos = self.root.findall('.//nfe:det', self.ns)
                        self.total_itens = len(self.detalhes_produtos)

                        self.v_prod_nota_el = self.root.find('.//nfe:total/nfe:ICMSTot/nfe:vProd', self.ns)
                        self.v_frete_nota_el = self.root.find('.//nfe:total/nfe:ICMSTot/nfe:vFrete', self.ns)
                        self.v_desc_nota_el = self.root.find('.//nfe:total/nfe:ICMSTot/nfe:vDesc', self.ns)
                        self.v_pag_el = self.root.find('.//nfe:pag/nfe:detPag/nfe:vPag', self.ns)

                        self.total_v_prod = float(self.v_prod_nota_el.text) if self.v_prod_nota_el is not None else 0.0
                        self.total_frete = float(self.v_frete_nota_el.text) if self.v_frete_nota_el is not None else 0.0
                        self.total_desconto = float(self.v_desc_nota_el.text) if self.v_desc_nota_el is not None else 0.0
                        self.valor_pag_total = float(self.v_pag_el.text) if self.v_pag_el is not None else 0.0

                        # loop para percorrer os itens internos da nota
                        for prod_det in self.detalhes_produtos:
                            prod = prod_det.find('nfe:prod', self.ns)
                            if prod is None:
                                continue

                            # Dados básicos do item
                            self.id_produto = prod.find('nfe:cProd', self.ns).text if prod.find('nfe:cProd', self.ns) is not None else ""
                            self.nome_produto = prod.find('nfe:xProd', self.ns).text if prod.find('nfe:xProd', self.ns) is not None else ""
                            self.qtd = float(prod.find('nfe:qCom', self.ns).text) if prod.find('nfe:qCom', self.ns) is not None else 0.0

                            if self.tipo == "Pedidos":
                                self.modelo = "MANG" if str(self.id_produto).upper().startswith("MANG") else "OUTRO"
                                
                                # Extração da metragem da descrição usando regex
                                self.metragem = 0.0
                                if self.modelo == "MANG":
                                    match = re.search(r'(\d+(?:[.,]\d+)?)\s*(?:metros|mt)', str(self.nome_produto).lower())
                                    if match:
                                        self.metragem = float(match.group(1).replace(',', '.'))
                                
                                # Cálculo dos metros totais
                                self.metros_total = (self.qtd * self.metragem) if self.modelo == "MANG" else 0.0
                                
                                # Cálculo de valor (Item único vs Múltiplos itens com rateio)
                                if self.total_itens == 1 and self.valor_pag_total > 0:
                                    self.valor_item = self.valor_pag_total
                                else:
                                    v_prod_item = float(prod.find('nfe:vProd', self.ns).text) if prod.find('nfe:vProd', self.ns) is not None else 0.0
                                    v_ipi_item_el = prod_det.find('.//nfe:IPI//nfe:vIPI', self.ns)
                                    v_ipi_item = float(v_ipi_item_el.text) if v_ipi_item_el is not None else 0.0
                                    
                                    peso_item = (v_prod_item / self.total_v_prod) if self.total_v_prod > 0 else 0.0
                                    frete_proporcional = self.total_frete * peso_item
                                    desconto_proporcional = self.total_desconto * peso_item
                                    
                                    self.valor_item = (v_prod_item + v_ipi_item) + frete_proporcional - desconto_proporcional
                                
                                # Estrutura a linha e adiciona à lista de vendas
                                linha_venda = {
                                    "PLATAFORMA": self.plataforma, 
                                    "MES": self.mes, 
                                    "DATA EMISSAO": self.data_final,
                                    "SEMANA": self.semana_texto, 
                                    "CLIENTE": self.cliente_nome, 
                                    "ID_CLIENTE": self.id_cliente,
                                    "NOTA": self.num_nota, 
                                    "ESTADO": self.estado, 
                                    "QTD": self.qtd, 
                                    "ID_PRODUTO": self.id_produto,
                                    "METRAGEM": self.metragem, 
                                    "MODELO": self.modelo, 
                                    "METROS TOTAL": self.metros_total,
                                    "VALOR": self.valor_item
                                }
                                self.dados_venda.append(linha_venda)

                            elif self.tipo == "Devoluções":
                                # Extração por fatiamento de caracteres na tag complementar
                                self.nf_venda = ""
                                if len(self.inf_cpl_texto) >= 34:
                                    self.nf_venda = self.inf_cpl_texto[30:34].strip()
                                
                                # Cálculo de dias em uso
                                self.dias_em_uso = ""
                                if len(self.inf_cpl_texto) >= 86 and self.dh_emi_text:
                                    data_venda_texto = self.inf_cpl_texto[77:86].strip()
                                    try:
                                        data_venda_limpa = data_venda_texto.replace('/', '-').replace('.', '-')
                                        data_venda_obj = datetime.strptime(data_venda_limpa, "%d-%m-%Y")
                                        data_dev_obj = datetime.strptime(self.dh_emi_text, "%Y-%m-%d")
                                        self.dias_em_uso = (data_dev_obj - data_venda_obj).days
                                    except Exception:
                                        self.dias_em_uso = "Erro no cálculo"
                                
                                # Determinação do valor da devolução
                                if self.total_itens == 1 and self.valor_pag_total > 0:
                                    self.valor_devolucao = self.valor_pag_total
                                else:
                                    self.valor_devolucao = float(prod.find('nfe:vProd', self.ns).text) if prod.find('nfe:vProd', self.ns) is not None else 0.0
                                
                                # Estrutura a linha e adiciona à lista de devoluções
                                linha_devolucao = {
                                    "PLATAFORMA": self.plataforma, 
                                    "MES": self.mes, 
                                    "SEMANA": self.semana_texto,
                                    "ID_PRODUTO": self.id_produto, 
                                    "NOME_PRODUTO": self.nome_produto, 
                                    "QTD": self.qtd,
                                    "ID_CLIENTE": self.id_cliente, 
                                    "NF VENDA": self.nf_venda, 
                                    "MOTIVO": "",
                                    "NF DEVOLUCAO": self.num_nota, 
                                    "ESTADO": self.estado, 
                                    "DATA DEVOLUCAO": self.data_final,
                                    "VALOR": self.valor_devolucao, 
                                    "DIAS EM USO": self.dias_em_uso
                                }
                                self.dados_devolucoes.append(linha_devolucao)

                    except Exception as e:
                        print(f"Erro ao ler o arquivo {arquivo}: {e}")

                    def exportar_excel(self):
                        # Transforma os dados em dataframes usando as colunas corretas
                        df_vendas = pd.DataFrame(self.dados_venda, columns=self.colunas_vendas)
                        df_devolucoes = pd.DataFrame(self.dados_devolucoes, columns=self.coluna_devolucoes)
                        
                        # Salva em arquivo Excel
                        caminho_arquivo = os.path.join(self.pasta_planilha, "Relatorio_Faturamento.xlsx")
                        with pd.ExcelWriter(caminho_arquivo, engine='openpyxl') as writer:
                            df_vendas.to_excel(writer, sheet_name='Pedidos', index=False)
                            df_devolucoes.to_excel(writer, sheet_name='Devolucoes', index=False)
                            
                        print(f"Processamento finalizado. Arquivo gerado em: {caminho_arquivo}")
