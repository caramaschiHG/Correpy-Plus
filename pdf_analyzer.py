"""
Módulo para análise detalhada de PDFs de notas de corretagem
"""
import io
import os
import pdfplumber
import pandas as pd
import re
from datetime import datetime

class NotaCorretagemAnalyzer:
    """Classe para análise avançada de PDFs de notas de corretagem"""
    
    def __init__(self, caminho_pdf):
        """Inicializa o analisador com o caminho para um arquivo PDF"""
        self.caminho_pdf = caminho_pdf
        self.nome_arquivo = os.path.basename(caminho_pdf)
        self.texto_completo = ""
        self.tabelas = []
        self.data_nota = None
        self.numero_nota = None
        self.corretora = None
        self.transacoes = []
        self.taxas = {}
        self.resultado = {}
        
    def extrair_conteudo(self):
        """Extrai o conteúdo completo do PDF"""
        try:
            with pdfplumber.open(self.caminho_pdf) as pdf:
                self.texto_completo = ""
                for pagina in pdf.pages:
                    self.texto_completo += pagina.extract_text() + "\n"
                    tabelas = pagina.extract_tables()
                    if tabelas:
                        self.tabelas.extend(tabelas)
                return True
        except Exception as e:
            print(f"Erro ao extrair conteúdo: {str(e)}")
            return False
            
    def analisar(self):
        """Analisa o conteúdo do PDF e extrai dados importantes"""
        if not self.extrair_conteudo():
            return False
            
        self._extrair_informacoes_basicas()
        self._extrair_transacoes()
        self._extrair_taxas()
        
        return {
            "sucesso": True,
            "nome_arquivo": self.nome_arquivo,
            "numero_nota": self.numero_nota,
            "data_nota": self.data_nota,
            "corretora": self.corretora,
            "transacoes": self.transacoes,
            "taxas": self.taxas,
            "resumo": self.resultado
        }
        
    def _extrair_informacoes_basicas(self):
        """Extrai informações básicas como número da nota, data e corretora"""
        # Extrair número da nota
        padrao_nr_nota = r"Nr\. nota:\s*(\d+)"
        match = re.search(padrao_nr_nota, self.texto_completo)
        if match:
            self.numero_nota = match.group(1)
            
        # Extrair data
        padrao_data = r"(Data pregão|Data):\s*(\d{2}/\d{2}/\d{4})"
        match = re.search(padrao_data, self.texto_completo)
        if match:
            try:
                data_str = match.group(2)
                self.data_nota = datetime.strptime(data_str, "%d/%m/%Y").date()
            except:
                pass
                
        # Extrair corretora
        # Procura pelo nome da corretora no cabeçalho
        primeiras_linhas = self.texto_completo.split('\n')[:5]
        for linha in primeiras_linhas:
            if "CCTVM" in linha or "Corretora" in linha or "CORRETORA" in linha:
                self.corretora = linha.strip()
                break
                
    def _extrair_transacoes(self):
        """Extrai as transações da nota de corretagem"""
        # Procura por tabelas que contêm dados de transações
        transacoes_encontradas = []
        
        # Tenta encontrar a tabela de negócios realizados
        for tabela in self.tabelas:
            # Verifica se é uma tabela de transações (procura por cabeçalhos típicos)
            cabecalhos = [str(col).lower() if col else "" for col in tabela[0]] if tabela else []
            cabecalho_str = " ".join(cabecalhos)
            
            if any(termo in cabecalho_str for termo in ["c/v", "tipo", "prazo", "quantidade", "preço", "valor"]):
                # Provável tabela de transações
                for linha in tabela[1:]:  # Pula o cabeçalho
                    if not linha or all(not cell for cell in linha):
                        continue
                        
                    # Processar a linha como uma transação
                    transacao = self._processar_linha_transacao(linha, cabecalhos)
                    if transacao:
                        transacoes_encontradas.append(transacao)
        
        # Se não encontrou transações nas tabelas, tenta extrair do texto
        if not transacoes_encontradas:
            # Procurar por padrões de texto que indicam transações
            padrao_transacao = r"([CV])\s+(\w+)\s+(\w+)\s+(\d+)\s+([\d,.]+)\s+([\d,.]+)"
            matches = re.finditer(padrao_transacao, self.texto_completo)
            
            for match in matches:
                try:
                    tipo_op = "C" if match.group(1) == "C" else "V"
                    ativo = match.group(2)
                    qtd = int(match.group(4).replace(".", ""))
                    preco = float(match.group(5).replace(".", "").replace(",", "."))
                    valor = float(match.group(6).replace(".", "").replace(",", "."))
                    
                    transacao = {
                        "tipo": tipo_op,
                        "ativo": ativo,
                        "quantidade": qtd,
                        "preco_unitario": preco,
                        "valor_total": valor
                    }
                    transacoes_encontradas.append(transacao)
                except:
                    continue
                    
        self.transacoes = transacoes_encontradas
        
    def _processar_linha_transacao(self, linha, cabecalhos):
        """Processa uma linha da tabela como uma transação"""
        try:
            transacao = {}
            
            # Mapeamento de índices comuns
            indices = {
                "tipo": next((i for i, h in enumerate(cabecalhos) if "c/v" in h or "tipo" in h), -1),
                "ativo": next((i for i, h in enumerate(cabecalhos) if "titulo" in h or "ativo" in h or "especificação" in h), -1),
                "quantidade": next((i for i, h in enumerate(cabecalhos) if "qtde" in h or "quantidade" in h), -1),
                "preco": next((i for i, h in enumerate(cabecalhos) if "preço" in h or "valor" in h), -1),
                "valor_total": next((i for i, h in enumerate(cabecalhos) if "total" in h or "valor op" in h), -1)
            }
            
            # Extrair dados da linha
            for campo, indice in indices.items():
                if indice >= 0 and indice < len(linha):
                    valor = linha[indice]
                    if campo == "tipo":
                        transacao[campo] = "C" if str(valor).upper() in ["C", "COMPRA"] else "V"
                    elif campo in ["quantidade", "preco", "valor_total"]:
                        try:
                            # Limpar e converter valores numéricos
                            if isinstance(valor, str):
                                valor = valor.replace(".", "").replace(",", ".")
                            transacao[campo] = float(valor)
                        except:
                            transacao[campo] = 0
                    else:
                        transacao[campo] = str(valor).strip()
                        
            return transacao if len(transacao) >= 3 else None  # Retorna apenas se tiver pelo menos 3 campos
        except:
            return None
            
    def _extrair_taxas(self):
        """Extrai informações sobre taxas e custos da nota"""
        # Padrões comuns para identificar taxas
        padroes_taxas = {
            "taxa_liquidacao": r"Taxa de [Ll]iquida[çc][ãa]o:?\s*([\d,.]+)",
            "taxa_registro": r"Taxa de [Rr]egistro:?\s*([\d,.]+)",
            "taxa_termo": r"Taxa de [Tt]ermo/[Oo]p[çc][õo]es:?\s*([\d,.]+)",
            "taxa_ana": r"Taxa A\.N\.A:?\s*([\d,.]+)",
            "emolumentos": r"[Ee]molumentos:?\s*([\d,.]+)",
            "taxa_operacional": r"Taxa [Oo]peracional:?\s*([\d,.]+)",
            "execucao": r"[Ee]xecu[çc][ãa]o:?\s*([\d,.]+)",
            "corretagem": r"[Cc]orretagem:?\s*([\d,.]+)",
            "iss": r"ISS:?\s*([\d,.]+)",
            "irrf": r"I\.R\.R\.F:?\s*([\d,.]+)",
            "outras_taxas": r"[Oo]utras [Tt]axas:?\s*([\d,.]+)",
            "valor_liquido": r"[Vv]alor [Ll][íi]quido:?\s*([\d,.]+)"
        }
        
        for nome_taxa, padrao in padroes_taxas.items():
            match = re.search(padrao, self.texto_completo)
            if match:
                try:
                    valor_str = match.group(1).replace(".", "").replace(",", ".")
                    self.taxas[nome_taxa] = float(valor_str)
                except:
                    self.taxas[nome_taxa] = 0.0
            else:
                self.taxas[nome_taxa] = 0.0
                
        # Extrair informações de resultado (valor de compras, vendas, etc.)
        padrao_compras = r"[Vv]alor de [Cc]ompras?:?\s*([\d,.]+)"
        match = re.search(padrao_compras, self.texto_completo)
        if match:
            try:
                self.resultado["valor_compras"] = float(match.group(1).replace(".", "").replace(",", "."))
            except:
                self.resultado["valor_compras"] = 0.0
                
        padrao_vendas = r"[Vv]alor de [Vv]endas?:?\s*([\d,.]+)"
        match = re.search(padrao_vendas, self.texto_completo)
        if match:
            try:
                self.resultado["valor_vendas"] = float(match.group(1).replace(".", "").replace(",", "."))
            except:
                self.resultado["valor_vendas"] = 0.0

# Função auxiliar para analisar um PDF e retornar os dados extraídos
def analisar_pdf_nota_corretagem(caminho_pdf):
    """Analisa um PDF de nota de corretagem e retorna os dados extraídos"""
    analyzer = NotaCorretagemAnalyzer(caminho_pdf)
    return analyzer.analisar()
