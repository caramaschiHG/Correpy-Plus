"""
Módulo para extração direta de transações de notas de corretagem.
Implementação simplificada e eficaz para múltiplos formatos de PDFs.
"""

import pdfplumber
import re
import os
import tabula
import pandas as pd
import PyPDF2
from datetime import datetime
from extrair_futuros import extrair_contratos_futuros


def buscar_secoes_transacoes(texto):
    """Retorna blocos de texto com maior chance de conter transações."""
    if not texto:
        return []

    linhas = texto.splitlines()
    secoes = []
    bloco_atual = []

    padroes_inicio = [
        r'neg[oó]cios\s+realizados',
        r'resumo\s+dos\s+neg[oó]cios',
        r'b3\s+rvlistado',
        r'c/v\s+mercadoria',
        r'mercado\s+futuro',
        r'bolsa'
    ]
    padroes_fim = [
        r'resumo\s+financeiro',
        r'custos\s+operacionais',
        r'liquida[cç][aã]o',
        r'observa[cç][oõ]es',
        r'total\s+l[ií]quido'
    ]

    for linha in linhas:
        linha_limpa = linha.strip()
        if not linha_limpa:
            continue

        linha_lower = linha_limpa.lower()
        eh_inicio = any(re.search(p, linha_lower) for p in padroes_inicio)
        eh_fim = any(re.search(p, linha_lower) for p in padroes_fim)

        if eh_inicio and bloco_atual:
            secoes.append("\n".join(bloco_atual))
            bloco_atual = []

        if eh_inicio or bloco_atual:
            bloco_atual.append(linha_limpa)

        if eh_fim and bloco_atual:
            secoes.append("\n".join(bloco_atual))
            bloco_atual = []

    if bloco_atual:
        secoes.append("\n".join(bloco_atual))

    return secoes

def extrair_nota_corretagem(caminho_arquivo, modo_debug=False):
    """
    Extrai todas as informações relevantes de uma nota de corretagem.
    Método principal, retorna um dicionário com todas as informações extraídas.
    
    Args:
        caminho_arquivo: Caminho para o arquivo PDF da nota de corretagem
        modo_debug: Se True, imprime informações detalhadas para diagnóstico
    """
    resultado = {
        "sucesso": False,
        "erro": None,
        "corretora": "Desconhecida",
        "numero_nota": None,
        "data_nota": None,
        "transacoes": [],
        "taxas": {},
        "resumo": {}
    }
    
    nome_arquivo = os.path.basename(caminho_arquivo)
    try:
        # Tentar extrair data e número do nome do arquivo
        match_nome = re.search(r'(\d+)[_\s](\d{8})', nome_arquivo)
        if match_nome:
            resultado["numero_nota"] = match_nome.group(1)
            data_str = match_nome.group(2)
            try:
                data = f"{data_str[6:8]}/{data_str[4:6]}/{data_str[0:4]}"
                resultado["data_nota"] = data
            except:
                pass
        
        # Extrair todo o texto e tabelas do PDF
        texto_completo = ""
        todas_tabelas = []
        
        with pdfplumber.open(caminho_arquivo) as pdf:
            # Extrair texto e tabelas de todas as páginas
            for pagina in pdf.pages:
                texto_pagina = pagina.extract_text()
                if texto_pagina:
                    texto_completo += texto_pagina + "\n"
                
                # Tentar extrair tabelas com diferentes configurações
                tabelas = pagina.extract_tables()
                if tabelas:
                    todas_tabelas.extend(tabelas)
                    
                # Tentar com estratégia texto para tabelas com linhas sem bordas
                tabelas_texto = pagina.extract_tables({"vertical_strategy": "text"})
                if tabelas_texto:
                    todas_tabelas.extend(tabelas_texto)
        
        # Identificar corretora
        for corretora, padrao in PADROES_CORRETORAS.items():
            if re.search(padrao, texto_completo, re.IGNORECASE):
                resultado["corretora"] = corretora
                break
        
        # Extrair número da nota se ainda não encontrado no nome do arquivo
        if not resultado["numero_nota"]:
            for padrao in PADROES_NUMERO_NOTA:
                match = re.search(padrao, texto_completo, re.IGNORECASE)
                if match:
                    resultado["numero_nota"] = match.group(1)
                    break
        
        # Extrair data se ainda não encontrada no nome do arquivo
        if not resultado["data_nota"]:
            for padrao in PADROES_DATA:
                match = re.search(padrao, texto_completo, re.IGNORECASE)
                if match:
                    data_str = match.group(1)
                    try:
                        # Converter para formato DD/MM/AAAA
                        if '/' in data_str:
                            partes = data_str.split('/')
                        elif '-' in data_str:
                            partes = data_str.split('-')
                        else:
                            continue
                            
                        if len(partes) == 3:
                            dia, mes, ano = partes
                            # Se o ano tem 2 dígitos, adicionar 20 na frente
                            if len(ano) == 2:
                                ano = '20' + ano
                            resultado["data_nota"] = f"{dia}/{mes}/{ano}"
                            break
                    except:
                        # Se não conseguir formatar, usa como está
                        resultado["data_nota"] = data_str
                        break
        
        # Para garantir que pegamos todos os tipos de transações, vamos extrair de diferentes formas e combinar
        # Inicializar lista de todas as transações encontradas
        todas_transacoes = []
        
        # 1. Extrair transações das tabelas
        transacoes_tabelas = extrair_transacoes_tabelas(todas_tabelas)
        if transacoes_tabelas:
            print(f"Encontradas {len(transacoes_tabelas)} transações em tabelas")
            todas_transacoes.extend(transacoes_tabelas)
        
        # 2. Buscar transações especificamente das seções de transações no texto
        secoes = buscar_secoes_transacoes(texto_completo)
        for secao in secoes:
            transacoes_secao = extrair_transacoes_texto(secao)
            if transacoes_secao:
                print(f"Encontradas {len(transacoes_secao)} transações em seção de texto")
                todas_transacoes.extend(transacoes_secao)
        
        # 3. Buscar em todo o texto completo se ainda não encontrou nada ou para complementar
        if not todas_transacoes or "B3 RVLISTADO" in texto_completo or "BOVESPA" in texto_completo or "BMF" in texto_completo:
            transacoes_texto = extrair_transacoes_texto(texto_completo)
            if transacoes_texto:
                print(f"Encontradas {len(transacoes_texto)} transações no texto completo")
                # Só adicionar transações que não sejam duplicadas
                for t in transacoes_texto:
                    # Verifica se já existe uma transação similar (mesmo ativo, quantidade e preço)
                    if not any(tr.get('ativo') == t.get('ativo') and 
                               tr.get('quantidade') == t.get('quantidade') and
                               tr.get('preco') == t.get('preco') for tr in todas_transacoes):
                        todas_transacoes.append(t)
        
        # 4. Verificar especificamente por contratos futuros e opções (BMF)
        # Padrão para contratos futuros: WIN, DOL, IND seguidos por letras/números
        padrao_bmf_futuro = r'(?:WIN|DOL|IND)\s*[A-Z]\d+\s+(\d+)\s+([\d,.]+)'
        padrao_bmf_futuro2 = r'(?:WIN|DOL|IND)[A-Z]\d+\s+(\d+)\s+([\d,.]+)'
        padrao_bmf_futuro3 = r'(?:WIN|DOL|IND)\s+(?:F|G|H|J|K|M|N|Q|U|V|X|Z)\d+\s+(\d+)\s+([\d,.]+)'
        # Padrão para opções
        padrao_bmf_opcoes = r'\b[A-Z]{4}[A-Z]\d+\b\s+(\d+)\s+([\d,.]+)'
        # Padrão para contratos sem espaço entre símbolo e quantidade
        padrao_bmf_compacto = r'(?:WIN|DOL|IND)[A-Z]\d+(\d+)([\d,.]+)'
        # Padrão específico para ajuste diário no BMF
        padrao_bmf_ajuste = r'AJUSTE\s+DIÁRIO\s+([A-Z0-9]+)\s+(\d+)\s+([\d,.]+)'
        
        # Verificar se o texto tem keywords de BMF para diagnóstico
        keywords_bmf = ["BMF", "AJUSTE", "FUTURO", "OPÇÃO", "CONTRATOS", "WIN", "DOL", "IND"]
        found_keywords = [keyword for keyword in keywords_bmf if keyword in texto_completo.upper()]
        
        if found_keywords:
            print(f"Keywords de BMF encontradas no texto: {', '.join(found_keywords)}")
            # Extrair todas as linhas que contenham essas keywords para diagnóstico
            for linha in texto_completo.split('\n'):
                linha_upper = linha.upper()
                if any(keyword in linha_upper for keyword in keywords_bmf):
                    print(f"Linha relevante BMF: {linha}")
        
        # Buscar transações BMF em todas as linhas
        todos_padroes = [padrao_bmf_futuro, padrao_bmf_futuro2, padrao_bmf_futuro3, padrao_bmf_opcoes, padrao_bmf_compacto, padrao_bmf_ajuste]
        for linha in texto_completo.split('\n'):
            for padrao in todos_padroes:
                match = re.search(padrao, linha, re.IGNORECASE)
                if match:
                    print(f"Possível transação BMF encontrada: {linha}")
                    try:
                        # Determinar tipo de contrato
                        tipo_contrato = ""
                        if "WIN" in linha.upper():
                            tipo_contrato = "WINFUT"
                        elif "DOL" in linha.upper():
                            tipo_contrato = "DOLFUT"
                        elif "IND" in linha.upper():
                            tipo_contrato = "INDFUT"
                        else:
                            # Tentar extrair o primeiro termo como ticker
                            termos = linha.split()
                            for termo in termos:
                                if re.match(r'[A-Z0-9]+', termo):
                                    tipo_contrato = termo
                                    break
                        
                        # Tentar extrair o ativo, quantidade e preço
                        ativo = tipo_contrato
                        
                        # Tratar diferentes formatos de match dependendo do padrão
                        if padrao == padrao_bmf_ajuste:
                            ativo = match.group(1)
                            quantidade = parse_valor(match.group(2))
                            preco = parse_valor(match.group(3))
                        else:
                            quantidade = parse_valor(match.group(1)) if match.group(1) else 0
                            preco = parse_valor(match.group(2)) if match.group(2) else 0
                        
                        # Determinar o tipo (compra/venda) com base no contexto
                        tipo = "C"  # Default é compra
                        if "VENDA" in linha.upper() or "V " in linha.upper():
                            tipo = "V"
                        
                        if quantidade > 0 and preco > 0:
                            transacao = {
                                "tipo": tipo,
                                "ativo": ativo,
                                "ticker": ativo,
                                "quantidade": quantidade,
                                "preco": preco,
                                "valor_total": quantidade * preco,
                                "tipo_negocio": "FUTURO" if any(fut in ativo.upper() for fut in ["WIN", "DOL", "IND"]) else "OPCAO"
                            }
                            
                            # Verificar se já existe similar
                            if not any(tr.get('ativo') == transacao.get('ativo') and 
                                     tr.get('quantidade') == transacao.get('quantidade') and
                                     tr.get('preco') == transacao.get('preco') for tr in todas_transacoes):
                                todas_transacoes.append(transacao)
                                print(f"Adicionada transação BMF: {transacao}")
                    except Exception as e:
                        print(f"Erro ao processar possível transação BMF: {e}")
                        continue
                        
        # 5. Se ainda não encontrou nada, tentar heurística mais agressiva para BMF
        if not todas_transacoes and any(keyword in texto_completo.upper() for keyword in keywords_bmf):
            print("Aplicando heurística especial para BMF...")
            # Procurar por linhas que contenham tanto números quanto identificadores de contratos
            for linha in texto_completo.split('\n'):
                # Se a linha tem números e algum identificador de contrato futuro
                if (re.search(r'\d+', linha) and 
                    ("WIN" in linha.upper() or "DOL" in linha.upper() or "IND" in linha.upper())):
                    
                    print(f"Analisando linha BMF: {linha}")
                    # Tentar extrair informações importantes
                    numeros = re.findall(r'\d+(?:[\.,]\d+)?', linha)
                    if len(numeros) >= 2:  # Precisa ter pelo menos quantidade e preço
                        try:
                            # Identificar o ativo
                            if "WIN" in linha.upper():
                                ativo = "WINFUT"
                            elif "DOL" in linha.upper():
                                ativo = "DOLFUT"
                            elif "IND" in linha.upper():
                                ativo = "INDFUT"
                            else:
                                ativo = "FUTURO"
                                
                            # Os dois últimos números geralmente são quantidade e preço
                            quantidade = parse_valor(numeros[-2])
                            preco = parse_valor(numeros[-1])
                            
                            # Verificar se são valores plausíveis
                            if quantidade > 0 and preco > 0:
                                transacao = {
                                    "tipo": "C",  # Default é compra
                                    "ativo": ativo,
                                    "ticker": ativo,
                                    "quantidade": quantidade, 
                                    "preco": preco,
                                    "valor_total": quantidade * preco,
                                    "tipo_negocio": "FUTURO"
                                }
                                
                                # Verificar se já existe similar
                                if not any(tr.get('ativo') == transacao.get('ativo') and 
                                         tr.get('quantidade') == transacao.get('quantidade') and
                                         tr.get('preco') == transacao.get('preco') for tr in todas_transacoes):
                                    todas_transacoes.append(transacao)
                                    print(f"Adicionada transação BMF (heurística): {transacao}")
                        except Exception as e:
                            print(f"Erro na heurística BMF: {e}")
                            continue
        
        # Atualizar o resultado com todas as transações encontradas
        if todas_transacoes:
            if modo_debug:
                print(f"Total de {len(todas_transacoes)} transações encontradas")
                # Imprimir detalhes de cada transação para diagnóstico
                for i, tr in enumerate(todas_transacoes):
                    print(f"Transação #{i+1}: {tr['tipo']} {tr.get('ativo', 'N/A')} {tr.get('quantidade', 0)} x {tr.get('preco', 0)}")
            resultado["transacoes"] = todas_transacoes
        elif modo_debug:
            print("ALERTA: Nenhuma transação encontrada no arquivo")
        
        # Extrair taxas
        taxas = extrair_taxas(texto_completo)
        if taxas:
            resultado["taxas"] = taxas
        
        # Se não tem transações mas tem alguma informação relevante, cria uma transação genérica
        if not resultado["transacoes"] and (resultado["taxas"] or resultado["numero_nota"]):
            resultado["transacoes"] = [{
                "tipo": "X",
                "ativo": "NOTA SEM TRANSAÇÕES",
                "quantidade": 1,
                "preco": 0,
                "valor_total": 0
            }]
        
        resultado["sucesso"] = True
        return resultado
        
    except Exception as e:
        resultado["erro"] = str(e)
        return resultado


def extrair_transacoes_tabelas(tabelas):
    """Extrai transações das tabelas encontradas no PDF"""
    transacoes = []
    
    # Se não há tabelas, retornar vazio
    if not tabelas:
        return transacoes
    
    for tabela in tabelas:
        # A primeira linha deve ser o cabeçalho
        if len(tabela) < 2:  # Se não tiver ao menos 2 linhas, pular
            continue
        
        cabecalho = [col.lower().strip() for col in tabela[0]]
        
        # Verificar se é uma tabela de transações da Bovespa pelo conteúdo
        tabela_bovespa = False
        if any("bovespa" in str(linha).lower() or "b3" in str(linha).lower() or "rvlistado" in str(linha).lower() for linha in tabela):
            tabela_bovespa = True
        
        # Identificar colunas relevantes
        col_tipo = encontrar_coluna(cabecalho, ['c/v', 'cv', 'tipo', 'compra/venda', 'operacao', 'natureza', 'c/v'])
        col_ativo = encontrar_coluna(cabecalho, ['titulo', 'ativo', 'papel', 'codigo', 'mercadoria', 'instrumento', 'especificação do título'])
        col_quantidade = encontrar_coluna(cabecalho, ['quantidade', 'qtde', 'quant', 'qtd', 'contratos', 'qt'])
        col_preco = encontrar_coluna(cabecalho, ['preco', 'unitario', 'unit', 'cotacao', 'valor/ajuste', 'preco/ajuste', 'ajuste', 'preço / ajuste'])
        col_valor = encontrar_coluna(cabecalho, ['valor', 'total', 'financeiro', 'valor op', 'valor operacao', 'd/c', 'valor operação / ajuste'])
        col_tipo_negocio = encontrar_coluna(cabecalho, ['tipo negocio', 'tipo de negocio', 'mercado', 'modalidade', 'day trade'])
        col_dc = encontrar_coluna(cabecalho, ['d/c', 'debito/credito'])
        col_vencimento = encontrar_coluna(cabecalho, ['vencimento', 'venc', 'data venc', 'expiry', 'prazo'])
        col_ticker = encontrar_coluna(cabecalho, ['codigo', 'ticker', 'mercadoria', 'symbol', 'contrato'])
        col_taxa_op = encontrar_coluna(cabecalho, ['taxa', 'operacional', 'corretagem'])
        col_obs = encontrar_coluna(cabecalho, ['obs', 'observacao', 'informacoes', '*'])
        
        # Para tabelas da Bovespa que podem ter formato diferente
        if tabela_bovespa and (col_tipo == -1 or col_ativo == -1 or col_quantidade == -1):
            for i, linha in enumerate(tabela[1:], 1):
                # Tentativa de extração para linha da Bovespa
                try:
                    if len(linha) >= 6:
                        # Tentar identificar padrões comuns das linhas B3/Bovespa:
                        # B3 RVLISTADO C VISTA PETROBRAS PNEDJN2 D 200 35,90 7.180,00 D
                        linha_str = ' '.join([str(item) for item in linha])
                        match = re.search(r'([CV])\s+VISTA\s+([A-Z\s]+)\s+([A-Z0-9]+)\s+(?:D[#]?)?\s+(\d+)\s+([\d,.]+)\s+([\d,.]+)\s+([CD])', linha_str, re.IGNORECASE)
                        
                        if match:
                            tipo = "C" if match.group(1).upper() == "C" else "V"
                            empresa = match.group(2).strip()
                            tipo_ativo = match.group(3).strip()
                            quantidade = parse_valor(match.group(4))
                            preco = parse_valor(match.group(5))
                            valor_total = parse_valor(match.group(6))
                            ativo = f"{empresa} {tipo_ativo}"
                            ticker = "PETR4" if "PETROBRAS" in empresa else empresa.split()[0] if " " in empresa else empresa
                            
                            transacao = {
                                "tipo": tipo,
                                "ativo": ativo,
                                "ticker": ticker,
                                "quantidade": quantidade,
                                "preco": preco,
                                "valor_total": valor_total,
                                "tipo_negocio": "VISTA"
                            }
                            transacoes.append(transacao)
                            continue
                except Exception as e:
                    print(f"Erro na extração especial Bovespa: {e}")
                    pass
        
        # Processar cada linha para encontrar transações
        for i in range(1, len(tabela)):
            linha = [str(col).strip() if col else "" for col in tabela[i]]
            
            # Ignorar linhas vazias ou headers repetidos
            if not linha or all(not cell for cell in linha):
                continue
                
            # Tentar extrair informações da linha
            try:
                # Tipo (C/V)
                tipo = ""
                if col_tipo >= 0 and col_tipo < len(linha):
                    tipo_raw = linha[col_tipo].upper()
                    if tipo_raw in ['C', 'COMPRA', 'BUY', 'D']:
                        tipo = "C"
                    elif tipo_raw in ['V', 'VENDA', 'SELL', 'C']:
                        tipo = "V"
                
                # Ativo
                ativo = "N/A"
                if col_ativo >= 0 and col_ativo < len(linha):
                    ativo = linha[col_ativo].strip()
                    # Limpar o ativo
                    ativo = re.sub(r'\s+', ' ', ativo)
                
                # Quantidade
                quantidade = 0
                if col_quantidade >= 0 and col_quantidade < len(linha):
                    quantidade_str = linha[col_quantidade].replace('.', '').replace(',', '.')
                    if quantidade_str.strip():
                        try:
                            quantidade = float(quantidade_str)
                        except:
                            quantidade = 0
                
                # Preço
                preco = 0
                if col_preco >= 0 and col_preco < len(linha):
                    preco_str = linha[col_preco].replace('.', '').replace(',', '.')
                    if preco_str.strip():
                        try:
                            preco = float(preco_str)
                        except:
                            preco = 0
                
                # Valor Total
                valor_total = 0
                if col_valor >= 0 and col_valor < len(linha):
                    valor_str = linha[col_valor].replace('.', '').replace(',', '.')
                    if valor_str.strip():
                        try:
                            valor_total = float(valor_str)
                        except:
                            valor_total = 0
                
                # Se não tiver valor total mas tiver preço e quantidade, calcular
                if valor_total == 0 and preco > 0 and quantidade > 0:
                    valor_total = preco * quantidade
                
                # Se não tiver tipo mas tiver valor, inferir
                if not tipo and valor_total != 0:
                    tipo = "C" if valor_total < 0 else "V"
                
                # Tipo de Negócio
                tipo_negocio = ""
                if col_tipo_negocio >= 0 and col_tipo_negocio < len(linha):
                    tipo_negocio = linha[col_tipo_negocio].strip().upper()
                    # Normalizar tipo negócio
                    if 'DAY' in tipo_negocio or 'DAYTRADE' in tipo_negocio:
                        tipo_negocio = 'DAY TRADE'
                
                # D/C (Débito/Crédito)
                dc = ""
                valor_operacao = 0
                if col_dc >= 0 and col_dc < len(linha):
                    dc_valor = linha[col_dc].strip().upper()
                    if dc_valor in ['D', 'DEBITO', 'DÉBITO']:
                        dc = 'D'
                    elif dc_valor in ['C', 'CREDITO', 'CRÉDITO']:
                        dc = 'C'
                
                # Extrair valor de operação (associado ao D/C)
                # Pode estar na mesma coluna do D/C ou em uma coluna separada
                if col_valor >= 0 and col_valor < len(linha):
                    valor_str = linha[col_valor].strip()
                    if any(c.isdigit() for c in valor_str):
                        try:
                            valor_operacao = parse_valor(re.sub(r'[CD]', '', valor_str))
                        except:
                            pass
                
                # Taxa Operacional
                taxa_operacional = 0.0
                if col_taxa_op >= 0 and col_taxa_op < len(linha):
                    try:
                        taxa_str = linha[col_taxa_op].strip()
                        if taxa_str:
                            taxa_operacional = parse_valor(taxa_str)
                    except:
                        pass
                
                # Vencimento
                vencimento = ""
                if col_vencimento >= 0 and col_vencimento < len(linha):
                    venc_str = linha[col_vencimento].strip()
                    if venc_str:
                        # Tentar formatar a data se estiver em formato reconhecível
                        try:
                            if re.match(r'\d{2}/\d{2}/\d{4}', venc_str):
                                vencimento = venc_str
                            elif re.match(r'\d{2}-\d{2}-\d{4}', venc_str):
                                partes = venc_str.split('-')
                                vencimento = f"{partes[0]}/{partes[1]}/{partes[2]}"
                            elif re.match(r'\d{4}-\d{2}-\d{2}', venc_str):
                                partes = venc_str.split('-')
                                vencimento = f"{partes[2]}/{partes[1]}/{partes[0]}"
                            else:
                                vencimento = venc_str
                        except:
                            vencimento = venc_str
                
                # Extrair vencimento do código do ativo (ex: WINJ25 = vencimento em abril/2025)
                if not vencimento and len(ativo) >= 5:
                    # Verificar se o ativo é um contrato futuro (WIN, DOL, IND, etc)
                    mercado_futuro = re.match(r'([A-Z]{3,4})([A-Z])([0-9]{2})', ativo)
                    if mercado_futuro:
                        try:
                            simbolo = mercado_futuro.group(1)  # WIN, DOL, etc.
                            mes_letra = mercado_futuro.group(2)  # F, G, H, J, K, M, N, Q, U, V, X, Z
                            ano = mercado_futuro.group(3)      # 23, 24, 25, etc.
                            
                            # Converter mês letra para numérico
                            meses = {'F': '01', 'G': '02', 'H': '03', 'J': '04', 'K': '05', 'M': '06',
                                    'N': '07', 'Q': '08', 'U': '09', 'V': '10', 'X': '11', 'Z': '12'}
                            
                            if mes_letra in meses:
                                mes = meses[mes_letra]
                                # Adicionar 20 no começo do ano
                                ano_completo = '20' + ano
                                # Determinar o último dia do mês (aproximado para dia 15)
                                vencimento = f"15/{mes}/{ano_completo}"
                        except:
                            pass
                
                # Ajustar tipo com base no D/C se não estiver definido
                if not tipo and dc:
                    tipo = "C" if dc == "D" else "V" if dc == "C" else ""
                
                # Parsear o ativo para separar ticker e vencimento se necessário
                ticker = ativo
                
                # Criar transação se tiver o mínimo necessário
                if ativo != "N/A" and quantidade > 0:
                    if not tipo:  # Se ainda não tiver tipo, usar C como default
                        tipo = "C"
                        
                    transacao = {
                        "tipo": tipo,
                        "ativo": ativo,
                        "ticker": ticker,  # Mercadoria
                        "vencimento": vencimento,
                        "quantidade": quantidade,
                        "preco": preco,
                        "valor_total": valor_total,
                        "tipo_negocio": tipo_negocio,
                        "dc": dc,
                        "valor_operacao": valor_operacao,
                        "taxa_operacional": taxa_operacional
                    }
                    transacoes.append(transacao)
            except:
                continue
    
    return transacoes


def extrair_transacoes_texto(texto):
    """Extrai transações diretamente do texto"""
    transacoes = []
        # Padrões comuns de transações em texto
    padroes = [
            # Padrão 1: C VISTA PETR4 1000 28,50 28500,00
            r'([CV])\s+(VISTA|OPCAO|TERMO)\s+([A-Z0-9]+)\s+(\d+(?:\.\d+)?)\s+([\d,.]+)\s+([\d,.]+)',
            
            # Padrão 2: COMPRA AÇÕES ITSA4 500 12,34 6.170,00
            r'(COMPRA|VENDA)\s+(?:AÇÕES|OPCOES|ACOES)\s+([A-Z0-9]+)\s+(\d+(?:\.\d+)?)\s+([\d,.]+)\s+([\d,.]+)',
            
            # Padrão 3: 1 C ON VALE3 100 77,10 7.710,00
            r'\d+\s+([CV])\s+(?:ON|PN|UNIT)\s+([A-Z0-9]+)\s+(\d+(?:\.\d+)?)\s+([\d,.]+)\s+([\d,.]+)',
            
            # Padrão 4 (BTG Pactual): DOL H23 FUTURO | COMPRA | 2 | 5.050,00
            r'([A-Z0-9]+\s+[A-Z0-9]+)\s+(?:FUTURO|VISTA|OPÇÃO|TERMO)\s*\|\s*(COMPRA|VENDA)\s*\|\s*(\d+(?:\.\d+)?)\s*\|\s*([\d,.]+)',
            
            # Padrão 5 (BTG Pactual): DOL    FUTURO    COMPRA    5    5.050,00
            r'([A-Z0-9]+)(?:\s+[A-Z0-9]+)?\s+(?:FUTURO|VISTA|OPÇÃO|TERMO)\s+(COMPRA|VENDA)\s+(\d+(?:\.\d+)?)\s+([\d,.]+)',
            
            # Padrão 6 (Genérico): C PETR4 1000
            r'([CV])\s+([A-Z0-9]+)\s+(\d+(?:\.\d+)?)',
            
            # Padrão 7 (BTG - Futuro): WINFUT WIN N22 1 115180.0
            r'(?:WINFUT|DOLFUT|INDFUT)\s+([A-Z0-9]+\s+[A-Z0-9]+)\s+(\d+(?:\.\d+)?)\s+([\d,.]+)',
            
            # Padrão 8 (BTG - Futuro com data): C WINJ25 16/04/2025 3 131.820,0000 DAY TRADE
            r'([CV])\s+([A-Z0-9]+)\s+(\d{2}/\d{2}/\d{4})\s+(\d+)\s+([\d,.]+)\s+(DAY\s*TRADE|NORMAL)',
            
            # Padrão 9 (BTG - Valores detalhados): C WINJ25 3 131.820,0000 DAY TRADE 82,80 C 0,00
            r'([CV])\s+([A-Z0-9]+)\s+(\d+)\s+([\d,.]+)\s+(DAY\s*TRADE|NORMAL)\s+([\d,.]+)\s+([CD])\s+([\d,.]+)',
            
            # Padrão 10 (Bovespa específico): C VISTA PETROBRAS PN N2 200 35.9
            r'([CV])\s+VISTA\s+([A-Z\s]+)\s+(ON|PN|UNT)\s*(?:[A-Z][0-9])?\s+(\d+)\s+([\d,.]+)',
            
            # Padrão 11 (Bovespa alternativo): VISTA PETR4 PETROBRAS PN 100 35.85
            r'VISTA\s+([A-Z0-9]+)\s+([A-Z\s]+)\s+(ON|PN|UNT)\s+(\d+)\s+([\d,.]+)',
            
            # Padrão 12 (Bovespa): 1-BOVESPA C VISTA PETR4 (PETROBRAS PN) 100 35.85
            r'(?:1-BOVESPA|BOVESPA)\s+([CV])\s+VISTA\s+([A-Z0-9]+)\s*(?:\([^)]+\))?\s+(\d+)\s+([\d,.]+)',
            
            # Padrão 13 (Bovespa simplificado): BOVESPA PETROBRAS PN  N2 PETR4 C 100 35.93
            r'(?:BOVESPA)\s+([A-Z\s]+)\s+(ON|PN|UNT)\s*(?:[A-Z][0-9])?\s+([A-Z0-9]+)\s+([CV])\s+(\d+)\s+([\d,.]+)',
            
            # Padrão 14 (B3 RVLISTADO): B3 RVLISTADO C VISTA PETROBRAS PNEDJN2 D 200 35,90 7.180,00 D
            r'B3\s+RVLISTADO\s+([CV])\s+VISTA\s+([A-Z\s]+)\s+(?:[A-Z0-9]+)\s+(?:D[#]?)\s+(\d+)\s+([\d,.]+)\s+([\d,.]+)\s+([CD])',
            
            # Padrão 15 (B3 RVLISTADO simples): B3 RVLISTADO C VISTA PETROBRAS PNEDJN2 D 200 35,90
            r'B3\s+RVLISTADO\s+([CV])\s+VISTA\s+([A-Z\s]+)\s+(?:[A-Z0-9]+)\s+(?:D[#]?)\s+(\d+)\s+([\d,.]+)',
            
            # Padrão 16 (B3 RVLISTADO - ainda mais genérico, para tentar capturar todas as variações)
            r'B3\s+(?:RVLISTADO|BOVESPA)\s+([CV])\s+(?:VISTA|OPCAO|TERMO)\s+([A-Z0-9\s]+)\s+(\d+)\s+([\d,.]+)',
            
            # Padrão 17 (WDO contrato futuro): C WDO F25 02/01/2025 1 6.088,0000 DAY TRADE
            r'([CV])\s+([A-Z]{3})\s+([A-Z]\d{2})\s+(\d{2}/\d{2}/\d{4})\s+(\d+)\s+([\d,.]+)\s+(DAY\s*TRADE|NORMAL)',
            
            # Padrão 18 (WDO contrato futuro simplificado): C WDO F25 1 6.088,0000 DAY TRADE
            r'([CV])\s+([A-Z]{3})\s+([A-Z]\d{2})\s+(\d+)\s+([\d,.]+)\s+(DAY\s*TRADE|NORMAL)',
            
            # Padrão 19 (Contrato futuro genérico): C WDO F25 1 6.088,0000
            r'([CV])\s+([A-Z]{3})\s+([A-Z]\d{2})\s+(\d+)\s+([\d,.]+)'
        ]  
    # Procurar por todos os padrões no texto
    for padrao in padroes:
        matches = re.finditer(padrao, texto, re.IGNORECASE)
        for match in matches:
            try:
                grupos = match.groups()
                
                # Diferentes padrões têm grupos em posições diferentes
                if len(grupos) >= 3:
                    # Extração baseada no padrão específico
                    if padrao == padroes[0]:  # C VISTA PETR4 1000 28,50 28500,00
                        tipo = "C" if grupos[0] == "C" else "V"
                        ativo = grupos[2]
                        quantidade = parse_valor(grupos[3])
                        preco = parse_valor(grupos[4])
                        valor_total = parse_valor(grupos[5]) if len(grupos) > 5 else quantidade * preco
                    
                    elif padrao == padroes[1]:  # COMPRA AÇÕES ITSA4 500 12,34 6.170,00
                        tipo = "C" if grupos[0].upper() == "COMPRA" else "V"
                        ativo = grupos[1]
                        quantidade = parse_valor(grupos[2])
                        preco = parse_valor(grupos[3])
                        valor_total = parse_valor(grupos[4]) if len(grupos) > 4 else quantidade * preco
                    
                    elif padrao == padroes[2]:  # 1 C ON VALE3 100 77,10 7.710,00
                        tipo = "C" if grupos[0] == "C" else "V"
                        ativo = grupos[1]
                        quantidade = parse_valor(grupos[2])
                        preco = parse_valor(grupos[3])
                        valor_total = parse_valor(grupos[4]) if len(grupos) > 4 else quantidade * preco
                    
                    elif padrao == padroes[3]:  # WIN N22 FUTURO | COMPRA | 2 | 115.180,00
                        tipo = "C" if "COMPRA" in grupos[1].upper() else "V"
                        ativo = grupos[0]
                        quantidade = parse_valor(grupos[2])
                        preco = parse_valor(grupos[3])
                        valor_total = preco * quantidade
                    
                    elif padrao == padroes[4]:  # DOL    FUTURO    COMPRA    5    5.050,00
                        tipo = "C" if "COMPRA" in grupos[1].upper() else "V"
                        ativo = grupos[0]
                        quantidade = parse_valor(grupos[2])
                        preco = parse_valor(grupos[3]) if len(grupos) > 3 else 0
                        valor_total = preco * quantidade
                    
                    elif padrao == padroes[5]:  # C WIN N22 5 119490.0
                        tipo = "C" if grupos[0] == "C" else "V"
                        ativo = grupos[1]
                        quantidade = parse_valor(grupos[2])
                        preco = parse_valor(grupos[3]) if len(grupos) > 3 else 0
                        valor_total = preco * quantidade
                    
                    elif padrao == padroes[6]:  # WIN N22 2 115180.0
                        tipo = "C"  # Default para C quando não especificado
                        ativo = grupos[0]
                        quantidade = parse_valor(grupos[1])
                        preco = parse_valor(grupos[2]) if len(grupos) > 2 and grupos[2] else 0
                        valor_total = preco * quantidade
                    
                    elif padrao == padroes[9]:  # C VISTA PETROBRAS PN N2 200 35.9
                        tipo = "C" if grupos[0] == "C" else "V"
                        empresa = grupos[1].strip()
                        tipo_ativo = grupos[2].strip()
                        ativo = f"{empresa} {tipo_ativo}"
                        quantidade = parse_valor(grupos[3])
                        preco = parse_valor(grupos[4])
                        valor_total = quantidade * preco
                    
                    elif padrao == padroes[10]:  # VISTA PETR4 PETROBRAS PN 100 35.85
                        tipo = "C"  # Default quando não especificado no padrão
                        ticker = grupos[0].strip()
                        empresa = grupos[1].strip()
                        tipo_ativo = grupos[2].strip()
                        ativo = f"{empresa} {tipo_ativo}"
                        quantidade = parse_valor(grupos[3])
                        preco = parse_valor(grupos[4])
                        valor_total = quantidade * preco
                    
                    elif padrao == padroes[11]:  # 1-BOVESPA C VISTA PETR4 (PETROBRAS PN) 100 35.85
                        tipo = "C" if grupos[0] == "C" else "V"
                        ticker = grupos[1].strip()
                        quantidade = parse_valor(grupos[2])
                        preco = parse_valor(grupos[3])
                        valor_total = quantidade * preco
                        ativo = ticker
                    
                    elif padrao == padroes[12]:  # BOVESPA PETROBRAS PN N2 PETR4 C 100 35.93
                        empresa = grupos[0].strip()
                        tipo_ativo = grupos[1].strip()
                        ticker = grupos[2].strip()
                        tipo = "C" if grupos[3] == "C" else "V"
                        quantidade = parse_valor(grupos[4])
                        preco = parse_valor(grupos[5])
                        valor_total = quantidade * preco
                        ativo = f"{empresa} {tipo_ativo}"
                    
                    elif padrao == padroes[13]:  # B3 RVLISTADO C VISTA PETROBRAS PNEDJN2 D 200 35,90 7.180,00 D
                        tipo = "C" if grupos[0] == "C" else "V"
                        empresa = grupos[1].strip()
                        quantidade = parse_valor(grupos[2])
                        preco = parse_valor(grupos[3])
                        valor_total = parse_valor(grupos[4]) if len(grupos) > 4 else quantidade * preco
                        ativo = empresa
                        ticker = "PETR4" if "PETROBRAS" in empresa else empresa.split()[0] if " " in empresa else empresa
                        print(f"Padrão 14 B3 RVLISTADO encontrado: {grupos}")
                        
                    elif padrao == padroes[14]:  # B3 RVLISTADO C VISTA PETROBRAS PNEDJN2 D 200 35,90
                        tipo = "C" if grupos[0] == "C" else "V"
                        empresa = grupos[1].strip()
                        quantidade = parse_valor(grupos[2])
                        preco = parse_valor(grupos[3])
                        valor_total = quantidade * preco
                        ativo = empresa
                        ticker = "PETR4" if "PETROBRAS" in empresa else empresa.split()[0] if " " in empresa else empresa
                        print(f"Padrão 15 B3 RVLISTADO encontrado: {grupos}")
                        
                    elif padrao == padroes[15]:  # Padrão mais genérico para B3
                        tipo = "C" if grupos[0] == "C" else "V"
                        empresa_completa = grupos[1].strip()
                        quantidade = parse_valor(grupos[2])
                        preco = parse_valor(grupos[3])
                        valor_total = quantidade * preco
                        
                        # Limpar o nome da empresa, removendo códigos de tipo (PN, ON, etc)
                        partes_empresa = []
                        for parte in empresa_completa.split():
                            if not (parte.startswith("PN") or parte.startswith("ON") or parte == "D" or parte == "D#"):
                                partes_empresa.append(parte)
                        
                        ativo = " ".join(partes_empresa)
                        ticker = "PETR4" if "PETROBRAS" in ativo else ativo.split()[0] if " " in ativo else ativo
                        print(f"Padrão 16 B3 Genérico encontrado: {grupos} - Ativo: {ativo}")
                        
                    # Padrão 17 - Contrato futuro completo: C WDO F25 02/01/2025 1 6.088,0000 DAY TRADE
                    elif padrao == padroes[16]:
                        tipo = "C" if grupos[0] == "C" else "V"
                        ativo_base = grupos[1].strip()  # WDO
                        vencimento = grupos[2].strip()   # F25
                        data = grupos[3].strip()         # 02/01/2025
                        quantidade = parse_valor(grupos[4])
                        preco = parse_valor(grupos[5])
                        day_trade = "DAY TRADE" in grupos[6].upper() if len(grupos) > 6 else False
                        valor_total = quantidade * preco
                        ativo = f"{ativo_base} {vencimento}"
                        print(f"Padrão 17 Contrato Futuro encontrado: {ativo} - {quantidade} x {preco} = {valor_total}")
                    
                    # Padrão 18 - Contrato futuro simplificado: C WDO F25 1 6.088,0000 DAY TRADE
                    elif padrao == padroes[17]:
                        tipo = "C" if grupos[0] == "C" else "V"
                        ativo_base = grupos[1].strip()  # WDO
                        vencimento = grupos[2].strip()   # F25
                        quantidade = parse_valor(grupos[3])
                        preco = parse_valor(grupos[4])
                        day_trade = "DAY TRADE" in grupos[5].upper() if len(grupos) > 5 else False
                        valor_total = quantidade * preco
                        ativo = f"{ativo_base} {vencimento}"
                        print(f"Padrão 18 Contrato Futuro simplificado encontrado: {ativo} - {quantidade} x {preco} = {valor_total}")
                    
                    # Padrão 19 - Contrato futuro genérico: C WDO F25 1 6.088,0000
                    elif padrao == padroes[18]:
                        tipo = "C" if grupos[0] == "C" else "V"
                        ativo_base = grupos[1].strip()  # WDO
                        vencimento = grupos[2].strip()   # F25
                        quantidade = parse_valor(grupos[3])
                        preco = parse_valor(grupos[4])
                        valor_total = quantidade * preco
                        ativo = f"{ativo_base} {vencimento}"
                        print(f"Padrão 19 Contrato Futuro genérico encontrado: {ativo} - {quantidade} x {preco} = {valor_total}")
                    
                    else:  # Caso genérico
                        continue
                    
                    # Só adicionar se tiver pelo menos ativo e quantidade
                    if ativo and quantidade > 0:
                        transacao = {
                            "tipo": tipo,
                            "ativo": ativo.strip(),
                            "quantidade": quantidade,
                            "preco": preco,
                            "valor_total": valor_total,
                            "ticker": ticker if 'ticker' in locals() else ativo.split()[0] if ' ' in ativo else ativo
                        }
                        transacoes.append(transacao)
                
            except Exception as e:
                print(f"Erro extraindo transação: {e}")
                continue
    
    # Considerar também seções específicas de transações em algumas corretoras
    secoes = buscar_secoes_transacoes(texto)
    if secoes:
        for linha in secoes:
            try:
                # Tentar processar linhas como "C VISTA PETR4 100 25,30"
                match = re.search(r'([CV])\s+(?:VISTA|OPCAO|TERMO)?\s+([A-Z0-9]+)\s+(\d+(?:\.\d+)?)\s+([\d,.]+)', linha)
                if match:
                    tipo = "C" if match.group(1) == "C" else "V"
                    ativo = match.group(2)
                    quantidade = parse_valor(match.group(3))
                    preco = parse_valor(match.group(4))
                    valor_total = quantidade * preco
                    
                    transacao = {
                        "tipo": tipo,
                        "ativo": ativo,
                        "quantidade": quantidade,
                        "preco": preco,
                        "valor_total": valor_total
                    }
                    transacoes.append(transacao)
            except:
                continue
    
    """Busca seções específicas que contêm transações em algumas corretoras"""
    secoes = []
    
    # Padrões para encontrar seções de transações
    padroes_secao = [
        r'(?is)Negócios\s+(?:realizados|realizados:).*?(?=\n\s*\n|$)',
        r'(?is)Títulos\s+Negociados.*?(?=\n\s*\n|$)',
        r'(?is)Resumo\s+(?:dos|dos\s+)(?:Negócios|investimentos).*?(?=\n\s*\n|$)',
        r'(?is)Operações\s+(?:realizadas|realizadas:).*?(?=\n\s*\n|$)',
        r'(?is)Negociação.*?(?=\n\s*\n|$)',
        r'(?is)\bBMF\b.*?(?=\n\s*\n|$)',  # Capturar qualquer seção com BMF
        r'(?is)\bBOVESPA\b.*?(?=\n\s*\n|$)',
        r'(?is)\bB3\b.*?(?=\n\s*\n|$)',  # Capturar qualquer seção com B3
        r'(?is)Mercado\s+(?:futuro|de\s+futuros|de\s+opções).*?(?=\n\s*\n|$)',
        r'(?is)(?:Ajuste|Posição|Liquidado).*?(?:WIN|DOL|IND).*?(?=\n\s*\n|$)',  # Especificamente para contratos futuros
        r'(?is)C/V\s+(?:Mercadoria|Tipo).*?(?=\n\s*\n|$)',  # Padrão comum em tabelas de transações
        r'(?is)BMF\s+(?:FUTURO|OPÇÃO|DAY TRADE).*?(?=\n\s*\n|$)',  # Padrão específico para BMF
        r'(?is)BMF\s+(?:CONTRATO|FUTURO|OPÇÃO).*?(?=\n\s*\n|$)',  # Padrão específico para BMF
    ]
    
    for padrao in padroes_secao:
        for match in re.finditer(padrao, texto):
            secoes.append(match.group(0))
    
    # Imprimir seções encontradas para diagnóstico
    if secoes:
        print(f"Encontradas {len(secoes)} seções de transações")
    
    if "BTG" in texto.upper() or "PACTUAL" in texto.upper():
        # Procurar pela seção de transações em formato de tabela
        match = re.search(r'(?:MERCADORIAS|AJUSTE|ESPECIFICAÇÃO|CONTRATOS)(.+?)(?:RESUMO FINANCEIRO|CUSTOS|TOTAL)', texto, re.DOTALL)
        if match:
            secao = match.group(1)
            # Adicionar cada linha que parece ter um ativo
            for linha in secao.split('\n'):
                if any(fut in linha.upper() for fut in ['WIN', 'DOL', 'IND', 'FUTURO']):
                    secoes.append(linha)
    
    return secoes


def extrair_taxas(texto):
    """Extrai taxas e valores da nota de corretagem"""
    taxas = {}
    
    # Padrões para diferentes taxas
    padroes_taxas = {
        'taxa_liquidacao': [
            r'taxa\s+de\s+liquida[cç][aã]o\s*:?\s*(?:R\$)?\s*([\d\.,]+)',
            r'liquida[cç][aã]o\s*:?\s*(?:R\$)?\s*([\d\.,]+)'
        ],
        'taxa_registro': [
            r'taxa\s+de\s+registro\s*:?\s*(?:R\$)?\s*([\d\.,]+)',
            r'registro\s*:?\s*(?:R\$)?\s*([\d\.,]+)'
        ],
        'emolumentos': [
            r'emolumentos\s*:?\s*(?:R\$)?\s*([\d\.,]+)'
        ],
        'taxa_operacional': [
            r'taxa\s+(?:de\s+)?operacional\s*:?\s*(?:R\$)?\s*([\d\.,]+)',
            r'operacional\s*:?\s*(?:R\$)?\s*([\d\.,]+)',
            r'taxa\s+(?:de\s+)?opera[çc][aã]o\s*:?\s*(?:R\$)?\s*([\d\.,]+)'
        ],
        'corretagem': [
            r'corretagem\s*:?\s*(?:R\$)?\s*([\d\.,]+)'
        ],
        'iss': [
            r'(?:imposto|i\.?s\.?s\.?)\s*:?\s*(?:R\$)?\s*([\d\.,]+)'
        ],
        'irrf': [
            r'(?:i\.?r\.?r\.?f\.?|imposto\s+de\s+renda)\s*:?\s*(?:R\$)?\s*([\d\.,]+)'
        ],
        'valor_liquido': [
            r'(?:valor|l[ií]quido)\s+(?:l[ií]quido|para|da\s+nota)(?:\s+\d{2}/\d{2}/\d{4})?\s*:?\s*(?:R\$)?\s*([\d\.,]+)',
            r'(?:total|l[ií]quido)\s+(?:l[ií]quido|para\s+liquidac[aã]o)(?:\s+\d{2}/\d{2}/\d{4})?\s*:?\s*(?:R\$)?\s*([\d\.,]+)'
        ],
        # Taxas adicionais específicas
        'taxa_ajuste': [
            r'(?:taxa\s+de\s+)?ajuste\s*:?\s*(?:R\$)?\s*([\d\.,]+)'
        ],
        'valor_operacao_dc': [
            r'valor\s+(?:de|da)?\s*opera[çc][aã]o\s*:?\s*(?:R\$)?\s*([\d\.,]+)',
            r'valor\s+(?:d/c|debito/credito)\s*:?\s*(?:R\$)?\s*([\d\.,]+)'
        ]
    }
    
    # Procurar cada taxa no texto
    for nome_taxa, padroes in padroes_taxas.items():
        for padrao in padroes:
            match = re.search(padrao, texto, re.IGNORECASE)
            if match:
                try:
                    valor_str = match.group(1).replace('.', '').replace(',', '.')
                    valor = float(valor_str)
                    taxas[nome_taxa] = valor
                    break
                except:
                    continue
    
    return taxas


def encontrar_coluna(cabecalho, possibilidades):
    """Encontra a coluna que melhor corresponde a uma das possibilidades"""
    # Primeira tentativa: correspondência exata
    for i, col in enumerate(cabecalho):
        col_lower = col.lower()
        for p in possibilidades:
            if p.lower() == col_lower:
                return i
    
    # Segunda tentativa: correspondência parcial
    for i, col in enumerate(cabecalho):
        col_lower = col.lower()
        for p in possibilidades:
            if p.lower() in col_lower:
                return i
    
    # Terceira tentativa: verificar se a coluna contém palavras-chave comuns
    keywords = ['titulo', 'ativo', 'papel', 'codigo', 'especificacao', 'especificação', 
                'qtd', 'quant', 'quantidade', 'preco', 'preço', 'valor', 'ajuste', 
                'tipo', 'c/v', 'compra', 'venda', 'cv', 'operacao', 'operação']
    
    for i, col in enumerate(cabecalho):
        col_lower = col.lower()
        for keyword in keywords:
            if keyword in col_lower:
                # Verificar se essa keyword está relacionada às possibilidades
                for p in possibilidades:
                    if keyword in p.lower():
                        return i
    
    return -1


def parse_valor(valor_str):
    """Converte string de valor para float"""
    if not valor_str:
        return 0
    
    # Remover R$, espaços e outros caracteres não numéricos
    valor_str = re.sub(r'[^\d,.-]', '', str(valor_str))
    
    # Padrão brasileiro: usar vírgula como decimal e ponto como separador de milhar
    if ',' in valor_str:
        # Se tem vírgula, é decimal brasileiro
        valor_str = valor_str.replace('.', '').replace(',', '.')
    
    try:
        return float(valor_str)
    except:
        return 0


# Padrões de identificação de corretoras
PADROES_CORRETORAS = {
    'xp': r'XP\s+INVESTIMENTOS|CORRETORA\s+XP',
    'clear': r'CLEAR\s+CORRETORA|CLEAR\s+CTVM',
    'rico': r'RICO\s+INVESTIMENTOS|RICO\s+CTVM',
    'modal': r'MODAL\s+DTVM|MODAL\s+MAIS',
    'inter': r'INTER\s+DTVM|BANCO\s+INTER',
    'guide': r'GUIDE\s+INVESTIMENTOS',
    'nuinvest': r'NU\s+INVEST|NUINVEST|EASYNVEST',
    'itau': r'ITA[Uu]\s+CORRETORA',
    'bradesco': r'BRADESCO\s+S/?A|BRADESCO\s+CORRETORA',
    'santander': r'SANTANDER\s+CORRETORA|SANTANDER\s+CTVM',
    'btg': r'BTG\s+PACTUAL',
    'genial': r'GENIAL\s+INVESTIMENTOS',
    'terra': r'TERRA\s+INVESTIMENTOS',
    'orama': r'ORAMA\s+DTVM',
    'necton': r'NECTON\s+INVESTIMENTOS',
    'nova_futura': r'NOVA\s+FUTURA\s+CTVM',
    'toro': r'TORO\s+INVESTIMENTOS',
    'c6': r'C6\s+CTVM|C6\s+BANK',
    'mirae': r'MIRAE\s+ASSET',
}

# Padrões para número da nota
PADROES_NUMERO_NOTA = [
    r'Nr\.\s*(?:nota|order|negoci):\s*(\d+)',
    r'N[o°º]\s*(?:da nota|nota):\s*(\d+)',
    r'N[uú]mero\s*(?:da nota|nota|folha):\s*(\d+)',
    r'(?:Nota|Folha)\s*(?:n[o°º]|n[uú]mero|\#):\s*(\d+)',
    r'(?:NOTA|BOLETA)\s*(?:DE CORRETAGEM|DE NEGOCIAÇÃO)\s*[^\d]*(\d+)',
    r'Nr\.?\s*Boleta:?\s*(\d+)',
    r'Boleta\s+Nº\s*(\d+)'
]

# Padrões para data da nota
PADROES_DATA = [
    r'(?:Data|Data pregão):\s*(\d{2}[/-]\d{2}[/-]\d{4}|\d{2}[/-]\d{2}[/-]\d{2})',
    r'(?:Date|Dia|Data):\s*(\d{2}[/-]\d{2}[/-]\d{4}|\d{2}[/-]\d{2}[/-]\d{2})',
    r'Pregão(?:\s+de)?\s*:?\s*(\d{2}[/-]\d{2}[/-]\d{4}|\d{2}[/-]\d{2}[/-]\d{2})',
    r'(?:Data|Date)\s*(?:de|da|do)?\s*(?:neg[oó]ci(?:o|ação)|operações)?\s*:?\s*(\d{2}[/-]\d{2}[/-]\d{4}|\d{2}[/-]\d{2}[/-]\d{2})',
    r'D\.?\s*Pregão:?\s*(\d{2}[/-]\d{2}[/-]\d{4}|\d{2}[/-]\d{2}[/-]\d{2})',
    r'(?:Data|Date)\s*Liquidação:?\s*(\d{2}[/-]\d{2}[/-]\d{4}|\d{2}[/-]\d{2}[/-]\d{2})'
]

# Função principal para análise
def analisar_pdf_nota_corretagem(caminho_pdf, modo_debug=False):
    """Função principal para análise de notas de corretagem"""
    try:
        # Extrair nota usando extractors
        nota = extrair_nota_corretagem(caminho_pdf, modo_debug)
        
        # Se estiver no modo debug, imprimir informações sobre as transações
        if modo_debug and nota.get("numero_nota"):
            print(f"\n===== DIAGNÓSTICO DETALHADO - NOTA {nota.get('numero_nota')} =====")
            print(f"Total de transações encontradas: {len(nota.get('transacoes', []))}")
            
            # Se não encontrou transações, tentar verificar o texto da nota para palavras-chave
            if not nota.get('transacoes'):
                with pdfplumber.open(caminho_pdf) as pdf:
                    texto_completo = ""
                    for pagina in pdf.pages:
                        texto_pagina = pagina.extract_text()
                        if texto_pagina:
                            texto_completo += texto_pagina + "\n"
                    
                    print("\nBuscando keywords importantes nas notas sem transações:")
                    for keyword in ["BMF", "B3", "BOVESPA", "WIN", "DOL", "IND", "FUTURO", "OPÇÃO", "VISTA", "MERCADO", "NEGOCIAÇÃO"]:
                        if keyword in texto_completo.upper():
                            print(f"- Encontrada keyword: {keyword}")
                            # Extrair uma amostra do texto ao redor dessa keyword para diagnóstico
                            pos = texto_completo.upper().find(keyword)
                            inicio = max(0, pos - 50)
                            fim = min(len(texto_completo), pos + len(keyword) + 100)
                            print(f"  Contexto: ...{texto_completo[inicio:fim]}...")
                    
                    # Tentar extrair algumas linhas importantes
                    print("\nLinhas potencialmente relevantes:")
                    for linha in texto_completo.split('\n'):
                        if any(kw in linha.upper() for kw in ["BMF", "B3", "BOVESPA", "NEGÓCIO", "REALIZADO", "C/V", "COMPRA", "VENDA", "AJUSTE"]):
                            print(f">> {linha}")
        
        # Verificar se há texto completo para análises adicionais
        with pdfplumber.open(caminho_pdf) as pdf:
            texto_completo = ""
            for pagina in pdf.pages:
                texto_pagina = pagina.extract_text()
                if texto_pagina:
                    texto_completo += texto_pagina + "\n"
    
        # Verificar se encontramos os campos adicionais nas taxas
        taxas = nota.get('taxas', {})
        
        # Se não encontrou os valores nas taxas, tentar outras abordagens
        if 'valor_operacao_dc' not in taxas or taxas['valor_operacao_dc'] == 0:
            # Tentar extrair Valor Operação de outros lugares
            match = re.search(r'valor\s+(?:de|da)?\s*opera[çc][aã]o\s*:?\s*(?:R\$)?\s*([\d\.,]+)', 
                              texto_completo, re.IGNORECASE)
            if match:
                try:
                    taxas['valor_operacao_dc'] = parse_valor(match.group(1))
                except:
                    pass
        
        # Procurar por ajustes se ainda não encontrou
        if 'taxa_ajuste' not in taxas or taxas['taxa_ajuste'] == 0:
            match = re.search(r'ajuste\s*:?\s*(?:R\$)?\s*([\d\.,]+)', texto_completo, re.IGNORECASE)
            if match:
                try:
                    taxas['taxa_ajuste'] = parse_valor(match.group(1))
                except:
                    pass
        
        # Procurar especificamente por transações de mercado futuro no formato da BTG
        # Exemplo: C WINJ25 16/04/2025 3 131.820,0000 DAY TRADE 82,80 C 0,00
        padrao_btg_futuro = r'([CV])\s+([A-Z0-9]+)\s+(\d{2}/\d{2}/\d{4})?\s*(\d+)\s+([\d\.,]+)\s+(DAY\s*TRADE|NORMAL)?\s*([\d\.,]+)?\s*([CD])?\s*([\d\.,]+)?'
        
        # Buscar esse padrão em todas as linhas do texto
        transacoes_btg = []
        for linha in texto_completo.split('\n'):
            match = re.search(padrao_btg_futuro, linha.strip(), re.IGNORECASE)
            if match:
                try:
                    # Extrair dados
                    tipo = 'C' if match.group(1) == 'C' else 'V'
                    ativo = match.group(2)
                    vencimento = match.group(3) if match.group(3) else ''
                    quantidade = int(match.group(4))
                    preco = parse_valor(match.group(5))
                    tipo_negocio = match.group(6) if match.group(6) else 'NORMAL'
                    valor_op = parse_valor(match.group(7)) if match.group(7) else 0
                    dc = match.group(8) if match.group(8) else ''
                    taxa_op = parse_valor(match.group(9)) if match.group(9) else 0
                    
                    # Criar transação
                    if ativo and quantidade > 0:
                        transacoes_btg.append({
                            "tipo": tipo,
                            "ativo": ativo,
                            "ticker": ativo,  # Mercadoria
                            "vencimento": vencimento,
                            "quantidade": quantidade,
                            "preco": preco,
                            "valor_total": preco * quantidade,
                            "tipo_negocio": tipo_negocio.strip().upper() if tipo_negocio else 'NORMAL',
                            "dc": dc,
                            "valor_operacao": valor_op,
                            "taxa_operacional": taxa_op
                        })
                except Exception as e:
                    print(f"Erro ao extrair transação de mercado futuro: {e}")
        
        # Se encontramos transações no formato BTG, as adicionamos às existentes ou 
        # substituímos se não houver nenhuma
        if transacoes_btg:
            if not nota.get('transacoes'):
                nota['transacoes'] = transacoes_btg
            else:
                nota['transacoes'].extend(transacoes_btg)
        
        # Adicionar os campos extras nas transações
        for transacao in nota.get('transacoes', []):
            # Garantir que todos os novos campos estejam presentes
            if 'tipo_negocio' not in transacao:
                transacao['tipo_negocio'] = ""
            if 'dc' not in transacao:
                transacao['dc'] = ""
            if 'preco_ajuste' not in transacao:
                transacao['preco_ajuste'] = transacao.get('preco', 0)
            if 'ticker' not in transacao:
                transacao['ticker'] = transacao.get('ativo', '')
            if 'vencimento' not in transacao:
                transacao['vencimento'] = ''
            if 'valor_operacao' not in transacao:
                transacao['valor_operacao'] = 0
            if 'taxa_operacional' not in transacao:
                transacao['taxa_operacional'] = 0
        
        return nota
    except Exception as e:
        print(f"Erro ao analisar nota de corretagem: {e}")
        return {
            "sucesso": False,
            "erro": str(e),
            "corretora": "Desconhecida",
            "numero_nota": None,
            "data_nota": None,
            "transacoes": [],
            "taxas": {},
            "resumo": {}
        }
    # Este código já está incorporado na função principal acima
