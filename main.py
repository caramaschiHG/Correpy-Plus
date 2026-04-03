import os
import io
import sys
import threading
import tkinter as tk
from tkinter import filedialog, ttk, scrolledtext, messagebox
import pandas as pd
from collections import defaultdict
import ttkthemes as themes
from datetime import datetime
import traceback
import warnings
import logging
from correpy.parsers.brokerage_notes.parser_factory import ParserFactory
import re
import pdfplumber

# Importar o módulo de extração de futuros
from extrair_futuros_direto import main as extrair_futuros_main

# Configurar para ignorar avisos específicos (como CropBox missing)
warnings.filterwarnings("ignore")

# Suprimir mensagens de log do pdfplumber
logging.getLogger("pdfminer").setLevel(logging.ERROR)

# Funções para extração de contratos futuros diretamente no main.py
def parse_valor(valor_str):
    """Converte string de valor para float"""
    if not valor_str:
        return 0
    
    # Remover caracteres não numéricos, exceto ponto e vírgula
    valor_limpo = re.sub(r'[^\d.,]', '', str(valor_str))
    
    # Se estiver vazio após limpeza
    if not valor_limpo:
        return 0
    
    # Remover pontos de milhar e substituir vírgula por ponto
    valor_limpo = valor_limpo.replace('.', '').replace(',', '.')
    
    try:
        return float(valor_limpo)
    except ValueError:
        return 0

def extrair_texto_pdf(caminho_pdf):
    """Extrai todo o texto de um arquivo PDF"""
    texto_completo = ""
    try:
        with pdfplumber.open(caminho_pdf) as pdf:
            for pagina in pdf.pages:
                texto = pagina.extract_text() or ""
                texto_completo += texto + "\n"
        return texto_completo
    except Exception as e:
        print(f"Erro ao extrair texto do PDF: {e}")
        return ""

def extrair_contratos_futuros(texto):
    """Extrai contratos futuros do texto"""
    transacoes = []
    
    # Verificar cada linha do texto
    for linha in texto.split('\n'):
        # Verificar padrões específicos para contratos futuros
        # Padrão: C WDO F25 02/01/2025 1 6.088,0000 DAY TRADE
        match = re.search(r'([CV])\s+([A-Z]{3})\s+([A-Z]\d{2}).*?(\d+)\s+([\d.,]+)', linha, re.IGNORECASE)
        if match:
            grupos = match.groups()
            tipo = grupos[0].upper()
            ativo_base = grupos[1].upper()
            vencimento = grupos[2].upper()
            quantidade = parse_valor(grupos[3])
            preco = parse_valor(grupos[4])
            valor_total = quantidade * preco
            
            # Criar transação
            transacao = {
                "tipo": tipo,
                "ativo": f"{ativo_base} {vencimento}",
                "quantidade": quantidade,
                "preco": preco,
                "valor_total": valor_total
            }
            transacoes.append(transacao)
            continue
        
        # Padrão genérico: buscar por WIN, WDO, DOL, IND seguidos de números
        for ativo in ["WIN", "WDO", "DOL", "IND"]:
            if ativo in linha.upper():
                match = re.search(r'([CV])\s+.*?(\d+).*?([\d.,]+)', linha, re.IGNORECASE)
                if match:
                    grupos = match.groups()
                    tipo = grupos[0].upper()
                    quantidade = parse_valor(grupos[1])
                    preco = parse_valor(grupos[2])
                    valor_total = quantidade * preco
                    
                    # Extrair vencimento (código como F25, G25, etc)
                    vencimento = ""
                    for termo in linha.split():
                        if re.match(r'[A-Z]\d{2}', termo, re.IGNORECASE):
                            vencimento = termo.upper()
                            break
                    
                    # Criar transação
                    transacao = {
                        "tipo": tipo,
                        "ativo": f"{ativo} {vencimento}" if vencimento else ativo,
                        "quantidade": quantidade,
                        "preco": preco,
                        "valor_total": valor_total
                    }
                    
                    # Verificar se não é duplicata
                    duplicata = False
                    for t in transacoes:
                        if (t["ativo"] == transacao["ativo"] and 
                            t["quantidade"] == transacao["quantidade"] and 
                            t["preco"] == transacao["preco"] and
                            t["tipo"] == transacao["tipo"]):
                            duplicata = True
                            break
                    
                    if not duplicata:
                        transacoes.append(transacao)
                    
                    break
    
    return transacoes

# Importar analisadores de PDF personalizados (do menos para o mais avançado)
try:
    from pdf_analyzer import analisar_pdf_nota_corretagem
    PDF_ANALYZER_DISPONIVEL = True
except ImportError:
    PDF_ANALYZER_DISPONIVEL = False

try:
    from advanced_parser import analisar_pdf_nota_corretagem as analisar_pdf_avancado
    ADVANCED_PARSER_DISPONIVEL = True
except ImportError:
    ADVANCED_PARSER_DISPONIVEL = False

# Importar o extrator direto e simplificado (maior prioridade)
try:
    from extrator_notas import analisar_pdf_nota_corretagem as extrair_nota_direto
    EXTRATOR_DIRETO_DISPONIVEL = True
except ImportError:
    EXTRATOR_DIRETO_DISPONIVEL = False

# Verificar se estamos executando como um arquivo .exe compilado
if getattr(sys, 'frozen', False):
    # Executando como arquivo compilado (.exe)
    application_path = os.path.dirname(sys.executable)
os.chdir(os.path.dirname(os.path.abspath(__file__)))

# Configurações globais com paleta moderna e mais vibrante
CORES = {
    "bg_escuro": "#1a1b26",       # Fundo principal mais escuro
    "bg_medio": "#24283b",        # Fundo dos painéis um pouco mais claro
    "bg_claro": "#414868",        # Fundo dos controles mais claro
    "texto": "#c0caf5",           # Texto principal com tom azulado suave
    "destaque": "#7aa2f7",        # Azul vibrante para destaques
    "sucesso": "#9ece6a",         # Verde mais vibrante para sucesso
    "erro": "#f7768e",            # Vermelho mais vibrante para erros
    "alerta": "#e0af68",          # Amarelo mais vibrante para alertas
    "roxo": "#bb9af7",            # Roxo para elementos especiais
    "cyan": "#7dcfff"             # Cyan para elementos de destaque secundário
}

class LogHandler:
    def __init__(self, text_widget):
        self.text_widget = text_widget
        
    def log(self, mensagem, tipo="normal"):
        timestamp = datetime.now().strftime("%H:%M:%S")
        if tipo == "erro":
            tag = "erro"
            prefixo = "❌ "
        elif tipo == "sucesso":
            tag = "sucesso"
            prefixo = "✅ "
        elif tipo == "alerta":
            tag = "alerta"
            prefixo = "⚠️ "
        elif tipo == "info":
            tag = "info"
            prefixo = "ℹ️ "
        else:
            tag = "normal"
            prefixo = "   "
            
        texto_formatado = f"[{timestamp}] {prefixo}{mensagem}\n"
        self.text_widget.config(state=tk.NORMAL)
        self.text_widget.insert(tk.END, texto_formatado, tag)
        self.text_widget.see(tk.END)
        self.text_widget.config(state=tk.DISABLED)
        self.text_widget.update()

# Função para processar um único arquivo PDF
def processar_arquivo_pdf(caminho_pdf, dados_por_mes, logger):
    total_notas = 0
    total_transacoes = 0
    nome_arquivo = os.path.basename(caminho_pdf)
    uso_analisador_custom = False
    
    try:
        # Primeiro, tenta com o correpy (biblioteca padrão)
        with open(caminho_pdf, 'rb') as f:
            conteudo = io.BytesIO(f.read())
            conteudo.seek(0)
            
            notas = ParserFactory(brokerage_note=conteudo).parse()
            arquivo_notas = len(notas)
            total_notas += arquivo_notas
            
            logger.log(f"  → Encontradas {arquivo_notas} notas em {nome_arquivo}", "info")
            
            # Verificar se há transações em alguma das notas
            total_transacoes_correpy = sum(len(nota.transactions) for nota in notas)
            
            # Se não há transações e algum analisador customizado está disponível, usar ele como fallback
            if total_transacoes_correpy == 0 and (PDF_ANALYZER_DISPONIVEL or ADVANCED_PARSER_DISPONIVEL or EXTRATOR_DIRETO_DISPONIVEL):
                logger.log("  → Nenhuma transação encontrada com correpy. Tentando com analisadores customizados...", "info")
                resultado_analise = tentar_analisador_customizado(caminho_pdf, logger)
                
                if resultado_analise and resultado_analise.get("transacoes"):
                    uso_analisador_custom = True
                    # Usar os dados do analisador customizado
                    processar_resultado_customizado(resultado_analise, dados_por_mes, logger)
                    total_notas = 1  # Consideramos uma nota bem-sucedida
                    total_transacoes = len(resultado_analise.get("transacoes", []))
                    
                    return True, total_notas, total_transacoes
            
            # Se chegou aqui, ou não usou o analisador customizado, ou ele não encontrou transações também
            # Continuar com o processamento normal do correpy
            for i, nota in enumerate(notas):
                # Debug de informações da nota
                logger.log(f"  → Detalhes Nota #{i+1}:", "info")
                logger.log(f"     - ID: {nota.reference_id}")
                logger.log(f"     - Data: {nota.reference_date}")
                logger.log(f"     - Corretora: {nota.brokerage_firm if hasattr(nota, 'brokerage_firm') else 'N/A'}")
                logger.log(f"     - Total de transações: {len(nota.transactions)}")
                
                # Se não há transações, vamos tentar extrair contratos futuros desta nota
                if len(nota.transactions) == 0:
                    # Tentativa de extração de contratos futuros para cada nota individual
                    logger.log("     - Tentando extrair contratos futuros nesta nota...", "alerta")
                    
                    try:
                        # Extrair texto do PDF para esta nota
                        texto_pdf = extrair_texto_pdf(caminho_pdf)
                        
                        # Usar a função aprimorada de detecção de contratos futuros
                        contratos_encontrados = extrair_contratos_futuros(texto_pdf)
                        
                        if contratos_encontrados and len(contratos_encontrados) > 0:
                            logger.log(f"     - Sucesso! Encontrados {len(contratos_encontrados)} contratos futuros!", "sucesso")
                            
                            # Mostrar detalhes dos contratos encontrados
                            for i, contrato in enumerate(contratos_encontrados[:5], 1):
                                mes_info = f" ({contrato.get('mes_vencimento', '')})" if contrato.get('mes_vencimento') else ""
                                logger.log(f"       → Contrato #{i}: {contrato['tipo']} {contrato['ativo']}{mes_info} - {contrato['quantidade']} x {contrato['preco']}", "info")
                            
                            if len(contratos_encontrados) > 5:
                                logger.log(f"       → ... e mais {len(contratos_encontrados) - 5} contratos", "info")
                            
                            # Verificar se já existem transações semelhantes antes de adicionar
                            # Lista para manter registro das transações já existentes
                            transacoes_existentes = []
                            for t in nota.transactions:
                                # Extrair informações da transação existente
                                chave_transacao = (t.get('asset', ''), 
                                                   t.get('quantity', 0), 
                                                   t.get('price', 0), 
                                                   t.get('operation', ''))
                                transacoes_existentes.append(chave_transacao)
                                
                            # Contador de novas transações adicionadas
                            novas_transacoes = 0
                            
                            # Adicionar os contratos encontrados como transações, evitando duplicatas
                            for contrato in contratos_encontrados:
                                # Determinar a operação
                                operacao = 'buy' if contrato['tipo'] == 'C' else 'sell'
                                
                                # Verificar se já existe uma transação idêntica
                                chave_nova = (contrato['ativo'], 
                                             contrato['quantidade'], 
                                             contrato['preco'], 
                                             operacao)
                                
                                # Só adiciona se não for duplicata
                                if chave_nova not in transacoes_existentes:
                                    # Converter para o formato de transação do correpy
                                    transaction = {
                                        'asset': contrato['ativo'],
                                        'quantity': contrato['quantidade'],
                                        'price': contrato['preco'],
                                        'operation': operacao,
                                        # Adicionar informações extras
                                        'market_type': 'future',
                                        'expiration': contrato.get('vencimento', ''),
                                        'expiration_month': contrato.get('mes_vencimento', '')
                                    }
                                    nota.transactions.append(transaction)
                                    transacoes_existentes.append(chave_nova)
                                    novas_transacoes += 1
                            
                            # Atualizar o log com as novas informações
                            logger.log(f"     - Total de transações agora: {len(nota.transactions)}", "sucesso")
                            logger.log(f"     - Novas transações adicionadas: {novas_transacoes}", "info")
                            
                            # Agora processar os contratos futuros encontrados
                            data_referencia = nota.reference_date
                            mes_ano = data_referencia.strftime('%Y_%m')
                            
                            # Garantir que o mês existe no dicionário
                            if mes_ano not in dados_por_mes:
                                dados_por_mes[mes_ano] = []
                            
                            # Verificar transacoes já existentes no Excel para evitar duplicatas
                            registros_existentes = []
                            for reg in dados_por_mes[mes_ano]:
                                # Criar uma chave única para cada registro já existente
                                try:
                                    data = reg.get('Data')
                                    if isinstance(data, datetime):
                                        data_str = data.strftime('%Y-%m-%d')
                                    else:
                                        data_str = str(data)
                                        
                                    chave_excel = (data_str,
                                               str(reg.get('Número da Nota', '')),
                                               str(reg.get('Tipo de Transação', '')),
                                               str(reg.get('Quantidade', 0)),
                                               str(reg.get('Preço Unitário', 0)),
                                               str(reg.get('Ativo', '')))
                                    registros_existentes.append(chave_excel)
                                except Exception:
                                    # Ignorar erros na criação de chaves
                                    pass
                            
                            # Contador de registros adicionados ao Excel
                            novos_registros = 0
                                
                            # Adicionar cada contrato ao relatório, evitando duplicatas
                            for contrato in contratos_encontrados:
                                # Determinar o tipo de operação (formatado para exibição)
                                if contrato['tipo'] == 'C':
                                    tipo_operacao = 'COMPRA'
                                elif contrato['tipo'] == 'V':
                                    tipo_operacao = 'VENDA'
                                else:
                                    tipo_operacao = contrato['tipo']
                                
                                # Formatar o nome do ativo com informações adicionais
                                ativo_info = contrato['ativo']
                                if contrato.get('mes_vencimento') and contrato.get('vencimento') and len(contrato.get('vencimento')) > 1:
                                    ativo_info = f"{ativo_info} ({contrato['mes_vencimento']}/{contrato['vencimento'][1:]})"
                                
                                # Criar chave para verificar duplicidade
                                data_str = data_referencia.strftime('%Y-%m-%d')
                                chave_novo_registro = (data_str,
                                                     str(nota.reference_id),
                                                     tipo_operacao,
                                                     str(contrato['quantidade']),
                                                     str(contrato['preco']),
                                                     ativo_info)
                                
                                # Só adiciona se não for duplicata
                                if chave_novo_registro not in registros_existentes:
                                    registro = {
                                        'Data': data_referencia,
                                        'Número da Nota': nota.reference_id,
                                        'Tipo de Transação': tipo_operacao,
                                        'Quantidade': contrato['quantidade'],
                                        'Preço Unitário': contrato['preco'],
                                        'Valor': contrato['valor_total'],
                                        'Ativo': ativo_info,
                                        'Tipo de Mercado': 'Futuro',
                                        'Código de Vencimento': contrato.get('vencimento', ''),
                                        'Mês de Vencimento': contrato.get('mes_vencimento', '')
                                    }
                                    
                                    # Adicionar campos padrão para relatório
                                    for campo in ['Taxa de Liquidação', 'Taxa de Registro', 'Taxa de Termo/Opções',
                                                'Taxa A.N.A', 'Emolumentos', 'Taxa Operacional', 'Execução',
                                                'Corretagem', 'ISS', 'IRRF Retido na Fonte', 'Taxa de Custódia', 'Impostos']:
                                        registro[campo] = 0.0
                                    
                                    dados_por_mes[mes_ano].append(registro)
                                    registros_existentes.append(chave_novo_registro)
                                    novos_registros += 1
                            
                            logger.log(f"     - Registros adicionados ao Excel: {novos_registros}", "info")
                                
                            # Atualizar contadores
                            total_transacoes += len(contratos_encontrados)
                            continue  # Pular para a próxima nota
                    except Exception as e:
                        logger.log(f"     - Erro ao extrair contratos futuros: {str(e)}", "erro")
                    
                    # Se continuou aqui, não encontrou contratos futuros
                    logger.log("     - Alerta: Nenhuma transação encontrada nesta nota", "alerta")
                    
                    # Adicionar dados da nota sem transações no relatório
                    data_referencia = nota.reference_date
                    mes_ano = data_referencia.strftime('%Y_%m')
                    
                    # Criar uma entrada com informações básicas da nota
                    # Verificar cada atributo antes de acessá-lo para evitar erros
                    registro = {
                        'Data': data_referencia,
                        'Número da Nota': nota.reference_id,
                        'Tipo de Transação': 'SEM TRANSAÇÕES',
                        'Quantidade': 0,
                        'Preço Unitário': 0.0,
                        'Ativo': 'N/A',
                    }
                    
                    # Adicionar taxas e valores apenas se existirem
                    campos = [
                        ('Taxa de Liquidação', 'settlement_fee'),
                        ('Taxa de Registro', 'registration_fee'),
                        ('Taxa de Termo/Opções', 'term_fee'),
                        ('Taxa A.N.A', 'ana_fee'),
                        ('Emolumentos', 'emoluments'),
                        ('Taxa Operacional', 'operational_fee'),
                        ('Execução', 'execution'),
                        ('Taxa de Custódia', 'custody_fee'),
                        ('IRRF Retido na Fonte', 'source_withheld_taxes'),
                        ('Impostos', 'taxes'),
                        ('Outros', 'others')
                    ]
                    
                    for nome_campo, atributo in campos:
                        if hasattr(nota, atributo):
                            try:
                                valor = getattr(nota, atributo)
                                registro[nome_campo] = float(valor) if valor is not None else 0.0
                            except (ValueError, TypeError):
                                registro[nome_campo] = 0.0
                        else:
                            registro[nome_campo] = 0.0
                            
                    # Registrar os valores obtidos no log para debug
                    taxas_valores = ", ".join([f"{k}: {v}" for k, v in registro.items() 
                                             if k not in ['Data', 'Número da Nota', 'Tipo de Transação', 'Quantidade', 'Preço Unitário', 'Ativo']])
                    logger.log(f"     - Taxas/Valores: {taxas_valores}", "info")
                    
                    dados_por_mes[mes_ano].append(registro)
                else:
                    # Processamento normal para notas com transações
                    data_referencia = nota.reference_date
                    mes_ano = data_referencia.strftime('%Y_%m')
                    nota_transacoes = len(nota.transactions)
                    total_transacoes += nota_transacoes
                    
                    # Debug das transações
                    logger.log(f"     - Transações detectadas: {nota_transacoes}")
                    
                    for j, transacao in enumerate(nota.transactions):
                        logger.log(f"       → Transação #{j+1}: {transacao.security.name} - {transacao.transaction_type.name} - {transacao.amount} x {float(transacao.unit_price)}")
                        
                        # Usar a mesma abordagem defensiva para evitar erros com atributos ausentes
                        registro = {
                            'Data': data_referencia,
                            'Número da Nota': nota.reference_id,
                            'Tipo de Transação': transacao.transaction_type.name,
                            'Quantidade': transacao.amount,
                            'Preço Unitário': float(transacao.unit_price),
                            'Ativo': transacao.security.name,
                        }
                        
                        # Adicionar taxas e valores apenas se existirem
                        campos = [
                            ('Taxa de Liquidação', 'settlement_fee'),
                            ('Taxa de Registro', 'registration_fee'),
                            ('Taxa de Termo/Opções', 'term_fee'),
                            ('Taxa A.N.A', 'ana_fee'),
                            ('Emolumentos', 'emoluments'),
                            ('Taxa Operacional', 'operational_fee'),
                            ('Execução', 'execution'),
                            ('Taxa de Custódia', 'custody_fee'),
                            ('IRRF Retido na Fonte', 'source_withheld_taxes'),
                            ('Impostos', 'taxes'),
                            ('Outros', 'others')
                        ]
                        
                        for nome_campo, atributo in campos:
                            if hasattr(nota, atributo):
                                try:
                                    valor = getattr(nota, atributo)
                                    registro[nome_campo] = float(valor) if valor is not None else 0.0
                                except (ValueError, TypeError):
                                    registro[nome_campo] = 0.0
                            else:
                                registro[nome_campo] = 0.0
                                
                        dados_por_mes[mes_ano].append(registro)
        return True, total_notas, total_transacoes
    except Exception as e:
        logger.log(f"Erro ao processar {nome_arquivo} com correpy: {str(e)}", "erro")
        logger.log(f"Detalhes do erro: {traceback.format_exc()}", "erro")
        
        # Se falhou com correpy, tenta com os analisadores customizados disponíveis
        if (PDF_ANALYZER_DISPONIVEL or ADVANCED_PARSER_DISPONIVEL or EXTRATOR_DIRETO_DISPONIVEL) and not uso_analisador_custom:
            logger.log("  → Tentando processar com analisadores customizados (modo diagnóstico ativado)...", "info")
            resultado_analise = tentar_analisador_customizado(caminho_pdf, logger, modo_debug=True)
            
            # Sempre verificar se existem contratos futuros no PDF, independentemente do resultado anterior
            logger.log("Verificando por contratos futuros diretamente no PDF...", "info")
            try:
                # Usar a função importada do módulo extrair_futuros_direto
                transacoes_futuros = extrair_futuros_main(caminho_pdf)
                if transacoes_futuros:
                    logger.log(f"Encontrados {len(transacoes_futuros)} contratos futuros!", "sucesso")
                    # Se não temos resultado anterior, criar um novo
                    if not resultado_analise:
                        resultado_analise = {
                            "data_nota": datetime.now().strftime("%d/%m/%Y"),
                            "numero_nota": "AUTO",
                            "transacoes": transacoes_futuros,
                            "taxas": {}
                        }
                    # Se já temos um resultado mas sem transações, adicionar os contratos
                    elif not resultado_analise.get("transacoes"):
                        resultado_analise["transacoes"] = transacoes_futuros
                    # Se já existem transações, adicionar os contratos como novas transações
                    else:
                        # Verificar se já não são as mesmas transações
                        transa_existentes = resultado_analise.get("transacoes", [])
                        for contrato in transacoes_futuros:
                            if not any(t.get('ativo') == contrato.get('ativo') and 
                                    t.get('quantidade') == contrato.get('quantidade') and 
                                    t.get('preco') == contrato.get('preco') and
                                    t.get('tipo') == contrato.get('tipo') for t in transa_existentes):
                                transa_existentes.append(contrato)
                        resultado_analise["transacoes"] = transa_existentes
                        logger.log(f"Total agora: {len(transa_existentes)} transações combinadas", "sucesso")
            except Exception as e:
                logger.log(f"Erro ao extrair contratos futuros: {str(e)}", "erro")

            # Usar os dados do analisador customizado quando houver resultado válido
            if resultado_analise:
                processar_resultado_customizado(resultado_analise, dados_por_mes, logger)
                total_notas = 1  # Consideramos uma nota bem-sucedida
                total_transacoes = len(resultado_analise.get("transacoes", []))
                return True, total_notas, total_transacoes

        return False, 0, 0

# Função para tentar ler PDF com analisador customizado
def tentar_analisador_customizado(caminho_pdf, logger, modo_debug=True):
    # Primeiro tenta com o extrator direto (maior prioridade)
    if EXTRATOR_DIRETO_DISPONIVEL:
        try:
            logger.log("Tentando extrair informações com o extrator direto...", "info")
            # Ativar modo de diagnóstico para obter mais informações
            resultado = extrair_nota_direto(caminho_pdf, modo_debug=modo_debug)
            
            if resultado and resultado.get("sucesso"):
                logger.log("  → Extrator direto conseguiu extrair dados da nota", "sucesso")
                # Mostrar informações básicas da nota
                logger.log(f"     - Número da Nota: {resultado.get('numero_nota', 'N/A')}")
                logger.log(f"     - Data: {resultado.get('data_nota', 'N/A')}")
                logger.log(f"     - Corretora: {resultado.get('corretora', 'N/A')}")
                logger.log(f"     - Transações detectadas: {len(resultado.get('transacoes', []))}")
                
                # Verificar transações
                transacoes = resultado.get('transacoes', [])
                logger.log(f"        - Total de transações: {len(transacoes)}")
                
                # Se não houver transações, tentar extrair contratos futuros diretamente
                if not transacoes:
                    logger.log("     - Tentando detectar contratos futuros (BMF, WIN, WDO, DOL)...", "alerta")
                    try:
                        # Extrair texto do PDF
                        texto_pdf = extrair_texto_pdf(caminho_pdf)
                        # Buscar contratos futuros no texto
                        transacoes_futuros = extrair_contratos_futuros(texto_pdf)
                        
                        if transacoes_futuros:
                            logger.log(f"     - Encontrados {len(transacoes_futuros)} contratos futuros!", "sucesso")
                            # Adicionar os contratos futuros às transações da nota
                            transacoes = transacoes_futuros
                            resultado['transacoes'] = transacoes_futuros
                        else:
                            logger.log("     - Alerta: Nenhuma transação encontrada nesta nota", "alerta")
                    except Exception as e:
                        logger.log(f"     - Erro ao tentar extrair contratos futuros: {str(e)}", "erro")
                        logger.log("     - Alerta: Nenhuma transação encontrada nesta nota", "alerta")
                
                # Mostrar transações encontradas (seja pelo parser original ou pela nossa função)
                if transacoes:
                    logger.log(f"        - Transações detectadas: {len(transacoes)}")
                    for j, transacao in enumerate(transacoes):
                        ativo = transacao.get('ativo', 'N/A')
                        quantidade = transacao.get('quantidade', 0)
                        valor_total = transacao.get('valor_total', 0)
                        tipo = transacao.get('tipo', 'N/A')
                        tipo_texto = 'Compra' if tipo == 'C' else 'Venda' if tipo == 'V' else 'Outro'
                        preco = transacao.get('preco', 0)
                        logger.log(f"          → Transação #{j+1}: {ativo} - {tipo_texto} - {quantidade} x {preco} = {valor_total:.2f}")
                
                # Mostrar taxas e valores
                taxas = resultado.get('taxas', {})
                if taxas:
                    logger.log("     - Taxas e valores:")
                    for nome, valor in taxas.items():
                        if valor > 0:
                            nome_formatado = nome.replace('_', ' ').title()
                            logger.log(f"       → {nome_formatado}: R$ {valor:.2f}")
                
                return resultado
            else:
                logger.log("     - Extrator direto não conseguiu extrair dados, tentando método alternativo...", "alerta")
        except Exception as e:
            logger.log(f"Erro ao usar extrator direto: {str(e)}", "erro")
    
    # Segundo tenta com o analisador avançado se disponível
    if ADVANCED_PARSER_DISPONIVEL:
        try:
            logger.log("Tentando extrair informações com o analisador avançado...", "info")
            resultado = analisar_pdf_avancado(caminho_pdf)
            
            if resultado and resultado.get("sucesso"):
                logger.log("  → Analisador avançado conseguiu extrair dados da nota", "sucesso")
                # Mostrar informações básicas da nota
                logger.log(f"     - Número da Nota: {resultado.get('numero_nota', 'N/A')}")
                logger.log(f"     - Data: {resultado.get('data_nota', 'N/A')}")
                logger.log(f"     - Corretora: {resultado.get('corretora', 'N/A')}")
                logger.log(f"     - Transações detectadas: {len(resultado.get('transacoes', []))}")
                
                # Listar transações encontradas
                for i, transacao in enumerate(resultado.get('transacoes', [])):
                    tipo = transacao.get('tipo', 'N/A')
                    ativo = transacao.get('ativo', 'N/A')
                    qtd = transacao.get('quantidade', 0)
                    preco = transacao.get('preco', 0)
                    valor_total = transacao.get('valor_total', 0)
                    logger.log(f"       → Transação #{i+1}: {ativo} - {'Compra' if tipo == 'C' else 'Venda'} - {qtd} x {preco} = {valor_total:.2f}")
                
                return resultado
            else:
                logger.log("     - Analisador avançado não conseguiu extrair dados, tentando método alternativo...", "alerta")
        except Exception as e:
            logger.log(f"Erro ao usar analisador avançado: {str(e)}", "erro")
            
    # Se não conseguiu com os anteriores, tenta com o analisador básico
    if PDF_ANALYZER_DISPONIVEL:
        try:
            logger.log("Tentando extrair informações com o analisador básico...", "info")
            resultado = analisar_pdf_nota_corretagem(caminho_pdf)
            
            if resultado and resultado.get("sucesso"):
                logger.log("  → Analisador básico conseguiu extrair dados da nota", "sucesso")
                # Mostrar informações básicas da nota
                logger.log(f"     - Número da Nota: {resultado.get('numero_nota', 'N/A')}")
                logger.log(f"     - Data: {resultado.get('data_nota', 'N/A')}")
                logger.log(f"     - Corretora: {resultado.get('corretora', 'N/A')}")
                logger.log(f"     - Transações detectadas: {len(resultado.get('transacoes', []))}")
                
                # Listar transações encontradas
                for i, transacao in enumerate(resultado.get('transacoes', [])):
                    tipo = transacao.get('tipo', 'N/A')
                    ativo = transacao.get('ativo', 'N/A')
                    qtd = transacao.get('quantidade', 0)
                    preco = transacao.get('preco', 0)
                    logger.log(f"       → Transação #{i+1}: {ativo} - {'Compra' if tipo == 'C' else 'Venda'} - {qtd} x {preco}")
                
                return resultado
            else:
                logger.log("     - Nenhum analisador conseguiu extrair dados", "alerta")
                return None
        except Exception as e:
            logger.log(f"Erro ao usar analisador básico: {str(e)}", "erro")
            return None
    else:
        logger.log("Nenhum analisador de PDF está disponível.", "erro")
        return None

# Função para detectar contratos futuros diretamente do PDF
def detectar_contratos_futuros(caminho_pdf, logger):
    try:
        logger.log("Verificando especificamente por contratos futuros (WDO, WIN, DOL, IND)...", "info")
        
        # Extrair texto do PDF usando a função importada
        texto_completo = extrair_texto_pdf(caminho_pdf)
        
        transacoes_futuros = []
        
        # Lista de símbolos de contratos futuros comuns
        ativos_futuros = ["WIN", "WDO", "DOL", "IND", "BGI", "CCM", "ICF"]
        
        # Verificar linha por linha
        for linha in texto_completo.split('\n'):
            linha_upper = linha.upper()
            
            # Pular linhas muito pequenas
            if len(linha.strip()) < 5:
                continue
                
            # Verificar se a linha contém algum dos ativos futuros
            if any(ativo in linha_upper for ativo in ativos_futuros):
                logger.log(f"Encontrada linha com contrato futuro: {linha}", "info")
                
                # Padrões para detectar contratos futuros
                padroes = [
                    # C WDO F25 02/01/2025 1 6.088,0000 DAY TRADE
                    r'([CV])\s+([A-Z]{3})\s+([A-Z]\d{2}).*?(\d+)\s+([\d.,]+)',
                    # Padrão mais genérico para capturar mais casos
                    r'([CV])\s+([A-Z]{3}).*?(\d+).*?([\d.,]+)'
                ]
                
                for padrao in padroes:
                    match = re.search(padrao, linha, re.IGNORECASE)
                    if match:
                        grupos = match.groups()
                        try:
                            tipo = "C" if grupos[0].upper() == "C" else "V"
                            ativo_base = grupos[1].upper()  # WDO, WIN, DOL, etc.
                            
                            # Extrair vencimento se disponível
                            if len(grupos) > 2 and re.match(r'[A-Z]\d{2}', grupos[2]):
                                vencimento = grupos[2].upper()
                                quantidade = int(re.sub(r'[^\d]', '', grupos[3]))
                                preco_str = re.sub(r'[^\d.,]', '', grupos[4])
                                preco = float(preco_str.replace('.', '').replace(',', '.'))                            
                            else:
                                # Para o padrão mais genérico
                                quantidade = int(re.sub(r'[^\d]', '', grupos[2]))
                                preco_str = re.sub(r'[^\d.,]', '', grupos[3])
                                preco = float(preco_str.replace('.', '').replace(',', '.'))
                                vencimento = ""
                                # Tentar extrair vencimento do contexto
                                for termo in linha.split():
                                    if re.match(r'[A-Z]\d{2}', termo):
                                        vencimento = termo
                                        break
                            
                            # Formar nome do ativo
                            ativo = f"{ativo_base} {vencimento}" if vencimento else ativo_base
                            
                            # Calcular valor total
                            valor_total = quantidade * preco
                            
                            # Criar transação
                            transacao = {
                                "tipo": tipo,
                                "ativo": ativo.strip(),
                                "quantidade": quantidade,
                                "preco": preco,
                                "valor_total": valor_total
                            }
                            
                            # Adicionar à lista se não for duplicata
                            if not any(t.get('ativo') == transacao.get('ativo') and 
                                    t.get('quantidade') == transacao.get('quantidade') and
                                    t.get('preco') == transacao.get('preco') and
                                    t.get('tipo') == transacao.get('tipo') for t in transacoes_futuros):
                                transacoes_futuros.append(transacao)
                                logger.log(f"Contrato futuro detectado: {tipo} {ativo} - {quantidade} x {preco} = {valor_total:.2f}", "sucesso")
                        except Exception as e:
                            logger.log(f"Erro ao processar contrato futuro: {e}", "erro")
        
        logger.log(f"Total de contratos futuros detectados: {len(transacoes_futuros)}", "info")
        return transacoes_futuros
    except Exception as e:
        logger.log(f"Erro ao detectar contratos futuros: {e}", "erro")
        return []

# Função para processar resultado do analisador customizado
def processar_resultado_customizado(resultado, dados_por_mes, logger):
    """Processa o resultado do analisador customizado"""
    if not resultado:
        return False
    
    try:
        # Extrair data
        data_nota = resultado.get("data_nota")
        if data_nota:
            try:
                if isinstance(data_nota, str):
                    # Se for string, tentar converter para data
                    # Verificar formato da data e converter
                    if "/" in data_nota:
                        dia, mes, ano = data_nota.split("/")
                        data = datetime(int(ano), int(mes), int(dia))
                    else:
                        data = datetime.strptime(data_nota, "%Y-%m-%d")
                else:
                    # Assumir que já é um objeto date/datetime
                    data = data_nota
                    
                mes_ano = data.strftime("%Y_%m")
            except Exception as e:
                logger.log(f"Erro ao processar data da nota: {e}", "erro")
                mes_ano = "sem_data"
                data = datetime.now()
        else:
            mes_ano = "sem_data"
            data = datetime.now()
            
        # Garantir que o mês existe no dicionário
        if mes_ano not in dados_por_mes:
            dados_por_mes[mes_ano] = []
            
        # Número da nota
        numero_nota = resultado.get("numero_nota", "N/A")
        
        # Processar transações
        transacoes = resultado.get('transacoes', [])
        taxas = resultado.get('taxas', {})
        
        # Verificar e corrigir preços grandes que possam estar sem decimal
        for transacao in transacoes:
            if 'preco' in transacao and isinstance(transacao['preco'], (int, float)):
                preco = transacao['preco']
                # Se o preço for muito alto, provavelmente está sem casas decimais
                if preco > 10000 and len(str(int(preco))) >= 5:
                    # Converter para formato com decimal (dividindo por 100)
                    transacao['preco'] = preco / 100
        
        # Define campos comuns de taxas
        campos_taxas = {
            'taxa_liquidacao': 'Taxa de Liquidação',
            'taxa_registro': 'Taxa de Registro',
            'taxa_termo': 'Taxa de Termo/Opções',
            'taxa_ana': 'Taxa A.N.A',
            'emolumentos': 'Emolumentos',
            'taxa_operacional': 'Taxa Operacional',
            'execucao': 'Execução',
            'corretagem': 'Corretagem',
            'iss': 'ISS',
            'irrf': 'IRRF Retido na Fonte',
            'outras_taxas': 'Outros'
        }

        # Caso onde não há transações
        if not transacoes:
            # Se não há transações, criar entrada apenas com as taxas
            registro = {
                'Data': data,
                'Número da Nota': numero_nota,
                'Tipo de Transação': 'SEM TRANSAÇÕES',
                'Quantidade': 0,
                'Preço Unitário': 0.0,
                'Ativo': 'N/A',
            }
            
            # Adicionar taxas
            for campo_origem, campo_destino in campos_taxas.items():
                registro[campo_destino] = float(taxas.get(campo_origem, 0.0))
                
            # Adicionar campos que não temos no analisador customizado
            for campo in ['Taxa de Custódia', 'Impostos']:
                if campo not in registro:
                    registro[campo] = 0.0
                    
            dados_por_mes[mes_ano].append(registro)
        # Caso onde há transações
        else:
            # Processar cada transação encontrada
            for transacao in transacoes:
                tipo = transacao.get('tipo', 'N/A')
                ativo = transacao.get('ativo', 'N/A')
                qtd = float(transacao.get('quantidade', 0))
                preco = float(transacao.get('preco', 0))
                tipo_negocio = transacao.get('tipo_negocio', '')
                dc = transacao.get('dc', '')
                
                # Extrair campos específicos para mercado futuro
                ticker = transacao.get('ticker', ativo)
                vencimento = transacao.get('vencimento', '')
                valor_operacao = transacao.get('valor_operacao', 0)
                taxa_operacional = transacao.get('taxa_operacional', 0)
                
                registro = {
                    'Data': data_nota,
                    'Número da Nota': numero_nota,
                    'C/V': tipo,  # C ou V direto
                    'Mercadoria': ticker,
                    'Vencimento': vencimento,
                    'Quantidade': qtd,
                    'Preço / Ajuste': preco,
                    'Tipo Negócio': tipo_negocio,
                    'Valor Operação / D/C': valor_operacao,
                    'D/C': dc,
                    'Taxa Operacional': taxa_operacional,
                    'Ativo Original': ativo  # Mantemos o ativo original como referência
                }
                
                # Adicionar taxas (divididas igualmente entre as transações)
                divisor = len(transacoes)
                campos_taxas = {
                    'taxa_liquidacao': 'Taxa de Liquidação',
                    'taxa_registro': 'Taxa de Registro',
                    'taxa_termo': 'Taxa de Termo/Opções',
                    'taxa_ana': 'Taxa A.N.A',
                    'emolumentos': 'Emolumentos',
                    'taxa_operacional': 'Taxa Operacional',
                    'execucao': 'Execução',
                    'corretagem': 'Corretagem',
                    'iss': 'ISS',
                    'irrf': 'IRRF Retido na Fonte',
                    'outras_taxas': 'Outros'
                }
                
                for campo_origem, campo_destino in campos_taxas.items():
                    registro[campo_destino] = float(taxas.get(campo_origem, 0.0)) / divisor
                    
                # Adicionar campos que não temos no analisador customizado
                for campo in ['Taxa de Custódia', 'Impostos']:
                    if campo not in registro:
                        registro[campo] = 0.0
                        
                dados_por_mes[mes_ano].append(registro)
        
        return True
    except Exception as e:
        logger.log(f"Erro ao processar resultado do analisador customizado: {str(e)}", "erro")
        logger.log(f"Detalhes do erro: {traceback.format_exc()}", "erro")
        return False

# Função principal para processar os PDFs e exportar para Excel
def processar_notas(modo, origem, log_widget, progress_bar, status_var):
    dados_por_mes = defaultdict(list)
    logger = LogHandler(log_widget)
    total_arquivos = 0
    total_notas = 0
    total_transacoes = 0
    arquivo_saida = None
    
    try:
        # Definir arquivos a processar com base no modo (pasta ou arquivos individuais)
        if modo == "pasta":
            arquivos = [os.path.join(origem, f) for f in os.listdir(origem) if f.lower().endswith('.pdf')]
            if not arquivos:
                logger.log("Nenhum PDF encontrado na pasta selecionada.", "alerta")
                return False
            logger.log(f"Encontrados {len(arquivos)} arquivos PDF na pasta para processamento", "info")
            # Gerar nome do arquivo Excel de saída no mesmo local da pasta
            arquivo_saida = gerar_nome_saida_automatico(origem)
        else:  # modo == "arquivos"
            arquivos = origem.split(";")
            if not arquivos or not arquivos[0]:
                logger.log("Nenhum arquivo selecionado.", "alerta")
                return False
            logger.log(f"Processando {len(arquivos)} arquivos selecionados", "info")
            # Gerar nome do arquivo Excel de saída no mesmo local do primeiro PDF
            arquivo_saida = gerar_nome_saida_automatico(arquivos[0])
        
        logger.log(f"O arquivo Excel será salvo como: {arquivo_saida}", "info")
        total_arquivos = len(arquivos)
        
        # Configurar barra de progresso
        progress_bar["maximum"] = total_arquivos
        progress_bar["value"] = 0
        
        # Processar cada arquivo PDF
        for i, caminho_pdf in enumerate(arquivos):
            nome_arquivo = os.path.basename(caminho_pdf)
            status_var.set(f"Processando: {nome_arquivo} ({i+1}/{total_arquivos})")
            logger.log(f"Processando: {nome_arquivo}")
            
            sucesso, arquivo_notas, arquivo_transacoes = processar_arquivo_pdf(caminho_pdf, dados_por_mes, logger)
            if sucesso:
                total_notas += arquivo_notas
                total_transacoes += arquivo_transacoes
            
            # Atualizar progresso
            progress_bar["value"] = i + 1
            progress_bar.update()
        
        if not dados_por_mes:
            logger.log("Nenhuma transação encontrada nos PDFs.", "alerta")
            return False
        
        # Estatísticas finais
        logger.log("Estatísticas do processamento:", "info")
        logger.log(f"  → Arquivos PDF processados: {total_arquivos}")
        logger.log(f"  → Total de notas encontradas: {total_notas}")
        logger.log(f"  → Total de transações: {total_transacoes}")
        logger.log(f"  → Períodos encontrados: {', '.join(sorted(dados_por_mes.keys()))}")
            
        # Exportar para Excel
        status_var.set("Gerando arquivo Excel...")
        logger.log(f"Salvando Excel em: {arquivo_saida}", "info")
        with pd.ExcelWriter(arquivo_saida, engine='openpyxl') as writer:
            for mes, dados in sorted(dados_por_mes.items()):
                # Corrigir os valores de preço quando necessário
                for registro in dados:
                    # Checar campos de preço que podem estar com formato incorreto
                    campos_preco = ['Preço Unitário', 'Preço / Ajuste']
                    for campo in campos_preco:
                        if campo in registro and isinstance(registro[campo], (int, float)):
                            valor = registro[campo]
                            if valor > 10000 and len(str(int(valor))) >= 5:  # Valores altos sem decimal
                                # Converter para formato com decimal (dividindo por 100)
                                registro[campo] = valor / 100
                    
                    # Corrigir o formato do vencimento se necessário
                    if 'Vencimento' in registro and isinstance(registro['Vencimento'], str):
                        vencimento = registro['Vencimento']
                        if vencimento and not vencimento.strip().startswith('0'):
                            # Garantir que a data está no formato DD/MM/AAAA
                            try:
                                if '/' in vencimento:
                                    partes = vencimento.split('/')
                                    if len(partes) == 3 and len(partes[0]) == 2 and len(partes[1]) == 2 and len(partes[2]) == 4:
                                        # Já está no formato correto
                                        pass
                                    elif len(partes) == 3:
                                        # Garantir o formato com zeros à esquerda
                                        registro['Vencimento'] = f"{int(partes[0]):02d}/{int(partes[1]):02d}/{partes[2]}"
                            except Exception:
                                pass  # Ignorar erros de formatação
                    
                df = pd.DataFrame(dados)
                nome_aba = mes.replace('-', '_')  # Substituir '-' por '_' para o nome da aba
                
                # Criar o DataFrame e salvar na planilha
                df.to_excel(writer, sheet_name=nome_aba, index=False)
                
                # Aplicar formatação às colunas numéricas
                workbook = writer.book
                worksheet = writer.sheets[nome_aba]
                
                # Aplicar formatação monetária brasileira (R$ #.##0,00)
                colunas_moeda = ['Preço Unitário', 'Preço / Ajuste', 'Valor Operação / D/C', 'Taxa Operacional',
                                'Taxa de Liquidação', 'Taxa de Registro', 'Taxa de Termo/Opções', 'Taxa A.N.A', 
                                'Emolumentos', 'Execução', 'Taxa de Custódia', 'IRRF Retido na Fonte', 
                                'Impostos', 'Outros', 'Corretagem', 'ISS']
                
                # Definir a ordem das colunas para o formato de futuros (quando disponível)
                ordem_colunas_futuro = ['Data', 'Número da Nota', 'C/V', 'Mercadoria', 'Vencimento', 'Quantidade', 
                                    'Preço / Ajuste', 'Tipo Negócio', 'Valor Operação / D/C', 'D/C', 'Taxa Operacional']
                
                # Reorganizar colunas quando possível
                colunas_existentes = [c for c in ordem_colunas_futuro if c in df.columns]
                outras_colunas = [c for c in df.columns if c not in ordem_colunas_futuro]
                
                # Se temos pelo menos as colunas básicas de futuros, reorganizamos
                colunas_basicas_futuro = ['C/V', 'Mercadoria', 'Vencimento', 'Preço / Ajuste']
                if all(c in df.columns for c in colunas_basicas_futuro):
                    df = df[colunas_existentes + outras_colunas]
                
                # Aplicar formatos para todas as colunas
                for idx, coluna in enumerate(df.columns):
                    # Obter a letra da coluna (A, B, C, etc.)
                    col_letter = chr(65 + idx) if idx < 26 else chr(64 + idx // 26) + chr(65 + idx % 26)
                    
                    # Aplicar formatação monetária
                    if coluna in colunas_moeda:
                        # Formato para moeda brasileira
                        for row in range(2, len(df) + 2):  # +2 porque Excel é 1-indexed e temos o cabeçalho
                            cell = f"{col_letter}{row}"
                            try:
                                # Aplicar formatação monetária brasileira
                                worksheet[cell].number_format = 'R$ #,##0.00'
                            except Exception:
                                pass  # Ignorar erros de formatação
                    
                    # Formatação específica para outras colunas
                    elif coluna == 'C/V':
                        # Deixar centralizado
                        worksheet.column_dimensions[col_letter].width = 6
                    elif coluna == 'Mercadoria':
                        worksheet.column_dimensions[col_letter].width = 12
                    elif coluna == 'Vencimento':
                        worksheet.column_dimensions[col_letter].width = 12
                        # Formato de data
                        for row in range(2, len(df) + 2):
                            cell = f"{col_letter}{row}"
                            try:
                                worksheet[cell].number_format = 'dd/mm/yyyy'
                            except Exception:
                                pass
                    elif coluna == 'Quantidade':
                        worksheet.column_dimensions[col_letter].width = 10
                        # Formato numérico
                        for row in range(2, len(df) + 2):
                            cell = f"{col_letter}{row}"
                            try:
                                worksheet[cell].number_format = '#,##0'
                            except Exception:
                                pass
                    elif coluna == 'Tipo Negócio':
                        worksheet.column_dimensions[col_letter].width = 15
                    elif coluna == 'D/C':
                        worksheet.column_dimensions[col_letter].width = 5
                        # Centralizar
                        for row in range(2, len(df) + 2):
                            cell = f"{col_letter}{row}"
                            try:
                                worksheet[cell].alignment = workbook.styles.Alignment(horizontal='center')
                            except Exception:
                                pass
                
                logger.log(f"  → Planilha '{nome_aba}' criada com {len(df)} transações")
                
        logger.log(f"Arquivo Excel '{os.path.basename(arquivo_saida)}' criado com sucesso.", "sucesso")
        return arquivo_saida
        
    except Exception as e:
        logger.log(f"Erro no processamento: {str(e)}", "erro")
        return False

# Função para gerar nome de saída automático no mesmo local que o PDF original
def gerar_nome_saida_automatico(caminho_origem):
    # Se caminho_origem é um arquivo PDF
    if os.path.isfile(caminho_origem) and caminho_origem.lower().endswith('.pdf'):
        # Usar o mesmo diretório e nome base, apenas alterando a extensão
        diretorio = os.path.dirname(caminho_origem)
        nome_base = os.path.splitext(os.path.basename(caminho_origem))[0]
        return os.path.join(diretorio, f"{nome_base}_exportado.xlsx")
    
    # Se caminho_origem é um diretório
    elif os.path.isdir(caminho_origem):
        # Usar o nome do diretório como base para o nome do arquivo
        nome_diretorio = os.path.basename(caminho_origem)
        if not nome_diretorio:  # Se for raiz do drive, usar algo genérico
            nome_diretorio = "notas_corretagem"
        return os.path.join(caminho_origem, f"{nome_diretorio}_exportado.xlsx")
    
    # Caso padrão (nem arquivo nem diretório válido)
    else:
        # Usar diretório de documentos do usuário
        diretorio_documentos = os.path.expanduser("~\\Documents")
        return os.path.join(diretorio_documentos, "relatorio_notas_corretagem.xlsx")

# Função para selecionar pasta
def selecionar_pasta():
    pasta = filedialog.askdirectory(title="Selecione a pasta dos PDFs das notas de corretagem")
    if pasta:
        entrada_pasta.delete(0, tk.END)
        entrada_pasta.insert(0, pasta)
        status_var.set(f"Pasta selecionada: {pasta}")
        # Mostrar caminho do Excel que será gerado
        arquivo_saida = gerar_nome_saida_automatico(pasta)
        log_handler.log(f"O arquivo Excel será salvo como: {arquivo_saida}", "info")

# Função para selecionar arquivos PDF individuais
def selecionar_arquivos():
    arquivos = filedialog.askopenfilenames(
        filetypes=[("PDF Files", "*.pdf")],
        title="Selecione os PDFs das notas de corretagem"
    )
    if arquivos:
        arquivos_string = ";".join(arquivos)
        entrada_arquivos.delete(0, tk.END)
        entrada_arquivos.insert(0, arquivos_string)
        qtd_arquivos = len(arquivos)
        status_var.set(f"{qtd_arquivos} arquivo{'s' if qtd_arquivos > 1 else ''} selecionado{'s' if qtd_arquivos > 1 else ''}")
        # Mostrar caminho do Excel que será gerado
        if arquivos:
            arquivo_saida = gerar_nome_saida_automatico(arquivos[0])
            log_handler.log(f"O arquivo Excel será salvo como: {arquivo_saida}", "info")

# Função para abrir o diretório do arquivo gerado
def abrir_diretorio_resultado(arquivo):
    try:
        diretorio = os.path.dirname(arquivo)
        os.startfile(diretorio)
    except Exception as e:
        messagebox.showerror("Erro", f"Não foi possível abrir o diretório: {str(e)}")

# Função para abrir o arquivo Excel gerado
def abrir_arquivo_excel(arquivo):
    try:
        os.startfile(arquivo)
    except Exception as e:
        messagebox.showerror("Erro", f"Não foi possível abrir o arquivo: {str(e)}")

# Função para obter o modo e origem atuais
def obter_modo_e_origem():
    modo_aba = notebook.tab(notebook.select(), "text").lower()
    
    if modo_aba == "pasta":
        pasta = entrada_pasta.get()
        return "pasta", pasta
    else:  # "arquivos individuais"
        arquivos = entrada_arquivos.get()
        return "arquivos", arquivos

# Função para iniciar processamento em thread separada
def iniciar_processamento_thread():
    # Verificar seleção
    modo, origem = obter_modo_e_origem()
    
    # Validar entradas
    if modo == "pasta":
        if not origem:
            messagebox.showerror("Erro", "Selecione uma pasta contendo PDFs de notas de corretagem.")
            return
        if not os.path.isdir(origem):
            messagebox.showerror("Erro", "A pasta selecionada não existe.")
            return
    else:  # modo == "arquivos"
        if not origem:
            messagebox.showerror("Erro", "Selecione pelo menos um arquivo PDF de nota de corretagem.")
            return
        for arquivo in origem.split(";"):
            if not os.path.isfile(arquivo) or not arquivo.lower().endswith('.pdf'):
                messagebox.showerror("Erro", f"Arquivo inválido: {arquivo}")
                return
    
    # Reiniciar log e progresso
    log_text.config(state=tk.NORMAL)
    log_text.delete(1.0, tk.END)
    log_text.config(state=tk.DISABLED)
    progress_bar["value"] = 0
    
    # Desativar controles durante processamento
    toggle_controles(False)
    
    # Criar e iniciar thread
    thread = threading.Thread(target=lambda: processar_thread(modo, origem))
    thread.daemon = True
    thread.start()

# Função para executar processamento em thread separada
def processar_thread(modo, origem):
    try:
        # Atualizar interface
        status_var.set("Iniciando processamento...")
        
        # Desativar controles
        toggle_controles(False)
        
        # Executar processamento - retorna o caminho do arquivo criado ou False
        resultado = processar_notas(modo, origem, log_text, progress_bar, status_var)
        
        # Atualizar interface após conclusão
        if isinstance(resultado, str) and os.path.exists(resultado):
            status_var.set("Processamento concluído com sucesso")
            root.after(0, lambda: mostrar_resultado_sucesso(resultado))
        else:
            status_var.set("Processamento concluído com erros")
            root.after(0, lambda: messagebox.showerror("Erro", "Ocorreram erros durante o processamento. Verifique o log para mais detalhes."))
    
    except Exception as e:
        status_var.set("Erro durante processamento")
        erro_msg = str(e)
        root.after(0, lambda msg=erro_msg: messagebox.showerror("Erro inesperado", msg))
    
    finally:
        # Reativar controles
        root.after(0, lambda: toggle_controles(True))

def toggle_controles(ativar):
    estado = tk.NORMAL if ativar else tk.DISABLED
    botao_processar.config(state=estado)
    botao_pasta.config(state=estado)
    botao_arquivos.config(state=estado)

def mostrar_resultado_sucesso(arquivo):
    # Mostrar popup simplificado de sucesso com opções para abrir o arquivo
    popup = tk.Toplevel(root)
    popup.title("Concluído")
    popup.geometry("450x220")
    popup.resizable(False, False)
    popup.transient(root)
    popup.grab_set()
    
    # Aplicar tema escuro com borda suave
    popup.configure(bg=CORES["bg_escuro"])
    
    # Frame de conteúdo principal
    content_frame = ttk.Frame(popup, style="Card.TFrame")
    content_frame.pack(fill="both", expand=True, padx=15, pady=15)
    
    # Ícone de sucesso e mensagem principal em uma única linha
    header_frame = ttk.Frame(content_frame, style="Card.TFrame")
    header_frame.pack(fill="x", padx=5, pady=5)
    
    # Ícone de sucesso
    ttk.Label(
        header_frame, 
        text="✅", # Ícone de sucesso mais simples
        font=("Segoe UI", 24),
        foreground=CORES["sucesso"],
        background=CORES["bg_medio"]
    ).pack(side=tk.LEFT, padx=(5, 10))
    
    # Texto de sucesso
    ttk.Label(
        header_frame, 
        text="Processamento Concluído", 
        font=("Segoe UI", 14, "bold"),
        foreground=CORES["texto"],
        background=CORES["bg_medio"]
    ).pack(side=tk.LEFT)
    
    # Caminho do arquivo
    info_frame = ttk.Frame(content_frame, style="Card.TFrame")
    info_frame.pack(fill="x", padx=5, pady=10)
    
    ttk.Label(
        info_frame,
        text="Arquivo gerado:",
        font=("Segoe UI", 10),
        foreground=CORES["texto"],
        background=CORES["bg_medio"]
    ).pack(anchor="w")
    
    # Nome do arquivo com caminho completo
    ttk.Label(
        info_frame,
        text=arquivo,
        font=("Consolas", 9),
        foreground=CORES["destaque"],
        background=CORES["bg_medio"],
        wraplength=400
    ).pack(anchor="w", padx=(10, 0), pady=2)
    
    # Frame para botões - versão simplificada
    botoes_frame = ttk.Frame(content_frame, style="Card.TFrame")
    botoes_frame.pack(pady=10, fill="x")
    
    # Botão principal em destaque
    ttk.Button(
        botoes_frame, 
        text="Abrir Arquivo", 
        style="Accent.TButton",
        command=lambda: [popup.destroy(), abrir_arquivo_excel(arquivo)]
    ).pack(side=tk.LEFT, padx=5, expand=True, fill="x")
    
    # Botões secundarios
    ttk.Button(
        botoes_frame, 
        text="Abrir Pasta", 
        command=lambda: [popup.destroy(), abrir_diretorio_resultado(arquivo)]
    ).pack(side=tk.LEFT, padx=5, expand=True, fill="x")
    
    ttk.Button(
        botoes_frame, 
        text="Fechar", 
        command=popup.destroy
    ).pack(side=tk.LEFT, padx=5, expand=True, fill="x")

# Criar janela principal
root = tk.Tk()
root.title("Correpy Plus")
root.geometry("1000x650")
root.minsize(800, 600)

# Adicionar ícone para a janela (opcional)
try:
    root.iconbitmap("icon.ico")  # Se tiver um ícone disponivel
except Exception:
    pass  # Continuar sem ícone se não encontrar

# Configurar tema escuro moderno
root.configure(bg=CORES["bg_escuro"])

# Configurar estilo personalizado
style = themes.ThemedStyle(root)
style.set_theme("equilux")  # Tema base escuro

# Personalizar estilos dos widgets para uma aparência mais moderna e elegante
style.configure("TFrame", background=CORES["bg_escuro"])
style.configure("Card.TFrame", background=CORES["bg_medio"], relief="flat", borderwidth=0, padding=15)
style.configure("TLabel", background=CORES["bg_escuro"], foreground=CORES["texto"], font=("Segoe UI", 10))
style.configure("Header.TLabel", font=("Segoe UI", 13, "bold"), foreground=CORES["destaque"])
style.configure("Title.TLabel", font=("Segoe UI", 20, "bold"), foreground=CORES["destaque"])
style.configure("Subtitle.TLabel", font=("Segoe UI", 11), foreground=CORES["alerta"])
style.configure("Status.TLabel", font=("Segoe UI", 9), foreground=CORES["texto"])

# Estilo para botões modernos
style.configure("TButton", font=("Segoe UI", 10), padding=5)
style.configure("Accent.TButton", background=CORES["destaque"], foreground=CORES["bg_escuro"], font=("Segoe UI", 10, "bold"))
style.map("Accent.TButton",
         background=[('active', CORES["sucesso"]), ('pressed', CORES["destaque"])],
         foreground=[('active', CORES["bg_escuro"]), ('pressed', CORES["bg_escuro"])])

# Estilo para notebook (abas)
style.configure("TNotebook", background=CORES["bg_escuro"], borderwidth=0)
style.configure("TNotebook.Tab", background=CORES["bg_medio"], foreground=CORES["texto"], padding=[15, 5], font=("Segoe UI", 10))
style.map("TNotebook.Tab",
         background=[('selected', CORES["destaque"])],
         foreground=[('selected', CORES["bg_escuro"])])

# Estilo para entradas
root.option_add("*TEntry*font", ("Segoe UI", 10))
style.configure("TEntry", fieldbackground=CORES["bg_claro"], foreground=CORES["texto"], borderwidth=1, padding=5)
style.configure("Status.TLabel", font=("Segoe UI", 9))

# Estilo de botões
style.configure("TButton", font=("Segoe UI", 10))
style.configure("Accent.TButton", background=CORES["destaque"])

# Frame para título e cabeçalho com visual moderno
frame_header = ttk.Frame(root, style="Card.TFrame")
frame_header.pack(fill=tk.X, padx=20, pady=(15, 5))

# Container para logo e título com efeito gradiente
logo_container = ttk.Frame(frame_header, style="Card.TFrame")
logo_container.pack(fill=tk.X, pady=(0, 5))

# Logo e título com ícone moderno
titulo_label = ttk.Label(
    logo_container, 
    text="📊 Correpy Plus", 
    style="Title.TLabel",
    font=("Segoe UI", 22, "bold")
)
titulo_label.pack(side=tk.LEFT, padx=10)

# Versão
versao_label = ttk.Label(
    logo_container,
    text="v1.0",
    style="Subtitle.TLabel"
)
versao_label.pack(side=tk.LEFT, padx=(0, 10))

# Data atual com formato mais elegante e ícone de calendário
data_atual = datetime.now().strftime("%d de %B de %Y")
data_label = ttk.Label(
    logo_container, 
    text=f"📅 {data_atual}", 
    style="Header.TLabel"
)
data_label.pack(side=tk.RIGHT, padx=10)

# Subtítulo descritivo com visual aprimorado
subtitulo_container = ttk.Frame(frame_header, style="Card.TFrame")
subtitulo_container.pack(fill=tk.X, pady=5)

subtitulo_label = ttk.Label(
    subtitulo_container,
    text="Extrator avançado de dados de notas de corretagem para Excel",
    style="Subtitle.TLabel",
    font=("Segoe UI", 11, "italic")
)
subtitulo_label.pack(side=tk.LEFT, padx=10)

# Adiciona informação de recursos especiais
ttk.Label(
    subtitulo_container,
    text="✨ Suporte completo para mercado futuro",
    style="Status.TLabel",
    foreground=CORES["destaque"]
).pack(side=tk.RIGHT, padx=10)

# Layout principal reorganizado para design moderno
main_container = ttk.Frame(root, style="TFrame")
main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=15)

# Painel de comandos lateral (esquerda)
cmd_panel = ttk.Frame(main_container, style="Card.TFrame")
cmd_panel.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10), pady=0, expand=False)

# Título do painel de comandos
cmd_title = ttk.Label(cmd_panel, text="📂 Ações", style="Header.TLabel")
cmd_title.pack(side=tk.TOP, padx=10, pady=(0, 10), anchor="w")

# Botões principais com ícones
btn_frame = ttk.Frame(cmd_panel, style="Card.TFrame")
btn_frame.pack(side=tk.TOP, fill=tk.X, padx=5, pady=5)

botao_processar = ttk.Button(
    btn_frame, 
    text="🚀 Processar Notas", 
    command=iniciar_processamento_thread,
    style="Accent.TButton",
    width=20
)
botao_processar.pack(padx=5, pady=5, fill=tk.X)

botao_limpar = ttk.Button(
    btn_frame, 
    text="🗑️ Limpar Campos", 
    command=lambda: [entrada_pasta.delete(0, tk.END), entrada_arquivos.delete(0, tk.END)],
    width=20
)
botao_limpar.pack(padx=5, pady=5, fill=tk.X)

botao_sair = ttk.Button(btn_frame, text="🚪 Sair", command=root.quit, width=20)
botao_sair.pack(padx=5, pady=5, fill=tk.X)

# Frame principal - corpo da aplicação
frame_principal = ttk.Frame(main_container, style="TFrame")
frame_principal.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0, 10))

# Painel esquerdo - Controles
frame_controles = ttk.Frame(frame_principal, style="Card.TFrame")
frame_controles.pack(side=tk.LEFT, fill=tk.BOTH, padx=(0, 10), expand=True)

# Título do painel de controles
ttk.Label(
    frame_controles, 
    text="Configurações", 
    style="Header.TLabel",
    background=CORES["bg_medio"]
).pack(anchor="w", padx=15, pady=15)

# Conteúdo do painel de controles - Frame interno com padding
frame_form = ttk.Frame(frame_controles, style="TFrame")
frame_form.pack(fill=tk.BOTH, expand=True, padx=15, pady=(0, 15))

# Criar notebook (sistema de abas)
notebook = ttk.Notebook(frame_form)
notebook.pack(fill=tk.BOTH, expand=True, pady=(0, 15))

# Aba 1: Processar pasta
frame_aba_pasta = ttk.Frame(notebook, style="TFrame")
notebook.add(frame_aba_pasta, text="Pasta", padding=10)

# Aba 2: Processar arquivos individuais
frame_aba_arquivos = ttk.Frame(notebook, style="TFrame")
notebook.add(frame_aba_arquivos, text="Arquivos Individuais", padding=10)

# Conteúdo da Aba 1: Pasta
ttk.Label(frame_aba_pasta, text="Selecione a pasta contendo as notas de corretagem em PDF:").pack(anchor="w", pady=(5, 5))

frame_pasta = ttk.Frame(frame_aba_pasta)
frame_pasta.pack(fill=tk.X, pady=(0, 15))

entrada_pasta = tk.Entry(frame_pasta, bg=CORES["bg_claro"], fg=CORES["texto"], insertbackground=CORES["texto"])
entrada_pasta.pack(side=tk.LEFT, fill=tk.X, expand=True)

botao_pasta = ttk.Button(frame_pasta, text="Escolher Pasta", command=selecionar_pasta)
botao_pasta.pack(side=tk.RIGHT, padx=(10, 0))

# Mensagem informativa sobre exportação automática
informacao_exportacao = ttk.Label(
    frame_aba_pasta, 
    text="ℹ️ O arquivo Excel será exportado automaticamente no mesmo local do PDF original", 
    foreground=CORES["destaque"],
    wraplength=400,
    justify="left",
    font=("Segoe UI", 9, "italic"),
)
informacao_exportacao.pack(anchor="w", pady=(15, 5))

ttk.Label(frame_aba_pasta, text="📄 Processa todos os PDFs dentro de uma pasta", foreground=CORES["alerta"]).pack(anchor="w", pady=(15, 0))

# Conteúdo da Aba 2: Arquivos Individuais
ttk.Label(frame_aba_arquivos, text="Selecione os PDFs das notas de corretagem:").pack(anchor="w", pady=(5, 5))

frame_selecao_arquivos = ttk.Frame(frame_aba_arquivos)
frame_selecao_arquivos.pack(fill=tk.X, pady=(0, 15))

entrada_arquivos = tk.Entry(frame_selecao_arquivos, bg=CORES["bg_claro"], fg=CORES["texto"], insertbackground=CORES["texto"])
entrada_arquivos.pack(side=tk.LEFT, fill=tk.X, expand=True)

botao_arquivos = ttk.Button(frame_selecao_arquivos, text="Escolher Arquivos", command=selecionar_arquivos)
botao_arquivos.pack(side=tk.RIGHT, padx=(10, 0))

# Mensagem informativa sobre exportação automática
informacao_exportacao2 = ttk.Label(
    frame_aba_arquivos, 
    text="ℹ️ O arquivo Excel será exportado automaticamente no mesmo local do primeiro PDF selecionado", 
    foreground=CORES["destaque"],
    wraplength=400,
    justify="left",
    font=("Segoe UI", 9, "italic")
)
informacao_exportacao2.pack(anchor="w", pady=(15, 5))

ttk.Label(frame_aba_arquivos, text="📋 Selecione apenas os arquivos específicos para processar", foreground=CORES["alerta"]).pack(anchor="w", pady=(15, 0))

# Separador visual
separador = ttk.Separator(frame_form, orient="horizontal")
separador.pack(fill=tk.X, pady=15)

# Layout - Botões de ação
botoes_frame = ttk.Frame(frame_form)
botoes_frame.pack(fill=tk.X, pady=(5, 0))

botao_processar = ttk.Button(
    botoes_frame, 
    text="📂 Processar Notas", 
    command=iniciar_processamento_thread,
    style="Accent.TButton",
    width=20
)
botao_processar.pack(padx=5, pady=5, fill=tk.X)

botao_sair = ttk.Button(botoes_frame, text="✖ Sair", command=root.quit, width=20)
botao_sair.pack(padx=5, pady=5, fill=tk.X)

# Painel direito - Log e progresso
frame_log = ttk.Frame(frame_principal, style="Card.TFrame")
frame_log.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

# Título do painel de log
ttk.Label(
    frame_log, 
    text="Log de Processamento", 
    style="Header.TLabel",
    background=CORES["bg_medio"]
).pack(anchor="w", padx=15, pady=15)

# Área de log
frame_log_interno = ttk.Frame(frame_log, style="TFrame")
frame_log_interno.pack(fill=tk.BOTH, expand=True, padx=15, pady=(0, 15))

# Barra de progresso
ttk.Label(frame_log_interno, text="Progresso:").pack(anchor="w", pady=(0, 5))
progress_bar = ttk.Progressbar(frame_log_interno, orient="horizontal", length=300, mode="determinate")
progress_bar.pack(fill=tk.X, pady=(0, 10))

# Status
status_var = tk.StringVar(value="Pronto para iniciar")
status_label = ttk.Label(frame_log_interno, textvariable=status_var, style="Status.TLabel")
status_label.pack(anchor="w", pady=(0, 10))

# Área de texto de log
log_text = scrolledtext.ScrolledText(
    frame_log_interno, 
    bg=CORES["bg_escuro"], 
    fg=CORES["texto"], 
    height=20,
    insertbackground=CORES["texto"],
    font=("Consolas", 9)
)
log_text.pack(fill=tk.BOTH, expand=True)
log_text.tag_configure("erro", foreground=CORES["erro"])
log_text.tag_configure("sucesso", foreground=CORES["sucesso"])
log_text.tag_configure("alerta", foreground=CORES["alerta"])
log_text.tag_configure("info", foreground=CORES["destaque"])
log_text.tag_configure("normal", foreground=CORES["texto"])

# Inserir mensagem inicial de boas-vindas
log_handler = LogHandler(log_text)
log_handler.log("Bem-vindo ao Correpy Plus!", "info")
log_handler.log("Selecione a pasta contendo as notas de corretagem em PDF e o arquivo Excel de saída.")
log_handler.log("Clique em 'Processar Notas' para iniciar a conversão.")

# Barra de status na parte inferior
frame_status = ttk.Frame(root, style="Card.TFrame")
frame_status.pack(fill=tk.X, padx=20, pady=10)

ttk.Label(frame_status, text="📊 Correpy Plus v1.0 | Desenvolvido com ❤️", style="Status.TLabel", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=10, pady=5)
ttk.Label(frame_status, text="© 2025 | github.com/thiagosalvatore/correpy", style="Status.TLabel").pack(side=tk.RIGHT, padx=10, pady=5)

# Iniciar loop de eventos
root.mainloop()
