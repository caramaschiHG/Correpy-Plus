"""
Script independente para extrair contratos futuros diretamente de PDFs de notas de corretagem
"""

import re
import pdfplumber

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
    contratos_detectados = set()  # Set para controlar quais contratos já foram detectados
    
    # Definir o mapeamento de códigos de meses
    meses_vencimento = {
        'F': 'Janeiro',
        'G': 'Fevereiro',
        'H': 'Março',
        'J': 'Abril',
        'K': 'Maio',
        'M': 'Junho',
        'N': 'Julho',
        'Q': 'Agosto',
        'U': 'Setembro',
        'V': 'Outubro',
        'X': 'Novembro',
        'Z': 'Dezembro'
    }
    
    # Lista de ativos futuros a detectar (ampliada para incluir mais contratos)
    ativos_futuros = ["WIN", "WDO", "DOL", "IND", "BGI", "CCM", "ICF", "DI1", "DAP", 
                      "SJC", "ISP", "EUR", "FRC", "BOI", "B3", "DI", "DDI"]
    for linha in texto.split('\n'):
        # Verificar padrões específicos para contratos futuros
        # Padrão: C|V WDO F25 02/01/2025 1 6.088,0000 DAY TRADE
        match = re.search(r'([CV])\s+([A-Z0-9]{2,5})\s+([A-Z]\d{1,2})\s+(?:\d{2}/\d{2}/\d{4})?\s*(\d+)[\s\.,]*([\.\d\,]+)', linha, re.IGNORECASE)
        if match:
            grupos = match.groups()
            tipo = grupos[0].upper()
            ativo_base = grupos[1].upper()
            vencimento = grupos[2].upper()
            quantidade = parse_valor(grupos[3])
            preco = parse_valor(grupos[4])
            valor_total = quantidade * preco
            
            # Adicionar informação do mês de vencimento
            mes_vencimento = ""
            if len(vencimento) >= 1 and vencimento[0] in meses_vencimento:
                mes_vencimento = meses_vencimento[vencimento[0]]
            
            # Criar chave única para verificar duplicatas
            chave_contrato = f"{tipo}-{ativo_base}-{vencimento}-{quantidade}-{preco}"
            
            # Verificar se já foi processado
            if chave_contrato in contratos_detectados:
                continue
                
            # Adicionar à lista de contratos detectados
            contratos_detectados.add(chave_contrato)
            
            # Criar transação
            transacao = {
                "tipo": tipo,
                "ativo": f"{ativo_base} {vencimento}",
                "quantidade": quantidade,
                "preco": preco,
                "valor_total": valor_total,
                "vencimento": vencimento,
                "mes_vencimento": mes_vencimento
            }
            transacoes.append(transacao)
            continue
        
        # Padrão: C WDOK23 10 5.278,50
        match = re.search(r'([CV])\s+([A-Z]+[A-Z0-9]\d{2})\s+(\d+)\s+([\d\.\,]+)', linha, re.IGNORECASE)
        if match:
            grupos = match.groups()
            tipo = grupos[0].upper()
            
            # Separar o código do ativo e o vencimento
            codigo_completo = grupos[1].upper()
            ativo_base = ""
            vencimento = ""
            
            # Extrair o ativo base e o código de vencimento
            for ativo in ativos_futuros:
                if ativo in codigo_completo:
                    ativo_base = ativo
                    vencimento = codigo_completo.replace(ativo, "")
                    if vencimento and vencimento[0] in meses_vencimento:
                        mes_vencimento = meses_vencimento[vencimento[0]]
                    else:
                        mes_vencimento = ""
                    break
            
            # Se não identificou o ativo, usa o código completo
            if not ativo_base:
                ativo_base = codigo_completo
                vencimento = ""
                mes_vencimento = ""
            
            quantidade = parse_valor(grupos[2])
            preco = parse_valor(grupos[3])
            valor_total = quantidade * preco
            
            # Criar chave única para verificar duplicatas
            chave_contrato = f"{tipo}-{ativo_base}-{vencimento}-{quantidade}-{preco}"
            
            # Verificar se já foi processado
            if chave_contrato in contratos_detectados:
                continue
                
            # Adicionar à lista de contratos detectados
            contratos_detectados.add(chave_contrato)
            
            # Criar transação
            transacao = {
                "tipo": tipo,
                "ativo": f"{ativo_base} {vencimento}".strip(),
                "quantidade": quantidade,
                "preco": preco,
                "valor_total": valor_total,
                "vencimento": vencimento,
                "mes_vencimento": mes_vencimento
            }
            transacoes.append(transacao)
            continue
        
        # Padrão genérico: buscar por ativos futuros específicos
        for ativo in ativos_futuros:
            if ativo in linha.upper():
                # Tentar extrair operação, quantidade e preço
                match = re.search(r'([CV])\s+(?:.*?\s+)?(\d+)\s+(?:.*?)([\d\.\,]+)', linha, re.IGNORECASE)
                if match:
                    grupos = match.groups()
                    tipo = grupos[0].upper()
                    quantidade = parse_valor(grupos[1])
                    preco = parse_valor(grupos[2])
                    valor_total = quantidade * preco
                    
                    # Extrair vencimento (código como F25, G25, etc)
                    vencimento = ""
                    mes_vencimento = ""
                    # Procurar por padrão de vencimento: letra + 1-2 números
                    for termo in linha.split():
                        codigo_venc = re.match(r'([A-Z])(\d{1,2})', termo, re.IGNORECASE)
                        if codigo_venc:
                            letra_mes = codigo_venc.group(1).upper()
                            if letra_mes in meses_vencimento:
                                vencimento = termo.upper()
                                mes_vencimento = meses_vencimento[letra_mes]
                                break
                    
                    # Criar nome do ativo formatado
                    ativo_formatado = f"{ativo} {vencimento}" if vencimento else ativo
                    
                    # Criar chave única para verificar duplicatas
                    chave_contrato = f"{tipo}-{ativo_formatado}-{quantidade}-{preco}"
                    
                    # Verificar se já foi processado
                    if chave_contrato in contratos_detectados:
                        continue
                    
                    # Adicionar à lista de contratos detectados
                    contratos_detectados.add(chave_contrato)
                    
                    # Criar transação
                    transacao = {
                        "tipo": tipo,
                        "ativo": ativo_formatado,
                        "quantidade": quantidade,
                        "preco": preco,
                        "valor_total": valor_total,
                        "vencimento": vencimento,
                        "mes_vencimento": mes_vencimento
                    }
                    
                    # Adicionar à lista de transações
                    transacoes.append(transacao)
                    
                    break
    
    # Padrão adicional para tabelas estruturadas
    # Procurar por linhas que contenham "C/V" (compra/venda) e algum dos ativos
    linhas_tabela = [linha for linha in texto.split('\n') if any(ativo in linha.upper() for ativo in ativos_futuros)]
    for linha in linhas_tabela:
        # Procurar por padrões do tipo "C/V Mercadoria"
        match = re.search(r'([CV])/[VC]\s+([A-Z0-9]+)\s+([A-Z]\d{1,2})\s+(\d+)\s+([\.\d\,]+)', linha, re.IGNORECASE)
        if match:
            grupos = match.groups()
            tipo = grupos[0].upper()
            ativo_base = grupos[1].upper()
            vencimento = grupos[2].upper()
            quantidade = parse_valor(grupos[3])
            preco = parse_valor(grupos[4])
            valor_total = quantidade * preco
            
            # Adicionar informação do mês de vencimento
            mes_vencimento = ""
            if len(vencimento) >= 1 and vencimento[0] in meses_vencimento:
                mes_vencimento = meses_vencimento[vencimento[0]]
            
            # Criar chave única para verificar duplicatas
            chave_contrato = f"{tipo}-{ativo_base}-{vencimento}-{quantidade}-{preco}"
            
            # Verificar se já foi processado
            if chave_contrato in contratos_detectados:
                continue
            
            # Adicionar à lista de contratos detectados
            contratos_detectados.add(chave_contrato)
            
            # Criar transação
            transacao = {
                "tipo": tipo,
                "ativo": f"{ativo_base} {vencimento}",
                "quantidade": quantidade,
                "preco": preco,
                "valor_total": valor_total,
                "vencimento": vencimento,
                "mes_vencimento": mes_vencimento
            }
            transacoes.append(transacao)
    
    # Imprimir diagnóstico quando executado diretamente
    if __name__ == "__main__" and len(transacoes) > 0:
        print("\nDetalhes dos contratos futuros detectados:")
        print(f"{'Tipo':<5} {'Ativo':<15} {'Vencimento':<10} {'Mês':<10} {'Qtd':<8} {'Preço':<12} {'Total':<12}")
        print("-" * 75)
        for t in transacoes:
            print(f"{t['tipo']:<5} {t['ativo']:<15} {t['vencimento']:<10} {t['mes_vencimento']:<10} "
                  f"{t['quantidade']:<8} {t['preco']:<12,.2f} {t['valor_total']:<12,.2f}")
    
    return transacoes

def main(caminho_pdf):
    texto = extrair_texto_pdf(caminho_pdf)
    transacoes = extrair_contratos_futuros(texto)
    
    print(f"\nEncontradas {len(transacoes)} transações de contratos futuros:")
    for i, t in enumerate(transacoes, 1):
        print(f"  {i}. {t['tipo']} {t['ativo']} - {t['quantidade']} x {t['preco']} = {t['valor_total']:.2f}")
    
    return transacoes

def exportar_para_excel(transacoes, caminho_saida=None):
    """Exporta as transações para um arquivo Excel"""
    try:
        import pandas as pd
        from datetime import datetime
        
        # Se caminho de saída não for fornecido, cria um nome padrão
        if not caminho_saida:
            data_hora = datetime.now().strftime("%Y%m%d_%H%M%S")
            caminho_saida = f"contratos_futuros_{data_hora}.xlsx"
        
        # Converter transacoes para DataFrame
        df = pd.DataFrame(transacoes)
        
        # Adicionar data atual se não existir
        if 'data' not in df.columns:
            df['data'] = datetime.now().strftime("%Y-%m-%d")
        
        # Organizar por tipo de contrato e operação
        if len(df) > 0:
            df = df.sort_values(by=['ativo', 'tipo'])
        
        # Salvar para Excel
        with pd.ExcelWriter(caminho_saida) as writer:
            mes_ano = datetime.now().strftime("%Y_%m")
            df.to_excel(writer, sheet_name=mes_ano, index=False)
        
        print(f"\nArquivo Excel criado com sucesso: {caminho_saida}")
        return caminho_saida
    except Exception as e:
        print(f"Erro ao exportar para Excel: {e}")
        return None

if __name__ == "__main__":
    import sys
    if len(sys.argv) > 1:
        caminho_pdf = sys.argv[1]
        transacoes = main(caminho_pdf)
        
        # Perguntar se deseja exportar para Excel
        resposta = input("\nDeseja exportar as transações para Excel? (s/n): ")
        if resposta.lower() in ['s', 'sim', 'y', 'yes']:
            nome_arquivo = input("Digite o nome do arquivo Excel (ou pressione Enter para usar o padrão): ")
            if nome_arquivo.strip():
                if not nome_arquivo.endswith(".xlsx"):
                    nome_arquivo += ".xlsx"
                exportar_para_excel(transacoes, nome_arquivo)
            else:
                exportar_para_excel(transacoes)
    else:
        print("Por favor, forneça o caminho do arquivo PDF como argumento.")
