import pathlib

from extrair_futuros_direto import extrair_contratos_futuros, parse_valor


def test_parse_valor_brasileiro():
    assert parse_valor("6.088,50") == 6088.50
    assert parse_valor("R$ 1.234,00") == 1234.0
    assert parse_valor("") == 0


def test_extrai_padrao_com_vencimento_e_mes():
    texto = "C WDO F25 02/01/2025 1 6.088,0000 DAY TRADE"
    transacoes = extrair_contratos_futuros(texto)

    assert len(transacoes) == 1
    t = transacoes[0]
    assert t["tipo"] == "C"
    assert t["ativo"] == "WDO F25"
    assert t["quantidade"] == 1
    assert t["preco"] == 6088.0
    assert t["mes_vencimento"] == "Janeiro"


def test_remove_duplicatas_de_mesma_linha():
    texto = "\n".join([
        "V WIN J25 10 126.500,00",
        "V WIN J25 10 126.500,00",
    ])
    transacoes = extrair_contratos_futuros(texto)
    assert len(transacoes) == 1


def test_extrai_formato_aglutinado():
    texto = "C WDOK23 10 5.278,50"
    transacoes = extrair_contratos_futuros(texto)

    assert len(transacoes) == 1
    t = transacoes[0]
    assert t["ativo"] == "WDO K23"
    assert t["mes_vencimento"] == "Maio"


def test_regressao_main_sem_campos_inexistentes_no_limpar():
    source = pathlib.Path("main.py").read_text(encoding="utf-8")
    assert "entrada_arquivo_pasta" not in source
    assert "entrada_arquivo_individual" not in source
