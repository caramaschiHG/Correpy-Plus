from pathlib import Path


def test_main_nao_tem_campos_gui_inexistentes():
    source = Path('main.py').read_text(encoding='utf-8')
    assert 'entrada_arquivo_pasta' not in source
    assert 'entrada_arquivo_individual' not in source


def test_extrator_notas_define_busca_secoes():
    source = Path('extrator_notas.py').read_text(encoding='utf-8')
    assert 'def buscar_secoes_transacoes(' in source
