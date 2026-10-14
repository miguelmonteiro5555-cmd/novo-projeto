import importlib.util
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
spec = importlib.util.spec_from_file_location("gerar_aplicativo", RAIZ / "web" / "gerar_aplicativo.py")
gerar_aplicativo = importlib.util.module_from_spec(spec)
spec.loader.exec_module(gerar_aplicativo)


def test_aplicativo_funciona_sem_internet(tmp_path):
    saida = gerar_aplicativo.gerar(tmp_path / "app.html")
    html = saida.read_text(encoding="utf-8")
    assert html.startswith("<!doctype html>")
    assert html.rstrip().endswith("</html>")
    assert "cdnjs.cloudflare.com" not in html
    assert "fonts.googleapis.com" not in html
    assert "/*! ExcelJS" in html
    assert "/*DADOS*/null/*FIM*/" in html  # sem planilha, abre vazio


def test_aplicativo_com_dados_da_planilha(tmp_path):
    from balancete.base_calculo import gerar_base, modelo_vazio
    from decimal import Decimal

    meses = modelo_vazio()
    meses["jan"]["receita_realizada"] = Decimal("1234.56")
    planilha = gerar_base(meses, tmp_path / "base.xlsx", 2031)
    html = gerar_aplicativo.gerar(tmp_path / "app.html", str(planilha)).read_text(encoding="utf-8")
    assert 'DADOS_INICIAIS = {"jan":{"receita_realizada":123456}' in html
    assert 'value="2031"' in html
