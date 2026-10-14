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
    assert "JSZip v3.10.1" in html
    assert "/*MODELO_PARECER*/null" not in html  # modelo do parecer embutido
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


def test_versao_online_usa_cdn_e_embute_o_modelo(tmp_path):
    html = gerar_aplicativo.gerar(tmp_path / "pagina.html", online=True).read_text(encoding="utf-8")
    assert not html.startswith("<!doctype")
    assert "cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0" in html
    assert "cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1" in html
    assert "/*MODELO_PARECER*/null" not in html


def test_modelo_do_parecer_tem_todos_os_marcadores():
    import re
    import zipfile

    with zipfile.ZipFile(RAIZ / "web" / "modelos" / "parecer_femta.docx") as z:
        xml = z.read("word/document.xml").decode("utf-8")
        assert not any(n.startswith("word/media/image1") for n in z.namelist())
    marcadores = set(re.findall(r"\{\{([A-Z_0-9]+)\}\}", xml))
    assert marcadores == {
        "NUMERO", "MES_ANO", "ANO", "LEI", "VALOR_LOA", "DATA",
        "QUADRO1", "QUADRO2", "QUADRO3", "QUADRO4",
        "TXT_Q1", "TXT_Q2", "TXT_DOTACAO", "TXT_RESULTADO_ORC", "TXT_RESUMO_BF",
        "TXT_RFP", "TXT_AVALIACAO", "TXT_CONCLUSAO", "TXT_INFORMACAO",
    }
