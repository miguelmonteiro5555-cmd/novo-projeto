import re
from io import BytesIO
from pathlib import Path

import pytest
from openpyxl import load_workbook

from balancete.app import criar_app

EXEMPLOS = Path(__file__).resolve().parent.parent / "exemplos"


@pytest.fixture
def cliente():
    return criar_app().test_client()


def _arquivo(nome: str):
    return (BytesIO((EXEMPLOS / nome).read_bytes()), nome)


def _id_download(html: str) -> str:
    return re.search(r"/baixar/([0-9a-f]+)\.xlsx", html).group(1)


def test_pagina_inicial(cliente):
    resposta = cliente.get("/")
    assert resposta.status_code == 200
    assert "Gerar balancete" in resposta.text


def test_gerar_com_upload_e_baixar(cliente):
    resposta = cliente.post(
        "/gerar",
        data={
            "mes": "2026-09",
            "empresa": "ACME",
            "plano": _arquivo("plano_de_contas.csv"),
            "lancamentos": _arquivo("lancamentos.csv"),
            "saldos": _arquivo("saldos_iniciais.csv"),
        },
        content_type="multipart/form-data",
    )
    assert resposta.status_code == 200
    html = resposta.text
    assert "ACME — Balancete de Verificação — Setembro/2026" in html
    assert "Balancete fechado" in html
    assert "77.385,00" in html

    identificador = _id_download(html)
    excel = cliente.get(f"/baixar/{identificador}.xlsx")
    assert excel.status_code == 200
    aba = load_workbook(BytesIO(excel.data)).active
    assert aba["A1"].value.startswith("ACME")

    csv = cliente.get(f"/baixar/{identificador}.csv")
    assert "1.1.02;Banco Conta Movimento" in csv.data.decode("utf-8-sig")


def test_gerar_com_dados_de_exemplo(cliente):
    resposta = cliente.post("/gerar", data={"mes": "2026-08", "exemplo": "1"})
    assert resposta.status_code == 200
    assert "Agosto/2026" in resposta.text
    assert "dados de exemplo" in resposta.text


def test_arquivos_obrigatorios(cliente):
    resposta = cliente.post("/gerar", data={"mes": "2026-09"}, content_type="multipart/form-data")
    assert resposta.status_code == 400
    assert "Plano de contas: selecione um arquivo." in resposta.text
    assert "Lançamentos: selecione um arquivo." in resposta.text


def test_erros_de_dados_aparecem_na_pagina(cliente):
    lancamentos = (BytesIO("data;debito;credito;valor\n01/09/2026;9.9;1.1.01;10\n".encode()), "l.csv")
    resposta = cliente.post(
        "/gerar",
        data={"mes": "2026-09", "plano": _arquivo("plano_de_contas.csv"), "lancamentos": lancamentos},
        content_type="multipart/form-data",
    )
    assert resposta.status_code == 400
    assert "conta 9.9 não existe no plano de contas" in resposta.text


def test_extensao_invalida(cliente):
    resposta = cliente.post(
        "/gerar",
        data={
            "mes": "2026-09",
            "plano": (BytesIO(b"x"), "plano.pdf"),
            "lancamentos": _arquivo("lancamentos.csv"),
        },
        content_type="multipart/form-data",
    )
    assert resposta.status_code == 400
    assert "use um arquivo .csv ou .xlsx" in resposta.text


def test_download_inexistente_e_modelos(cliente):
    assert cliente.get("/baixar/naoexiste.xlsx").status_code == 404
    assert cliente.get("/modelo/plano_de_contas.csv").status_code == 200
    assert cliente.get("/modelo/../pyproject.toml").status_code == 404
