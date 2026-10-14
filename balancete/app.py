"""Aplicação web: python -m balancete.app (abre em http://127.0.0.1:5000)."""

from __future__ import annotations

import tempfile
import uuid
from collections import OrderedDict
from datetime import date
from io import BytesIO
from pathlib import Path

from flask import Flask, abort, render_template, request, send_file

from .calculo import Balancete, gerar_balancete
from .exportacao import MESES, exportar_csv, exportar_excel, titulo
from .formato import formatar_saldo, formatar_valor
from .leitura import ErroDeDados, ler_lancamentos, ler_plano_de_contas, ler_saldos_iniciais

EXEMPLOS = Path(__file__).resolve().parent.parent / "exemplos"
EXTENSOES = {".csv", ".xlsx", ".xlsm"}
MAX_RESULTADOS = 20


def _mes_padrao() -> str:
    hoje = date.today()
    ano, mes = (hoje.year, hoje.month - 1) if hoje.month > 1 else (hoje.year - 1, 12)
    return f"{ano}-{mes:02d}"


def criar_app() -> Flask:
    app = Flask(__name__)
    app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024
    app.jinja_env.filters["valor"] = formatar_valor
    app.jinja_env.filters["saldo"] = formatar_saldo
    # Balancetes gerados, para download (mantém só os mais recentes em memória).
    resultados: OrderedDict[str, tuple[Balancete, str]] = OrderedDict()

    def formulario(erros=None, **dados):
        dados.setdefault("mes", _mes_padrao())
        return render_template("index.html", erros=erros or [], dados=dados), 400 if erros else 200

    @app.get("/")
    def inicio():
        return formulario()

    @app.post("/gerar")
    def gerar():
        dados = {
            "mes": request.form.get("mes", ""),
            "empresa": request.form.get("empresa", "").strip(),
            "incluir_zeradas": bool(request.form.get("incluir_zeradas")),
        }
        try:
            ano, mes = (int(p) for p in dados["mes"].split("-"))
            if not 1 <= mes <= 12:
                raise ValueError
        except ValueError:
            return formulario(["Informe o mês de referência."], **dados)

        usar_exemplo = "exemplo" in request.form
        with tempfile.TemporaryDirectory() as pasta:
            if usar_exemplo:
                arquivos = {
                    "plano": EXEMPLOS / "plano_de_contas.csv",
                    "lancamentos": EXEMPLOS / "lancamentos.csv",
                    "saldos": EXEMPLOS / "saldos_iniciais.csv",
                }
            else:
                arquivos, erros = _salvar_uploads(Path(pasta))
                if erros:
                    return formulario(erros, **dados)
            try:
                plano = ler_plano_de_contas(arquivos["plano"])
                lancamentos = ler_lancamentos(arquivos["lancamentos"])
                saldos = ler_saldos_iniciais(arquivos["saldos"]) if arquivos.get("saldos") else {}
                balancete = gerar_balancete(
                    plano, lancamentos, ano, mes, saldos, dados["incluir_zeradas"]
                )
            except ErroDeDados as erro:
                return formulario(erro.problemas, **dados)

        identificador = uuid.uuid4().hex
        resultados[identificador] = (balancete, dados["empresa"])
        while len(resultados) > MAX_RESULTADOS:
            resultados.popitem(last=False)

        return render_template(
            "resultado.html",
            b=balancete,
            titulo=titulo(balancete, dados["empresa"]),
            mes_nome=f"{MESES[mes - 1]}/{ano}",
            id=identificador,
            exemplo=usar_exemplo,
        )

    @app.get("/baixar/<identificador>.<formato>")
    def baixar(identificador: str, formato: str):
        if identificador not in resultados or formato not in ("xlsx", "csv"):
            abort(404)
        balancete, empresa = resultados[identificador]
        nome = f"balancete_{balancete.ano}-{balancete.mes:02d}.{formato}"
        with tempfile.TemporaryDirectory() as pasta:
            caminho = Path(pasta) / nome
            if formato == "xlsx":
                exportar_excel(balancete, caminho, empresa)
            else:
                exportar_csv(balancete, caminho)
            conteudo = BytesIO(caminho.read_bytes())
        return send_file(conteudo, as_attachment=True, download_name=nome)

    @app.get("/modelo/<nome>")
    def modelo(nome: str):
        permitidos = {"plano_de_contas.csv", "lancamentos.csv", "saldos_iniciais.csv"}
        if nome not in permitidos:
            abort(404)
        return send_file(EXEMPLOS / nome, as_attachment=True, download_name=nome)

    @app.errorhandler(413)
    def arquivo_grande(_erro):
        return formulario(["Arquivo muito grande (limite de 20 MB)."])

    return app


def _salvar_uploads(pasta: Path) -> tuple[dict[str, Path], list[str]]:
    arquivos: dict[str, Path] = {}
    erros = []
    rotulos = {"plano": "Plano de contas", "lancamentos": "Lançamentos", "saldos": "Saldos iniciais"}
    for campo, rotulo in rotulos.items():
        enviado = request.files.get(campo)
        if not enviado or not enviado.filename:
            if campo != "saldos":
                erros.append(f"{rotulo}: selecione um arquivo.")
            continue
        extensao = Path(enviado.filename).suffix.lower()
        if extensao not in EXTENSOES:
            erros.append(f"{rotulo}: use um arquivo .csv ou .xlsx.")
            continue
        destino = pasta / f"{campo}{extensao}"
        enviado.save(destino)
        arquivos[campo] = destino
    return arquivos, erros


def main() -> None:
    import argparse

    parser = argparse.ArgumentParser(prog="balancete-app", description="Abre a aplicação web do balancete.")
    parser.add_argument("--porta", type=int, default=5000)
    parser.add_argument("--host", default="127.0.0.1", help="use 0.0.0.0 para acessar de outros computadores da rede")
    args = parser.parse_args()
    print(f"Balancete disponível em http://{'localhost' if args.host == '127.0.0.1' else args.host}:{args.porta}")
    criar_app().run(host=args.host, port=args.porta)


if __name__ == "__main__":
    main()
