"""Gera o aplicativo da Base de Cálculo (com o parecer mensal em Word).

Versão para o computador (padrão): um único arquivo .html que abre com duplo
clique, funciona sem internet (as bibliotecas do Excel e do Word vão embutidas)
e pode ser enviado por e-mail ou pen drive.

Versão online (--online): a página para publicar no claude.ai, que carrega as
bibliotecas da CDN.

Uso:
    python web/gerar_aplicativo.py --saida "Base de Calculo FEMTA.html"
    python web/gerar_aplicativo.py --de BASE_2025.xlsx --saida "Base de Calculo FEMTA.html"
    python web/gerar_aplicativo.py --online --de BASE_2025.xlsx --saida pagina.html

Com --de, o aplicativo já abre com os valores da planilha informada.
"""

from __future__ import annotations

import argparse
import base64
import json
import re
import sys
from pathlib import Path

PASTA = Path(__file__).resolve().parent
PAGINA = PASTA / "base_calculo.html"
MODELO_PARECER = PASTA / "modelos" / "parecer_femta.docx"
BIBLIOTECAS = {
    '<script src="https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js"></script>':
        PASTA / "vendor" / "exceljs-4.4.0.min.js",
    '<script src="https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js"></script>':
        PASTA / "vendor" / "jszip-3.10.1.min.js",
}
MARCADOR_DADOS = "/*DADOS*/null/*FIM*/"
MARCADOR_MODELO = "/*MODELO_PARECER*/null/*FIM*/"

CABECALHO = """<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<style>:root{color-scheme:light}body{margin:0;font:14px system-ui,sans-serif}img{max-width:100%}[hidden]{display:none!important}</style>
</head>
<body>
"""


def dados_da_planilha(caminho: str) -> dict:
    sys.path.insert(0, str(PASTA.parent))
    from balancete.base_calculo import ler_base

    base = ler_base(caminho)
    meses = {m: {k: int(round(v * 100)) for k, v in valores.items() if v} for m, valores in base["meses"].items()}
    return {"meses": meses, "ano": base["ano"]}


def gerar(saida: Path, planilha: str | None = None, online: bool = False) -> Path:
    pagina = PAGINA.read_text(encoding="utf-8")
    faltando = [m for m in [MARCADOR_DADOS, MARCADOR_MODELO, *BIBLIOTECAS] if m not in pagina]
    if faltando:
        raise SystemExit(f"web/base_calculo.html não tem os marcadores esperados: {faltando}")

    modelo = base64.b64encode(MODELO_PARECER.read_bytes()).decode("ascii")
    pagina = pagina.replace(MARCADOR_MODELO, json.dumps(modelo))

    if not online:
        # Sem internet: usa as fontes do sistema e embute as bibliotecas.
        pagina = re.sub(r'<link rel="(preconnect|stylesheet)"[^>]*>\n', "", pagina)
        for tag, arquivo in BIBLIOTECAS.items():
            codigo = arquivo.read_text(encoding="utf-8").replace("</script", "<\\/script")
            pagina = pagina.replace(tag, f"<script>\n{codigo}\n</script>")

    if planilha:
        dados = dados_da_planilha(planilha)
        pagina = pagina.replace(MARCADOR_DADOS, json.dumps(dados["meses"], separators=(",", ":")))
        if dados["ano"]:
            pagina = pagina.replace('id="ano" min="2000" max="2100" value="2025"',
                                    f'id="ano" min="2000" max="2100" value="{dados["ano"]}"')

    saida.parent.mkdir(parents=True, exist_ok=True)
    conteudo = pagina if online else CABECALHO + pagina + "\n</body>\n</html>\n"
    saida.write_text(conteudo, encoding="utf-8")
    return saida


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--saida", default="Base de Calculo.html", help="arquivo .html a gerar")
    parser.add_argument("--de", help="planilha da base de cálculo para o aplicativo já abrir preenchido")
    parser.add_argument("--online", action="store_true", help="gera a página para publicar (bibliotecas pela internet)")
    args = parser.parse_args()
    caminho = gerar(Path(args.saida), args.de, args.online)
    print(f"Aplicativo gerado: {caminho} ({caminho.stat().st_size / 1024:.0f} KB)")


if __name__ == "__main__":
    main()
