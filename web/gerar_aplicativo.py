"""Gera a versão para o computador da Base de Cálculo: um único arquivo .html.

O arquivo abre com duplo clique em qualquer navegador, funciona sem internet
(a biblioteca do Excel vai embutida) e pode ser enviado por e-mail ou pen drive.

Uso:
    python web/gerar_aplicativo.py --saida "Base de Calculo FEMTA.html"
    python web/gerar_aplicativo.py --de BASE_2025.xlsx --saida "Base de Calculo FEMTA.html"

Com --de, o aplicativo já abre com os valores da planilha informada.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

PASTA = Path(__file__).resolve().parent
PAGINA = PASTA / "base_calculo.html"
EXCELJS = PASTA / "vendor" / "exceljs-4.4.0.min.js"
MARCADOR_DADOS = "/*DADOS*/null/*FIM*/"
SCRIPT_CDN = '<script src="https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js"></script>'

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


def gerar(saida: Path, planilha: str | None = None) -> Path:
    pagina = PAGINA.read_text(encoding="utf-8")
    if SCRIPT_CDN not in pagina or MARCADOR_DADOS not in pagina:
        raise SystemExit("web/base_calculo.html não tem os marcadores esperados.")

    # Sem internet: usa as fontes do sistema e embute a biblioteca do Excel.
    pagina = re.sub(r'<link rel="(preconnect|stylesheet)"[^>]*>\n', "", pagina)
    biblioteca = EXCELJS.read_text(encoding="utf-8").replace("</script", "<\\/script")
    pagina = pagina.replace(SCRIPT_CDN, f"<script>\n{biblioteca}\n</script>")

    if planilha:
        dados = dados_da_planilha(planilha)
        pagina = pagina.replace(MARCADOR_DADOS, json.dumps(dados["meses"], separators=(",", ":")))
        if dados["ano"]:
            pagina = pagina.replace('id="ano" min="2000" max="2100" value="2025"',
                                    f'id="ano" min="2000" max="2100" value="{dados["ano"]}"')

    saida.parent.mkdir(parents=True, exist_ok=True)
    saida.write_text(CABECALHO + pagina + "\n</body>\n</html>\n", encoding="utf-8")
    return saida


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--saida", default="Base de Calculo.html", help="arquivo .html a gerar")
    parser.add_argument("--de", help="planilha da base de cálculo para o aplicativo já abrir preenchido")
    args = parser.parse_args()
    caminho = gerar(Path(args.saida), args.de)
    print(f"Aplicativo gerado: {caminho} ({caminho.stat().st_size / 1024:.0f} KB)")


if __name__ == "__main__":
    main()
