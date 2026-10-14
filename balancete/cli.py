"""Linha de comando: python -m balancete --mes 2026-09 ..."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from .calculo import gerar_balancete
from .exportacao import exportar_csv, exportar_excel, titulo
from .formato import formatar_saldo, formatar_valor
from .leitura import ErroDeDados, ler_lancamentos, ler_plano_de_contas, ler_saldos_iniciais


def _mes(texto: str) -> tuple[int, int]:
    for separador in ("-", "/"):
        partes = texto.split(separador)
        if len(partes) == 2 and all(p.isdigit() for p in partes):
            a, b = (int(p) for p in partes)
            ano, mes = (a, b) if a > 12 else (b, a)  # aceita 2026-09 ou 09/2026
            if 1 <= mes <= 12:
                return ano, mes
    raise argparse.ArgumentTypeError(f"mês inválido: {texto!r} (use AAAA-MM ou MM/AAAA)")


def criar_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="balancete",
        description="Gera o balancete de verificação mensal a partir dos lançamentos.",
    )
    parser.add_argument("--mes", required=True, type=_mes, help="mês de referência (AAAA-MM ou MM/AAAA)")
    parser.add_argument("--plano", required=True, help="plano de contas (.csv ou .xlsx)")
    parser.add_argument("--lancamentos", required=True, help="lançamentos contábeis (.csv ou .xlsx)")
    parser.add_argument("--saldos", help="saldos iniciais (.csv ou .xlsx), opcional")
    parser.add_argument("--saida", help="arquivo de saída .xlsx ou .csv (padrão: saida/balancete_AAAA-MM.xlsx)")
    parser.add_argument("--empresa", default="", help="nome da empresa para o título")
    parser.add_argument("--incluir-zeradas", action="store_true", help="mostrar contas sem saldo nem movimento")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = criar_parser().parse_args(argv)
    ano, mes = args.mes

    try:
        plano = ler_plano_de_contas(args.plano)
        lancamentos = ler_lancamentos(args.lancamentos)
        saldos = ler_saldos_iniciais(args.saldos) if args.saldos else {}
        balancete = gerar_balancete(plano, lancamentos, ano, mes, saldos, args.incluir_zeradas)
    except ErroDeDados as erro:
        print("Erro nos dados de entrada:", file=sys.stderr)
        for problema in erro.problemas:
            print(f"  - {problema}", file=sys.stderr)
        return 1

    saida = Path(args.saida or f"saida/balancete_{ano}-{mes:02d}.xlsx")
    if saida.suffix.lower() == ".csv":
        exportar_csv(balancete, saida)
    else:
        exportar_excel(balancete, saida, args.empresa)

    print(titulo(balancete, args.empresa))
    print(f"  Contas no relatório: {len(balancete.linhas)}")
    print(f"  Saldo anterior:      {formatar_saldo(balancete.total_saldo_anterior)}")
    print(f"  Débitos do mês:      {formatar_valor(balancete.total_debitos)}")
    print(f"  Créditos do mês:     {formatar_valor(balancete.total_creditos)}")
    print(f"  Saldo atual:         {formatar_saldo(balancete.total_saldo_atual)}")
    print(f"  Situação:            {'FECHADO ✔' if balancete.fechado else 'NÃO FECHOU ✘'}")
    for aviso in balancete.avisos:
        print(f"  Aviso: {aviso}")
    print(f"Arquivo gerado: {saida}")
    return 0 if balancete.fechado else 2
