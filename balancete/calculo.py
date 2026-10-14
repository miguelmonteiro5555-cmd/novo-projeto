"""Cálculo do balancete de verificação mensal."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from decimal import Decimal

from .formato import formatar_saldo, formatar_valor
from .leitura import Conta, ErroDeDados, Lancamento

ZERO = Decimal("0.00")


@dataclass
class LinhaBalancete:
    conta: Conta
    nivel: int
    sintetica: bool
    # Saldos com sinal: positivo = devedor, negativo = credor.
    saldo_anterior: Decimal = ZERO
    debitos: Decimal = ZERO
    creditos: Decimal = ZERO

    @property
    def saldo_atual(self) -> Decimal:
        return self.saldo_anterior + self.debitos - self.creditos

    @property
    def zerada(self) -> bool:
        return not (self.saldo_anterior or self.debitos or self.creditos)


@dataclass
class Balancete:
    ano: int
    mes: int
    linhas: list[LinhaBalancete]
    total_debitos: Decimal
    total_creditos: Decimal
    total_saldo_anterior: Decimal
    total_saldo_atual: Decimal
    avisos: list[str] = field(default_factory=list)

    @property
    def fechado(self) -> bool:
        """Débitos = créditos e a soma dos saldos é zero."""
        return self.total_debitos == self.total_creditos and self.total_saldo_atual == 0


def _chave_ordenacao(codigo: str) -> tuple:
    return tuple((0, int(p), "") if p.isdigit() else (1, 0, p) for p in codigo.split("."))


def _ancestrais(codigo: str, plano: dict[str, Conta]) -> list[str]:
    """Contas sintéticas acima de `codigo` (hierarquia separada por pontos)."""
    partes = codigo.split(".")
    prefixos = (".".join(partes[:i]) for i in range(len(partes) - 1, 0, -1))
    return [p for p in prefixos if p in plano]


def gerar_balancete(
    plano: dict[str, Conta],
    lancamentos: list[Lancamento],
    ano: int,
    mes: int,
    saldos_iniciais: dict[str, Decimal] | None = None,
    incluir_zeradas: bool = False,
) -> Balancete:
    """Gera o balancete do mês.

    Saldo anterior = saldo inicial + movimento de todos os lançamentos
    anteriores ao mês. Lançamentos posteriores ao mês são ignorados.
    Os valores são lançados apenas em contas analíticas e somados nas
    sintéticas (contas-pai, pelo código: 1 > 1.1 > 1.1.01).
    """
    if not 1 <= mes <= 12:
        raise ErroDeDados([f"Mês inválido: {mes}"])
    saldos_iniciais = saldos_iniciais or {}
    inicio = date(ano, mes, 1)
    fim = date(ano + 1, 1, 1) if mes == 12 else date(ano, mes + 1, 1)

    pais = {codigo: _ancestrais(codigo, plano) for codigo in plano}
    sinteticas = {pai for lista in pais.values() for pai in lista}
    linhas = {
        codigo: LinhaBalancete(conta, len(pais[codigo]) + 1, codigo in sinteticas)
        for codigo, conta in plano.items()
    }

    problemas = _validar(plano, sinteticas, lancamentos, saldos_iniciais)
    if problemas:
        raise ErroDeDados(problemas)

    def acumular(codigo: str, atributo: str, valor: Decimal) -> None:
        for alvo in [codigo, *pais[codigo]]:
            linha = linhas[alvo]
            setattr(linha, atributo, getattr(linha, atributo) + valor)

    for codigo, saldo in saldos_iniciais.items():
        acumular(codigo, "saldo_anterior", saldo)

    for lanc in lancamentos:
        if lanc.data < inicio:
            acumular(lanc.debito, "saldo_anterior", lanc.valor)
            acumular(lanc.credito, "saldo_anterior", -lanc.valor)
        elif lanc.data < fim:
            acumular(lanc.debito, "debitos", lanc.valor)
            acumular(lanc.credito, "creditos", lanc.valor)

    analiticas = [l for c, l in linhas.items() if c not in sinteticas]
    balancete = Balancete(
        ano=ano,
        mes=mes,
        linhas=[
            linhas[c]
            for c in sorted(linhas, key=_chave_ordenacao)
            if incluir_zeradas or not linhas[c].zerada
        ],
        total_debitos=sum((l.debitos for l in analiticas), ZERO),
        total_creditos=sum((l.creditos for l in analiticas), ZERO),
        total_saldo_anterior=sum((l.saldo_anterior for l in analiticas), ZERO),
        total_saldo_atual=sum((l.saldo_atual for l in analiticas), ZERO),
    )
    balancete.avisos = _avisos(balancete, analiticas)
    return balancete


def _validar(
    plano: dict[str, Conta],
    sinteticas: set[str],
    lancamentos: list[Lancamento],
    saldos_iniciais: dict[str, Decimal],
) -> list[str]:
    problemas = []

    def checar_conta(codigo: str, origem: str) -> None:
        if not codigo:
            problemas.append(f"{origem}: conta não informada")
        elif codigo not in plano:
            problemas.append(f"{origem}: conta {codigo} não existe no plano de contas")
        elif codigo in sinteticas:
            problemas.append(f"{origem}: conta {codigo} é sintética (não aceita lançamentos)")

    for lanc in lancamentos:
        origem = f"Lançamentos, linha {lanc.linha}"
        checar_conta(lanc.debito, f"{origem} (débito)")
        checar_conta(lanc.credito, f"{origem} (crédito)")
        if lanc.debito and lanc.debito == lanc.credito:
            problemas.append(f"{origem}: débito e crédito na mesma conta")
        if lanc.valor <= 0:
            problemas.append(f"{origem}: valor deve ser maior que zero")

    for codigo in saldos_iniciais:
        checar_conta(codigo, "Saldos iniciais")
    return problemas


def _avisos(balancete: Balancete, analiticas: list[LinhaBalancete]) -> list[str]:
    avisos = []
    if balancete.total_saldo_anterior != 0:
        avisos.append(
            f"A soma dos saldos anteriores não é zero ({formatar_valor(balancete.total_saldo_anterior)}): "
            "verifique os saldos iniciais."
        )
    if balancete.total_debitos == 0 and balancete.total_creditos == 0:
        avisos.append("Nenhum lançamento no mês.")
    for linha in analiticas:
        natureza, saldo = linha.conta.natureza, linha.saldo_atual
        if (natureza == "D" and saldo < 0) or (natureza == "C" and saldo > 0):
            avisos.append(
                f"Conta {linha.conta.codigo} ({linha.conta.descricao}) com saldo invertido: "
                f"natureza {natureza}, saldo {formatar_saldo(saldo)}."
            )
    return avisos
