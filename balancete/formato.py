"""Formatação de valores no padrão brasileiro."""

from decimal import Decimal


def formatar_valor(valor: Decimal) -> str:
    """1234.5 -> "1.234,50"."""
    return f"{valor:,.2f}".replace(",", "_").replace(".", ",").replace("_", ".")


def formatar_saldo(valor: Decimal) -> str:
    """Saldo com indicador D/C: 100 -> "100,00 D", -100 -> "100,00 C"."""
    if valor == 0:
        return "0,00"
    return f"{formatar_valor(abs(valor))} {'D' if valor > 0 else 'C'}"
