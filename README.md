# Balancete Mensal Automatizado

Gera o **balancete de verificação** de um mês a partir do plano de contas e dos
lançamentos contábeis (CSV ou Excel). O resultado é uma planilha Excel formatada
com saldo anterior, débitos, créditos e saldo atual de cada conta, com as contas
sintéticas totalizadas e a conferência de fechamento.

## Instalação

Requer Python 3.10+.

```bash
pip install -e ".[dev]"
```

## Aplicação web

```bash
python -m balancete.app        # ou: balancete-app
```

Abra **http://localhost:5000** no navegador, escolha o mês, envie o plano de
contas e os lançamentos (e, se quiser, os saldos iniciais) e clique em
**Gerar balancete**. O balancete aparece na tela, com os totais, a situação de
fechamento e os avisos, e pode ser baixado em Excel ou CSV. O botão
**Testar com dados de exemplo** mostra o resultado sem enviar arquivos, e a
seção *Formato dos arquivos* tem modelos para baixar.

Para usar a partir de outros computadores da rede: `python -m balancete.app --host 0.0.0.0`.

## Linha de comando

```bash
python -m balancete --mes 2026-09 \
  --plano exemplos/plano_de_contas.csv \
  --lancamentos exemplos/lancamentos.csv \
  --saldos exemplos/saldos_iniciais.csv \
  --empresa "Comércio Exemplo Ltda"
```

Saída no terminal:

```
Comércio Exemplo Ltda — Balancete de Verificação — Setembro/2026
  Contas no relatório: 26
  Saldo anterior:      0,00
  Débitos do mês:      77.385,00
  Créditos do mês:     77.385,00
  Saldo atual:         0,00
  Situação:            FECHADO ✔
Arquivo gerado: saida/balancete_2026-09.xlsx
```

| Opção | Descrição |
|---|---|
| `--mes` | Mês de referência: `AAAA-MM` ou `MM/AAAA` (obrigatório) |
| `--plano` | Plano de contas `.csv` ou `.xlsx` (obrigatório) |
| `--lancamentos` | Lançamentos `.csv` ou `.xlsx` (obrigatório) |
| `--saldos` | Saldos iniciais (opcional) |
| `--saida` | Arquivo de saída `.xlsx` ou `.csv` (padrão: `saida/balancete_AAAA-MM.xlsx`) |
| `--empresa` | Nome da empresa no título |
| `--incluir-zeradas` | Mostra também contas sem saldo nem movimento |

Códigos de saída: `0` balancete fechado, `1` erro nos dados, `2` balancete não fechou.

## Formato dos arquivos de entrada

CSV com separador `;` ou `,`, ou Excel (primeira aba). Os cabeçalhos aceitam
maiúsculas e acentos (`Débito` = `debito`). Valores podem vir como `1.234,56`,
`1234.56` ou `R$ 1.234,56`; datas como `31/12/2026` ou `2026-12-31`.

**Plano de contas**: `codigo`, `descricao`, `natureza` (opcional, `D` ou `C`)

```
codigo;descricao;natureza
1;ATIVO;D
1.1;Ativo Circulante;D
1.1.01;Caixa;D
```

A hierarquia vem do código separado por pontos (`1` > `1.1` > `1.1.01`). Contas
que têm filhas são **sintéticas**: somam as filhas e não recebem lançamentos.
A natureza serve para avisar quando uma conta fica com saldo invertido.

**Lançamentos**: `data`, `debito`, `credito`, `valor`, `historico` (opcional).
Cada linha é uma partida dobrada (uma conta a débito e outra a crédito).

```
data;debito;credito;valor;historico
15/09/2026;1.1.02;3.1.01;22.000,00;Venda à vista
```

Os lançamentos podem conter vários meses: os anteriores ao mês de referência
formam o saldo anterior e os posteriores são ignorados.

**Saldos iniciais** (opcional): `codigo`, `saldo`, `dc` (opcional, `D` ou `C`).
São os saldos antes do primeiro lançamento do arquivo. Sem a coluna `dc`,
saldo positivo é devedor e negativo é credor.

## Validações

O programa para e lista **todos** os problemas encontrados quando há:

- conta inexistente no plano ou lançamento em conta sintética;
- débito e crédito na mesma conta, valor zero/negativo;
- datas ou valores inválidos, colunas obrigatórias ausentes.

E emite avisos (na tela e na aba **Avisos** do Excel) para saldos iniciais que
não somam zero, contas com saldo invertido e mês sem lançamentos.

## Testes

```bash
python -m pytest
```

## Estrutura

```
balancete/
  leitura.py     leitura e validação de CSV/Excel
  calculo.py     cálculo do balancete e totalização das contas sintéticas
  exportacao.py  geração do Excel e do CSV
  formato.py     formatação de valores (1.234,56)
  cli.py         linha de comando
  app.py         aplicação web (Flask)
  templates/     páginas da aplicação web
exemplos/        arquivos de exemplo
tests/           testes automatizados
```
