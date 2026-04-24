import re

ARQUIVO_ENTRADA          = r"data/input/planilha-teste.xlsb"
ABA_BALANCETE            = "Balancete"
ARQUIVO_SAIDA            = r"data/output/resultado_comparacao.xlsx"

ABAS_IGNORAR_EXTRA       = {'MENU', 'Instruções', 'Aux', 'Config'}

COLUNA_CONTA_BALANCETE   = "Conta"
COLUNA_NOME_BALANCETE    = "Descrição"
COLUNA_CR_BALANCETE      = None          # extraído do código da conta
COLUNA_DEBITO_BALANCETE  = "Débitos"
COLUNA_CREDITO_BALANCETE = "Créditos"
COLUNA_VALOR_BALANCETE   = "SALDO MÊS"
# ─────────────────────────────────────────────────────────────────

REGEX_FORMULA         = re.compile(r'CONTAS_BALANCETE\s*\(\s*"([^"]+)"\s*\)', re.IGNORECASE)
REGEX_CODIGO          = re.compile(r'^\d[\d\-]*\d-\d{6}$')
REGEX_CR_VAL          = re.compile(r'C\.R\.\s*(\d{6})', re.IGNORECASE)
REGEX_CONTA_ANALITICA = re.compile(r'^\d{8,}-\d{6}$')   # mínimo 8 dígitos antes do hífen

XL_CELL_TYPE_FORMULAS = -4123
