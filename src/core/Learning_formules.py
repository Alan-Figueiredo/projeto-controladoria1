import win32com.client as win32
import os
from src.config import *
from src.utils import Utils

class Learning_formules:

    def ler_formulas_e_crs_win32(wb_path, abas_ignorar):

        """
        Para cada aba retorna:
        formulas_por_aba : {aba: {(row,col): {formula, contas, value}}}
        crs_por_aba      : {aba: set(cr_codes)}
        """
        import win32com.client as win32

        abs_path = os.path.abspath(wb_path)
        print("  🔌 Abrindo arquivo com Excel (win32com)...")

        xl = win32.DispatchEx("Excel.Application")
        try:
            xl.Visible = False
        except Exception:
            pass
        try:
            xl.DisplayAlerts = False
        except Exception:
            pass

        formulas_por_aba = {}
        crs_por_aba      = {}

        try:
            wb = xl.Workbooks.Open(abs_path, ReadOnly=True, UpdateLinks=False)

            for sheet in wb.Sheets:
                nome = sheet.Name
                if nome in abas_ignorar:
                    print(f"  ⏭️  Ignorando aba: '{nome}'")
                    continue

                print(f"  📖 Lendo aba: '{nome}'")
                celulas_formula = {}
                crs_da_aba      = set()

                try:
                    formula_range = sheet.UsedRange.SpecialCells(XL_CELL_TYPE_FORMULAS)
                except Exception:
                    print(f"     ⏭️  Nenhuma fórmula encontrada. Pulando.")
                    continue

                for area in formula_range.Areas:
                    for cell in area.Cells:
                        try:
                            formula    = str(cell.Formula or '')
                            formula_br = str(cell.FormulaLocal or '')
                        except Exception:
                            continue

                        formula_up    = formula.upper()
                        formula_br_up = formula_br.upper()

                        # ── CONTAS_BALANCETE ──────────────────────────
                        if 'CONTAS_BALANCETE' in formula_up:
                            contas = REGEX_FORMULA.findall(formula)
                            if contas:
                                try:
                                    val = float(cell.Value) if cell.Value is not None else 0.0
                                except (ValueError, TypeError):
                                    val = 0.0
                                celulas_formula[(cell.Row, cell.Column)] = {
                                    'formula': formula,
                                    'contas':  contas,
                                    'value':   val,
                                }

                        # ── PROCV / VLOOKUP → captura CR ─────────────
                        tem_procv = (any(x in formula_up    for x in ('PROCV', 'VLOOKUP')) or
                                    any(x in formula_br_up for x in ('PROCV', 'VLOOKUP')))

                        if tem_procv:
                            refs_chave = re.findall(
                                r'(?:PROCV|VLOOKUP)\s*\(\s*\$?([A-Z]+)\$?(\d+)\s*[;,]',
                                formula_br, re.IGNORECASE
                            )
                            for col_letra, row_num in refs_chave:
                                try:
                                    val_celula = str(
                                        sheet.Cells(int(row_num),
                                                    Utils._letra_para_col(col_letra)).Value or ''
                                    ).strip()
                                    match = REGEX_CR_VAL.search(val_celula)
                                    if match:
                                        cr_code = match.group(1)
                                        crs_da_aba.add(cr_code)
                                        print(f"     🏷️  CR capturado: {cr_code}  ('{val_celula}')")
                                except Exception:
                                    pass

                if celulas_formula:
                    formulas_por_aba[nome] = celulas_formula
                    n_contas = sum(len(c['contas']) for c in celulas_formula.values())
                    print(f"     → {len(celulas_formula)} células CONTAS_BALANCETE | "
                        f"{n_contas} referências | {len(crs_da_aba)} CRs via PROCV")
                else:
                    print(f"     ⏭️  Nenhuma fórmula CONTAS_BALANCETE. Pulando.")

                if crs_da_aba:
                    crs_por_aba[nome] = crs_da_aba

            wb.Close(SaveChanges=False)
        finally:
            xl.Quit()

        return formulas_por_aba, crs_por_aba