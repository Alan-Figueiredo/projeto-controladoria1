from src.utils.Utils import Utils
import os
import pandas as pd

class Report:


    def gerar_relatorio(mapa_balancete, referencias, abas_com_conta,
                    formulas_por_aba, crs_por_aba, arquivo_saida):

        # 1. Inverte a lógica: Descobre quais Abas cuidam de cada CR
        abas_por_cr = {}
        for aba, crs in crs_por_aba.items():
            for cr in crs:
                abas_por_cr.setdefault(cr, set()).add(aba)

        # Garante que abas sem PROCV (mas com contas) entrem no mapa
        for aba, contas in abas_com_conta.items():
            if aba not in crs_por_aba or not crs_por_aba[aba]:
                for c in contas:
                    cr = Utils.extrair_cr(c)
                    abas_por_cr.setdefault(cr, set()).add(aba)

        # ─────────────────────────────────────────────────────────────
        # NOVO: Trava de Limpeza (Remove FIN/ADM das abas específicas)
        for cr in list(abas_por_cr.keys()):
            # Pega todas as abas vinculadas ao CR, mas ignora ADM e FIN
            abas_especificas = {a for a in abas_por_cr[cr] if a not in ("ADM e Dir", "FIN e N.O.")}
            
            # Se houver alguma aba específica (ex: Acessórios), ela ganha exclusividade
            if abas_especificas:
                abas_por_cr[cr] = abas_especificas
        # ─────────────────────────────────────────────────────────────

        # 2. Agrupa todas as contas encontradas nas planilhas por CR
        contas_nas_planilhas_por_cr = {}
        for conta in referencias.keys():
            cr = Utils.extrair_cr(conta)
            contas_nas_planilhas_por_cr.setdefault(cr, set()).add(conta)

        COLUNAS  = ['CR', 'Abas Relacionadas', 'Conta', 'Nome da Conta', 
                    'Débito', 'Crédito', 'Saldo', 'Status', 'Observação']
        COLS_NUM = [5, 6, 7] # Índices das colunas de valores

        # Protege contra arquivo aberto
        if os.path.exists(arquivo_saida):
            try:
                os.rename(arquivo_saida, arquivo_saida)
            except OSError:
                base, ext    = os.path.splitext(arquivo_saida)
                arquivo_saida = f"{base}_novo{ext}"
                print(f"  ⚠️  Arquivo em uso — salvando como '{arquivo_saida}'")

        with pd.ExcelWriter(arquivo_saida, engine='openpyxl') as writer:
            linhas, tipos = [], []
            todos_crs = sorted(abas_por_cr.keys())

            for cr in todos_crs:
                abas_do_cr = " e ".join(sorted(abas_por_cr[cr]))
                contas_planilha = contas_nas_planilhas_por_cr.get(cr, set())

                # Puxa do Balancete apenas as contas que pertencem a este CR
                contas_bal_cr = {
                    cod: info for cod, info in mapa_balancete.items() if info['cr'] == cr
                }
                contas_bal_cr_keys = set(contas_bal_cr.keys())

                pendencias = []
                for cod, info in contas_bal_cr.items():
                    if cod in contas_planilha:
                        continue   # Conta já lançada nas abas -> OK

                    status = Utils._status_cobertura(cod, contas_bal_cr_keys, contas_planilha)

                    if status == 'coberto':
                        continue   # Conta pai já cobre

                    pendencias.append((cod, info, status))

                if not pendencias:
                    continue   # CR sem pendências, pula para o próximo

                pendencias.sort(key=lambda x: x[0])

                for cod, info, status in pendencias:
                    if status == 'parcial':
                        st_cell = '🟡 Parcial'
                        obs     = 'Filhos cobertos, mas conta pai não referenciada'
                    else:
                        st_cell = '🔴 Falta na Aba'
                        obs     = f'Conta não encontrada na(s) aba(s): {abas_do_cr}'

                    linhas.append({
                        'CR':                cr,
                        'Abas Relacionadas': abas_do_cr,
                        'Conta':             cod,
                        'Nome da Conta':     info['nome'],
                        'Débito':            info['debito']  if info['debito']  != 0 else None,
                        'Crédito':           info['credito'] if info['credito'] != 0 else None,
                        'Saldo':             info['valor']   if info['valor']   != 0 else None,
                        'Status':            st_cell,
                        'Observação':        obs,
                    })
                    tipos.append('ausente' if status == 'faltando' else 'parcial')

                

            df = pd.DataFrame(linhas, columns=COLUNAS)
            df.to_excel(writer, sheet_name='Contas Faltando', index=False)

            ws = writer.sheets['Contas Faltando']
            Utils._fmt_header(ws)

            larguras = {'A': 10, 'B': 25, 'C': 25, 'D': 42,
                        'E': 16, 'F': 16, 'G': 16, 'H': 18, 'I': 60}
            for col_letra, w in larguras.items():
                ws.column_dimensions[col_letra].width = w

            Utils._aplicar_estilo(ws, tipos, colunas_valor=COLS_NUM)

        print(f"\n✅ Relatório gerado: {arquivo_saida}")