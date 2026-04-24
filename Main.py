from src.config import *
from src.utils.Utils import Utils
from src.core.Learning_formules import Learning_formules as LF
from src.core.Learning_balancete import Learning_balancete as LB
from src.services.Report import Report as RP
import os
import shutil


def main():
    print("=" * 60)
    print("  Comparador de Balancete focado em Centro de Custo (CR)")
    print("=" * 60)

    ARQUIVO_ENTRADA = Utils.preparar_arquivo()

    if not ARQUIVO_ENTRADA:
        return

    
    if not os.path.exists(ARQUIVO_ENTRADA):
        print(f"\n❌ Arquivo '{ARQUIVO_ENTRADA}' não encontrado.")
        return

    abas_ignorar = ABAS_IGNORAR_EXTRA | {ABA_BALANCETE}

    # 1. Extrai fórmulas CONTAS_BALANCETE + CRs via PROCV
    formulas_por_aba, crs_por_aba = LF.ler_formulas_e_crs_win32(ARQUIVO_ENTRADA, abas_ignorar)

    if not formulas_por_aba:
        print("\n❌ Nenhuma fórmula CONTAS_BALANCETE encontrada.")
        return

    # 2. Consolida referências
    referencias, abas_com_conta = Utils.consolidar_referencias(formulas_por_aba)
    print(f"\n✔  Contas únicas nas abas:  {len(referencias)}")
    print(f"✔  Abas com CRs via PROCV:  {len(crs_por_aba)}")

    # 3. Lê o Balancete
    mapa_balancete = LB.ler_balancete(ARQUIVO_ENTRADA, ABA_BALANCETE)

    # 4. Gera relatório (agora por CR)
    RP.gerar_relatorio(mapa_balancete, referencias, abas_com_conta,
                    formulas_por_aba, crs_por_aba, ARQUIVO_SAIDA)

    # 5. Resumo no terminal (agora focado em CR)
    print("\n📊 Resumo de Pendências por Centro de Custo (CR):")
    
    shutil.rmtree("data/temp")
    # Reconstrói o mapa de CRs para o resumo visual
    abas_por_cr = {}
    for aba, crs in crs_por_aba.items():
        for cr in crs:
            abas_por_cr.setdefault(cr, set()).add(aba)

    # ─────────────────────────────────────────────────────────────
    # NOVO: Mesma Trava de Limpeza no resumo de tela
    for cr in list(abas_por_cr.keys()):
        abas_especificas = {a for a in abas_por_cr[cr] if a not in ("ADM e Dir", "FIN e N.O.")}
        if abas_especificas:
            abas_por_cr[cr] = abas_especificas
    # ─────────────────────────────────────────────────────────────

    for cr in sorted(abas_por_cr.keys()):
        abas_do_cr = ", ".join(sorted(abas_por_cr[cr]))
        contas_planilha = {c for c in referencias.keys() if Utils.extrair_cr(c) == cr}
        contas_bal_cr = {cod for cod, info in mapa_balancete.items() if info['cr'] == cr}
        
        faltando = sum(
            1 for cod in contas_bal_cr
            if cod not in contas_planilha
            and Utils._status_cobertura(cod, contas_bal_cr, contas_planilha) != 'coberto'
        )
        
        status = '✅' if faltando == 0 else f'⚠️  {faltando} pendência(s)'
        print(f"   {status}  [CR {cr}]  Planilhas: {abas_do_cr}")

if __name__ == "__main__":
    main()