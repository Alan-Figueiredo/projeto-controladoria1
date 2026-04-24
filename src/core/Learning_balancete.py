from src.config import *
import pandas as pd

class Learning_balancete:

    def _engine(path):
        return 'pyxlsb' if str(path).lower().endswith('.xlsb') else 'openpyxl'

    def _br_to_float(s):
        try:
            return float(str(s).strip().replace('.', '').replace(',', '.'))
        except (ValueError, AttributeError):
            return None


    def _detectar_col_conta(df):
        for col in df.columns:
            if str(col).strip().lower() == 'conta':
                return col
        for col in df.columns:
            try:
                amostra = df[col].dropna().astype(str).head(100)
                if sum(1 for x in amostra if REGEX_CODIGO.match(x.strip())) >= 3:
                    return col
            except Exception:
                continue
        return None

    def _detectar_col_nome(df, col_conta):
        candidatos = ['descrição', 'descricao', 'nome', 'name', 'description',
                    'histórico', 'historico', 'título', 'titulo']
        for col in df.columns:
            if col == col_conta:
                continue
            if str(col).strip().lower() in candidatos:
                return col
        for col in df.columns:
            if col == col_conta:
                continue
            amostra = df[col].dropna().astype(str).head(20)
            if amostra.str.len().mean() > 10:
                return col
        return None


    def _detectar_col_cr(df):
        candidatos = ['cr', 'centro de resultado', 'centro resultado',
                    'centro de custo', 'centro custo', 'cc']
        for col in df.columns:
            if str(col).strip().lower() in candidatos:
                return col
        for col in df.columns:
            amostra = df[col].dropna().astype(str).head(50)
            if sum(1 for x in amostra if re.match(r'^\d{6}$', x.strip())) >= 3:
                return col
        return None


    def _detectar_col_debito(df):
        candidatos = ['débito', 'debito', 'débitos', 'debitos', 'déb', 'deb', 'debit']
        for col in df.columns:
            if str(col).strip().lower() in candidatos:
                return col
        return None


    def _detectar_col_credito(df):
        candidatos = ['crédito', 'credito', 'créditos', 'creditos', 'créd', 'cred', 'credit']
        for col in df.columns:
            if str(col).strip().lower() in candidatos:
                return col
        return None


    def _detectar_col_valor(df, col_conta):
        prioridade = ['saldo mês', 'saldo mes', 'saldo atual', 'saldo_mes', 'saldo_atual']
        for col in df.columns:
            if str(col).strip().lower() in prioridade:
                return col
        for col in reversed(df.columns):
            if col == col_conta:
                continue
            amostra = df[col].dropna().astype(str).head(50)
            validos = sum(1 for x in amostra if Learning_balancete._br_to_float(x) is not None)
            if validos >= 3:
                return col
        return None


    def ler_balancete(wb_path, nome_aba):
        """Retorna {codigo_conta: {nome, cr, debito, credito, valor}}"""
        print(f"\n📊 Lendo balancete: '{nome_aba}'")

        # Detecta linha real do cabeçalho procurando por 'Conta'
        header_row = 0
        engine     = Learning_balancete._engine(wb_path)
        for skip in range(10):
            df_test = pd.read_excel(wb_path, sheet_name=nome_aba, header=skip,
                                    nrows=1, dtype=str, engine=engine)
            cols = [str(c).strip().lower() for c in df_test.columns]
            if any(c == 'conta' for c in cols):
                header_row = skip
                break

        df = pd.read_excel(wb_path, sheet_name=nome_aba, header=header_row,
                        dtype=str, engine=engine)
        for col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(df[col])

        print(f"   → Colunas encontradas ({len(df.columns)}):")
        for c in df.columns:
            dtype   = df[c].dtype
            n_num   = pd.to_numeric(df[c], errors='coerce').notna().sum()
            amostra = df[c].dropna().astype(str).head(3).tolist()
            print(f"      [{dtype}] '{c}'  nums={n_num}  ex: {amostra}")

        col_conta   = COLUNA_CONTA_BALANCETE   or Learning_balancete._detectar_col_conta(df)
        if col_conta is None:
            raise ValueError(f"Coluna de conta não encontrada em '{nome_aba}'. "
                            "Defina COLUNA_CONTA_BALANCETE.")

        col_nome    = COLUNA_NOME_BALANCETE    or Learning_balancete._detectar_col_nome(df, col_conta)
        col_cr      = COLUNA_CR_BALANCETE      or Learning_balancete._detectar_col_cr(df)
        col_debito  = COLUNA_DEBITO_BALANCETE  or Learning_balancete._detectar_col_debito(df)
        col_credito = COLUNA_CREDITO_BALANCETE or Learning_balancete._detectar_col_credito(df)
        col_valor   = COLUNA_VALOR_BALANCETE   or Learning_balancete._detectar_col_valor(df, col_conta)

        if col_valor is None:
            raise ValueError("Coluna de saldo não encontrada. Defina COLUNA_VALOR_BALANCETE.")

        print(f"   → Conta: '{col_conta}' | Nome: '{col_nome}' | CR: '{col_cr}' | "
            f"Débito: '{col_debito}' | Crédito: '{col_credito}' | Saldo: '{col_valor}'")

        def _float(row, col):
            if col is None:
                return 0.0
            v = row.get(col)
            if v is None or (isinstance(v, float) and pd.isna(v)):
                return 0.0
            try:
                return float(v)
            except (ValueError, TypeError):
                return 0.0

        mapa = {}
        for _, row in df.iterrows():
            cod = str(row[col_conta]).strip() if pd.notna(row[col_conta]) else ''
            if not REGEX_CODIGO.match(cod):
                continue

            nome_conta = ''
            if col_nome:
                v = row.get(col_nome)
                if v is not None and not (isinstance(v, float) and pd.isna(v)):
                    nome_conta = str(v).strip()

            cr_val = ''
            if col_cr:
                v = row.get(col_cr)
                if v is not None and not (isinstance(v, float) and pd.isna(v)):
                    cr_val = str(v).strip()
            if not cr_val:
                partes = cod.split('-')
                cr_val = partes[-1].strip() if len(partes) >= 2 else ''

            debito  = _float(row, col_debito)
            credito = _float(row, col_credito)
            valor   = _float(row, col_valor)

            if cod in mapa:
                mapa[cod]['debito']  += debito
                mapa[cod]['credito'] += credito
                mapa[cod]['valor']   += valor
            else:
                mapa[cod] = {
                    'nome':    nome_conta,
                    'cr':      cr_val,
                    'debito':  debito,
                    'credito': credito,
                    'valor':   valor,
                }

        print(f"   → {len(mapa)} contas no Balancete")
        return mapa
