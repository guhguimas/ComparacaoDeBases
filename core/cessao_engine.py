import pandas as pd
from core.data_transformer import DataTransformer


class CessaoEngine:
    def __init__(self):
        self.allowed_operations = {"DIG", "GDC", "SEM CESSAO", "CAPITAL"}

    def process(self, df_cessao, df_base1, df_base2):
        df_cessao = df_cessao.copy()
        df_base1 = df_base1.copy()
        df_base2 = df_base2.copy()

        self._validate_required_columns(df_cessao, df_base1, df_base2)

        # Trata CCB como número antes do merge
        df_cessao["CCB INVESTIDOR"] = df_cessao["CCB INVESTIDOR"].apply(
            DataTransformer.ccb_to_number
        )

        df_base1["CCB"] = df_base1["CCB"].apply(DataTransformer.ccb_to_number)
        df_base2["CCB"] = df_base2["CCB"].apply(DataTransformer.ccb_to_number)

        df_base1["Operação"] = df_base1["Operação"].apply(DataTransformer.normalize_text)
        df_base2["Operação"] = df_base2["Operação"].apply(DataTransformer.normalize_text)

        # Junta as bases
        df_bases = pd.concat(
            [df_base1[["CCB", "Operação"]], df_base2[["CCB", "Operação"]]],
            ignore_index=True
        )

        df_bases = df_bases[df_bases["CCB"].notna()].copy()
        df_bases = df_bases.drop_duplicates(subset=["CCB"], keep="first")

        # Merge
        df_result = df_cessao.merge(
            df_bases,
            how="left",
            left_on="CCB INVESTIDOR",
            right_on="CCB"
        )

        # IMPORTANTE:
        # resultado do merge fica em coluna separada
        df_result["OPERACAO_FRONT"] = df_result["Operação"].fillna("#N/D")

        # Remove colunas temporárias
        df_result = df_result.drop(columns=["CCB", "Operação"], errors="ignore")

        # Filtro deve ser feito com base na OPERACAO_FRONT
        df_filtered = df_result[df_result["OPERACAO_FRONT"].apply(self._must_keep_row)].copy()

        return df_filtered

    def _must_keep_row(self, value):
        if pd.isna(value):
            return True

        text = str(value).strip()

        if text == "":
            return True

        if text.upper() == "#N/D":
            return True

        if text.upper() in self.allowed_operations:
            return True

        return False

    def _validate_required_columns(self, df_cessao, df_base1, df_base2):
        cessao_required = ["CCB INVESTIDOR", "CONTRATO CRED"]

        # se a Cessão já traz OPERACAO e você quer usar ela na final, vale validar também
        if "OPERACAO" not in df_cessao.columns:
            raise ValueError("Coluna obrigatória não encontrada na Cessão: OPERACAO")

        for column in cessao_required:
            if column not in df_cessao.columns:
                raise ValueError(f"Coluna obrigatória não encontrada na Cessão: {column}")

        base_required = ["CCB", "Operação"]

        for column in base_required:
            if column not in df_base1.columns:
                raise ValueError(f"Coluna obrigatória não encontrada na Base 1: {column}")

        for column in base_required:
            if column not in df_base2.columns:
                raise ValueError(f"Coluna obrigatória não encontrada na Base 2: {column}")