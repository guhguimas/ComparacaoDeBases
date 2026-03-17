class FinalSheetMapper:
    def __init__(self, mapping):
        self.mapping = mapping

    def fill_dataframe(self, source_df, target_df):
        """
        Preenche o DataFrame final com base no mapeamento:
        COLUNA_FINAL <- COLUNA_ORIGEM
        """
        for coluna_final, coluna_origem in self.mapping.items():
            if coluna_origem in source_df.columns:
                target_df[coluna_final] = source_df[coluna_origem]
            else:
                target_df[coluna_final] = None

        return target_df