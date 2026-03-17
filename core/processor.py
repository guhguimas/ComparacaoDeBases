from core.layout_engine import LayoutEngine
from core.final_sheet_builder import FinalSheetBuilder
from core.mapping_engine import MappingEngine


class CessaoProcessor:

    def __init__(self):
        self.layout_engine = LayoutEngine()
        self.mapping_engine = MappingEngine()
        self.final_sheet_builder = None

    def executar(self):
        print("Processador iniciado")

        colunas_layout = self.layout_engine.load_layout()
        print("Colunas do layout carregadas com sucesso:")
        for coluna in colunas_layout:
            print(f"- {coluna}")

        mapping = self.mapping_engine.load_mapping()
        print("\nMapeamento carregado com sucesso:")
        for coluna_final, coluna_origem in mapping.items():
            print(f"{coluna_final} <- {coluna_origem}")

        self.final_sheet_builder = FinalSheetBuilder(colunas_layout)
        df_final_vazio = self.final_sheet_builder.create_empty_dataframe()

        print("\nDataFrame final criado com sucesso.")
        print("Colunas do DataFrame final:")
        print(df_final_vazio.columns.tolist())