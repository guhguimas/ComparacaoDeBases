import pandas as pd

from core.layout_engine import LayoutEngine
from core.final_sheet_builder import FinalSheetBuilder
from core.mapping_engine import MappingEngine
from core.final_sheet_mapper import FinalSheetMapper


class CessaoProcessor:

    def __init__(self):
        self.layout_engine = LayoutEngine()
        self.mapping_engine = MappingEngine()
        self.final_sheet_builder = None
        self.final_sheet_mapper = None

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

        print("\nDataFrame final vazio criado com sucesso.")
        print(df_final_vazio.columns.tolist())

        # DataFrame de teste simulando a planilha Cessão já tratada
        df_origem_teste = pd.DataFrame([
            {
                "CONTRATO CRED": "123456",
                "CCB INVESTIDOR": 987654321,
                "TAXA CESSÃO": "1.89%",
                "OPERACAO": "DIG",
                "FUNDO": "FUNDO A",
                "DATA CESSÃO": "2026-03-17",
                "ORIGEM": "CESSAO"
            },
            {
                "CONTRATO CRED": "789012",
                "CCB INVESTIDOR": 123456789,
                "TAXA CESSÃO": "2.10%",
                "OPERACAO": "CAPITAL",
                "FUNDO": "FUNDO B",
                "DATA CESSÃO": "2026-03-18",
                "ORIGEM": "CESSAO"
            }
        ])

        self.final_sheet_mapper = FinalSheetMapper(mapping)
        df_final_preenchido = self.final_sheet_mapper.fill_dataframe(
            source_df=df_origem_teste,
            target_df=df_final_vazio
        )

        print("\nDataFrame final preenchido com sucesso:")
        print(df_final_preenchido)