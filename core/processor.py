from config.settings import OUTPUT_FILE_NAME

from core.layout_engine import LayoutEngine
from core.final_sheet_builder import FinalSheetBuilder
from core.mapping_engine import MappingEngine
from core.final_sheet_mapper import FinalSheetMapper
from core.file_loader import FileLoader
from core.cessao_engine import CessaoEngine
from core.excel_writer import ExcelWriter


class CessaoProcessor:

    def __init__(self):
        self.layout_engine = LayoutEngine()
        self.mapping_engine = MappingEngine()
        self.file_loader = FileLoader()
        self.cessao_engine = CessaoEngine()
        self.excel_writer = ExcelWriter()

        self.final_sheet_builder = None
        self.final_sheet_mapper = None

    def executar(self):
        print("Processador iniciado")

        colunas_layout = self.layout_engine.load_layout()
        mapping = self.mapping_engine.load_mapping()

        self.final_sheet_builder = FinalSheetBuilder(colunas_layout)
        self.final_sheet_mapper = FinalSheetMapper(mapping)

        # ajuste os caminhos abaixo para teste real local
        cessao_path = "dados/cessao.xlsx"
        base1_path = "dados/base1.csv"
        base2_path = "dados/base2.csv"
        output_dir = "output"

        df_cessao = self.file_loader.load_excel(cessao_path)
        df_base1 = self.file_loader.load_csv(base1_path)
        df_base2 = self.file_loader.load_csv(base2_path)

        df_filtered = self.cessao_engine.process(df_cessao, df_base1, df_base2)

        print("Cessão processada com sucesso.")
        print(f"Total de linhas filtradas: {len(df_filtered)}")

        df_final_empty = self.final_sheet_builder.create_empty_dataframe()

        df_final = self.final_sheet_mapper.fill_dataframe(
            source_df=df_filtered,
            target_df=df_final_empty
        )

        output_path = self.excel_writer.save_dataframe(
            dataframe=df_final,
            output_dir=output_dir,
            file_name=OUTPUT_FILE_NAME
        )

        print("\nPlanilha final salva com sucesso.")
        print(f"Arquivo gerado em: {output_path}")