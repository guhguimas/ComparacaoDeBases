import json
from config.settings import FINAL_LAYOUT_PATH


class LayoutEngine:
    def __init__(self, layout_path=FINAL_LAYOUT_PATH):
        self.layout_path = layout_path
        self.layout_data = None
        self.columns = []

    def load_layout(self):
        with open(self.layout_path, "r", encoding="utf-8") as file:
            self.layout_data = json.load(file)

        self._validate_layout()
        self.columns = self.layout_data["colunas"]
        return self.columns

    def _validate_layout(self):
        if not isinstance(self.layout_data, dict):
            raise ValueError("O layout precisa estar em formato de objeto.")

        if "colunas" not in self.layout_data:
            raise ValueError("O layout precisa conter a chave 'colunas'.")

        if not isinstance(self.layout_data["colunas"], list):
            raise ValueError("A chave 'colunas' precisa ser uma lista.")

        if len(self.layout_data["colunas"]) == 0:
            raise ValueError("A lista de colunas do layout não pode estar vazia.")