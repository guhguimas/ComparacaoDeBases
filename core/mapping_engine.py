import json
from config.settings import COLUMN_MAPPING_PATH


class MappingEngine:
    def __init__(self, mapping_path=COLUMN_MAPPING_PATH):
        self.mapping_path = mapping_path
        self.mapping_data = None

    def load_mapping(self):
        with open(self.mapping_path, "r", encoding="utf-8") as file:
            self.mapping_data = json.load(file)

        self._validate_mapping()
        return self.mapping_data

    def _validate_mapping(self):
        if not isinstance(self.mapping_data, dict):
            raise ValueError("O mapeamento precisa estar em formato de objeto.")

        if len(self.mapping_data) == 0:
            raise ValueError("O mapeamento não pode estar vazio.")