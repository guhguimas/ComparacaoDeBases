import pandas as pd


class FileLoader:
    def load_excel(self, file_path):
        return pd.read_excel(file_path, dtype=str)

    def load_csv(self, file_path, separator=";", encoding="utf-8-sig"):
        try:
            return pd.read_csv(
                file_path,
                sep=separator,
                dtype=str,
                encoding=encoding
            )
        except UnicodeDecodeError:
            return pd.read_csv(
                file_path,
                sep=separator,
                dtype=str,
                encoding="latin1"
            )