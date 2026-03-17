import pandas as pd


class FinalSheetBuilder:
    def __init__(self, columns):
        self.columns = columns

    def create_empty_dataframe(self):
        return pd.DataFrame(columns=self.columns)