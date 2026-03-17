import pandas as pd


class DataTransformer:
    @staticmethod
    def normalize_text(value):
        if pd.isna(value):
            return ""
        return str(value).strip()

    @staticmethod
    def ccb_to_number(value):
        text = DataTransformer.normalize_text(value)
        text = text.replace("-", "").replace(" ", "")

        if text == "":
            return pd.NA

        if text.endswith(".0"):
            text = text[:-2]

        if text.isdigit():
            return int(text)

        only_digits = "".join(char for char in text if char.isdigit())
        if only_digits:
            return int(only_digits)

        return pd.NA