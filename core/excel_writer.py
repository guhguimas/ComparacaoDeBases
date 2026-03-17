import os


class ExcelWriter:
    def save_dataframe(self, dataframe, output_dir, file_name):
        os.makedirs(output_dir, exist_ok=True)

        output_path = os.path.join(output_dir, file_name)
        dataframe.to_excel(output_path, index=False)

        return output_path