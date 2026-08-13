import pandas as pd
from pathlib import Path


class DataProcessor:
    """
    Nhận một list các file path, biến đổi và trả về 2 dataframe: Dữ liệu gốc và dữ liệu chứa kết quả
    """

    def __init__(self, data_folder_path: Path):
        self.data_folder = data_folder_path 

    def check_file_data(self):
        # Lấy danh sách file
        files = [file for file in self.data_folder.iterdir()
                 if file.suffix() in [".xlsx", ".xlsm"]]

        # Kiểm tra nếu danh sách rỗng
        if len(files) == 0:
            print(f"Không tìm thấy file trong thư mục {self.data_folder}")
            return

        # Kiểm tra các sheet thuộc file
        for file in files:
            sheet_in_file = pd.ExcelFile(file).sheet_names
        