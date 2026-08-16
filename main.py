from pathlib import Path
import pandas as pd
from pandas import DataFrame


def process_result_data(df: pd.DataFrame):
    df = df.iloc[9:, :19]
    col_mapping = {
        "stt": "int64",
        "order_id": "str",
        "sku_id": "str",
        "product_name": "str",
        "variation": "str",
        "sku_unit_original_price": "float64",
        "quantity": "float64",
        "sku_subtotal_before_discount": "float64",
        "sku_subtotal_after_discount": "float64",
        "shipping_fee_after_discount": "float64",
        "original_shipping_fee": "float64",
        "order_amount": "float64",
        "shipping_provider_name": "str",
        "date": "str",
        "item_code": "str",
        "item_name": "str",
        "item_price": "float64",
        "item_quantity": "float64",
        "total_order_value": "float64"
    }

    df.columns = [col for col in col_mapping.keys()]
    df.to_excel("result.xlsx")

    for col, dtype in col_mapping.items():
        if dtype == "float64":
            df[col] = pd.to_numeric(df[col], errors="raise")
        elif dtype == "str":
            df[col] = df[col].astype(str)

    clear_cols = ["order_id", 'sku_id']
    for col in clear_cols:
        df[col] = df[col].str.strip()

    df = df.iloc[:, :-3]
    df = df.ffill()
    return df

def process_source_data(df:DataFrame):
    df = df.loc[:, ["Order ID", "SKU Seller Discount"]]

    col_mapping = {
            "order_id": "str",
            "sku_seller_discount": "float64"
    }

    df.columns = [col for col in col_mapping.keys()]

    for col, dtype in col_mapping.items():
        if dtype == "float64":
            df[col] = pd.to_numeric(df[col], errors="raise")
        elif dtype == "str":
            df[col] = df[col].astype(str)
    return df

def process_ecom_data(ecom_folder: Path):
    files = [file for file in ecom_folder.glob("*.xlsm") if file.name.startswith("~")]
    for file in ecom_folder.glob("*.xlsm"):
        sheet_in_file = pd.ExcelFile(file).sheet_names

        sheet_to_read = ["NGUON", "GHEPSP"]
        has_error = False
        for sheet in sheet_to_read:
            if sheet not in sheet_in_file:
                print(f"Không tìm thấy sheet {sheet} in {files}")
                has_error = True

        if has_error:
            print(f"Có lỗi, đang bỏ qua file {file}")
            continue

        result_data_df = pd.read_excel(file, sheet_name = "GHEPSP")
        result_data_df = process_result_data(result_data_df)

        source_data_df = pd.read_excel(file, sheet_name = "NGUON")
        source_data_df = process_source_data(source_data_df)

        return result_data_df, source_data_df

def process_product_data(file_path: Path):
    df = pd.read_excel(file_path)

    col_mapping = {
        "product_code": "str",
        "product_name": "str",
        "tax_percent": "float64"
    }

    for col, dtype in col_mapping.items():
        if dtype == "float64":
            df[col] = pd.to_numeric(df[col], errors="raise")
        else:
            df[col] = df[col].astype("str")

    return df

def check_file_num(data_folder: Path):
    """
    Đảm bảo các điều kiện sau:
    - Thư mục ecom_process_data (chứa dữ liệu của ecom) chỉ chứa file có đuôi .xlsm)
    - Thư mục invoice_template chỉ chứa đúng 1 file hoá đơn
    - Thư mục product_data chỉ chứa đúng 1 file thông tin sản phẩm
    """
    has_error = False

    # Kiểm tra thư mục ecom_processed_data
    ecom_files = [file for file in Path(data_folder / "ecom_processed_data").glob("*.xlsm")]
    if len(ecom_files) == 0:
        print("Không có file chứa dữ liệu ecom nào, vui lòng thêm vào")

    invoice_files = [file for file in Path(data_folder / "invoice_template").glob("*.xls")]
    if len(invoice_files) == 0:
        print(f"Không có file hoá đơn nào trong thư mục {data_folder / "invoice_template"}")
    elif len(invoice_files) > 1:
        print(f"Có nhiều hơn 1 file hoá đơn, vui lòng chỉ sử dụng 1 file")

    product_files = [file for file in Path(data_folder / "product_data").glob("*.xlsx")]
    if len(product_files) == 0:
        print(f"Không có file dữ liệu sản phẩm nào nằm trong thư mục {Path(data_folder / "product_data")}")

    if not has_error:
        print("Không có lỗi, tiếp tục đọc file")

    return has_error


BASE_PATH = Path().cwd()
DATA_FOLDER_PATH = BASE_PATH / "data"


def main():
    check_result = check_file_num(DATA_FOLDER_PATH)
    if check_result:
        return

    result_data, source_data = process_ecom_data(
        ecom_folder = DATA_FOLDER_PATH / "ecom_processed_data"
    )

    product_data = process_product_data(DATA_FOLDER_PATH / "product_data" / "Thông tin sản phẩm.xlsx")

main()