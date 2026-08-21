from pathlib import Path
import pandas as pd
import duckdb
import win32com.client as win32


def process_result_data(df: pd.DataFrame):
    df = df.iloc[9:, :19]
    
    col_mapping = {
        "order_group": "int64",
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

    for col, dtype in col_mapping.items():
        if dtype == "float64":
            df[col] = pd.to_numeric(df[col], errors="raise")
        elif dtype == "str":
            df[col] = df[col].astype(str)

    clear_cols = ["order_id", 'sku_id']
    for col in clear_cols:
        df[col] = df[col].str.strip()

    df = df.ffill()
    return df

def process_source_data(df: pd.DataFrame):
    df = df.loc[:, ["Order ID", "SKU ID", "SKU Seller Discount"]]

    col_mapping = {
            "order_id": "str",
            "sku_id": "str",
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
        "item_code": "str",
        "item_name": "str",
        "item_unit": "str",
        "tax_ratio": "float64"
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

def sql_process(result_data, source_data, product_data):
    result_data = result_data
    source_data = source_data
    product_data = product_data
    
    sql_query = """
        WITH temp_1 AS ( -- thêm các sequence vào dữ liệu kết quả
            SELECT 
                row_number() OVER () AS raw_seq, -- Thứ tự gốc của các dòng trong file
                row_number() OVER (PARTITION BY order_id ORDER BY order_group) AS item_group,
                rd.order_group,
                rd.order_id,
                rd.sku_id,
                rd.item_code,
                CASE
                    WHEN 
                        pd.item_name IS NOT NULL AND rd.item_price = 0 THEN pd.item_name || '(Hàng khuyến mãi không thu tiền)'
                        ELSE pd.item_name
                END AS item_name,
                pd.item_name,
                rd.item_price,
                pd.item_unit,
                pd.tax_ratio,
                rd.item_quantity,
                ROUND (rd.item_price / (1 + COALESCE(pd.tax_ratio, 0)), 3) AS item_price_novat, -- Tính toán tiền chưa vat, làm tròn 3 số thập phân
                ROUND (ROUND (rd.item_price / (1 + COALESCE(pd.tax_ratio, 0)), 3) * item_quantity) AS total_value_novat, -- Thành tiền chưa vat
                ROUND (ROUND (rd.item_price / (1 + COALESCE(pd.tax_ratio, 0)), 3) * item_quantity * pd.tax_ratio) AS tax_amount, -- Tiền thuế 
                CAST(tax_ratio * 100 AS INTEGER) AS tax_percentage
            FROM result_data rd
            LEFT JOIN product_data pd ON rd.item_code = pd.item_code
            ORDER BY raw_seq
        ),
        temp_2 AS (
            SELECT
                item_group,
                order_group,
                NULL AS item_code,
                'Mã giảm giá và shop voucher' AS item_name,
                'Lần' AS item_unit,
                NULL AS item_quantity,
                ROUND(SUM(discount_per_item), 3) AS item_price_novat, -- Không cần làm tròn giá sản phẩm
                tax_percentage,
                ROUND(SUM(tax_amount)) AS tax_amount, -- Bắt buộc làm tròn
                ROUND(SUM(discount_per_item)) AS total_value_novat, -- Bắt buộc làm tròn
            FROM (
            -- Phân chia tiền voucher vào các mã sản phẩm theo order_id, sku_id trước sau đó gom lại theo % vat
            SELECT
                NULL AS item_group,
                t1.order_group,
                tax_percentage,
                sku_seller_discount,
                -- Phân bổ tiền voucher vào các sản phẩm trong đơn order_id theo từng sku sau đó tính toán tiền trước thuế (Voucher trong file gốc đã bao gồm tiền thuế)
                (sku_seller_discount / COUNT(*) OVER (PARTITION BY t1.order_id, t1.sku_id)) / (1 + t1.tax_ratio) AS discount_per_item,
                (sku_seller_discount / COUNT(*) OVER (PARTITION BY t1.order_id, t1.sku_id)) / (1 + t1.tax_ratio) * t1.tax_ratio AS tax_amount
            FROM temp_1 t1
            LEFT JOIN source_data sd ON t1.order_id = sd.order_id
            AND t1.sku_id = sd.sku_id) AS sub_query
            GROUP BY
                item_group,
                order_group,
                tax_percentage
            HAVING
                tax_amount <> 0 AND total_value_novat <> 0 AND item_price_novat <> 0
        )
        SELECT
            item_group,
            order_group,
            1 AS receive_receipt,
            '11204625' AS buyer_code,
            'Bán cho người tiêu dùng' AS buyer_legal_name,
            'TM/CK' AS pay_method,
            2 AS pay_status,
            'VND' AS currency_code,
            CASE
                WHEN item_price_novat = 0 AND item_code IS NOT NULL THEN 5
                WHEN item_code IS NOT NULL THEN 1
                ELSE 3
            END AS selection,
            item_code,
            item_name,
            item_unit,
            item_quantity,
            item_price_novat,
            total_value_novat,
            tax_percentage,
            tax_amount
        FROM 
        (SELECT 
            item_group,
            order_group,
            item_code,
            item_name,
            item_unit,
            item_quantity,
            item_price_novat,
            total_value_novat,
            tax_percentage,
            tax_amount
        FROM temp_1 t1
        UNION ALL
        SELECT
            item_group,
            order_group,
            item_code,
            item_name,
            item_unit,
            item_quantity,
            item_price_novat,
            total_value_novat,
            tax_percentage,
            tax_amount
        FROM temp_2
        ORDER BY 
        order_group, 
        item_group
        ) AS temp
        """
    result = duckdb.query(sql_query).df()
    return result 

def excel_writer(target_folder_path: Path, data: pd.DataFrame):
    folder_path = target_folder_path 
    files_check = [f for f in folder_path.glob("*.xls")] 
    file_num = len(files_check)
    if not file_num == 1:
        print(f"Lỗi: Phát hiện {file_num} trong thư mục {folder_path}")
        return

    target_template_file = str(files_check[0].resolve())

    # Ghi dữ liệu
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = True

    wb = excel.Workbooks.Open(target_template_file)
    ws = wb.Sheets(1)

    start_row = 11
    start_col = 1
    end_row = start_row + len(data) - 1
    
    COLUMN_MAP = {
        'item_group': 1,
        'order_group': 2,
        'receive_receipt': 3,
        'buyer_code': 4,
        'buyer_legal_name': 13,
        'pay_method': 19,
        'pay_status': 20,
        'currency_code': 21,
        'selection': 25,
        'item_code': 27,
        'item_name': 28,
        'item_unit': 32,
        'item_quantity': 33,
        'item_price_novat': 34,
        'total_value_novat': 35,
        'tax_percentage': 36,
        'tax_amount': 37
    }   

    ws.Range(
        ws.Cells(start_row, start_col),
        ws.Cells(10000, 40)
    ).ClearContents()

    for col, col_idx in COLUMN_MAP.items():
        data_col = data[col]
        data_to_write = [[None if pd.isna(val) else val] for val in data_col.tolist()]

        ws.Range(
            ws.Cells(start_row, col_idx), ws.Cells(end_row, col_idx)
        ).Value = data_to_write
    
    wb.Save()
    del excel


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
    result_data = sql_process(result_data, source_data, product_data)
    excel_writer(DATA_FOLDER_PATH / "invoice_template", result_data)

main()