import pandas as pd
from DatabaseHandler import DatabaseHandler
from Converter import ExcelConverter


def main():
    excel_file = 'Marine Hull - Direct 2022.xlsx'
    df_excel = pd.read_excel(excel_file, sheet_name=None, skiprows=4)
    db_handler = DatabaseHandler()
    converter = ExcelConverter()

    for sheet_name, df in df_excel.items():
        df = df.dropna(axis=1, how='all')
        if df.empty:
            continue
        try:
            converted_data = converter.convert_to_dataframe(df)
            db_handler.insert_data(sheet_name, converted_data)
        except Exception as e:
            print(f"Error processing {sheet_name}: {str(e)}")


if __name__ == "__main__":
    main()