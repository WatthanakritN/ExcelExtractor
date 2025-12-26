import pandas as pd


class ExcelExtractor:
    def __init__(self, sheet_name: str):
        self.sheet_name = sheet_name

    def extract_single_file(self, file_path: str) -> pd.DataFrame:
        try:
            # อ่านทั้งหมดก่อน
            df = pd.read_excel(
                file_path,
                sheet_name=self.sheet_name,
                header=1   # A2 เป็น header
            )
        except Exception as e:
            print(f"[SKIP] {file_path} : {e}")
            return pd.DataFrame()

        if df.empty:
            return df

        # จำกัด column ไม่เกิน AH (index 33) A - AH
        df = df.iloc[:, :34]

        # หยุดเมื่อ column A ว่าง
        first_col = df.columns[0]
        stop_index = df[df[first_col].isna()].index
        if len(stop_index) > 0:
            df = df.loc[: stop_index[0] - 1]

        return df
