import pandas as pd

class CsvHandler:
    def __init__(self):
        pass

    def process_csv(self, file):
        # 尝试用多种编码读取，并捕获空数据异常
        for encoding in ["utf-8", "gbk", "latin1"]:
            try:
                df = pd.read_csv(file, encoding=encoding)
                if not df.empty:
                    break
            except pd.errors.EmptyDataError:
                return pd.DataFrame()  # 文件为空，返回空DataFrame
            except Exception:
                continue
        else:
            return pd.DataFrame()  # 全部尝试失败，返回空DataFrame

        # 只保留Reference和Description列（如果存在）
        needed_cols = []
        if "Reference" in df.columns:
            needed_cols.append("Reference")
        if "Description" in df.columns:
            needed_cols.append("Description")
        if needed_cols:
            df = df[needed_cols]
            df["Reference"] = df["Reference"].astype(str)
            df["Description"] = df["Description"].astype(str)
        return df