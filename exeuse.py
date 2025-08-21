import sys
import pandas as pd
import re
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
    QFileDialog, QLabel, QProgressBar, QMessageBox, QTableWidget, QTableWidgetItem, QHeaderView
)
from handlers.dat_handler import DatHandler
from handlers.csv_handler import CsvHandler
from handlers.dcl_handler import DclHandler
from utils.excel_utils import fill_template_excel

def expand_reference(ref_str):
    refs = []
    for part in str(ref_str).replace(' ', '').split(','):
        m = re.match(r'([A-Za-z]+)(\d+)-([A-Za-z]+)?(\d+)', part)
        if m:
            prefix1, start, prefix2, end = m.groups()
            prefix2 = prefix2 if prefix2 else prefix1
            for i in range(int(start), int(end)+1):
                refs.append(f"{prefix2.upper()}{i}")
        else:
            if part:
                refs.append(part.upper())
    return refs

def build_ref_to_desc(csv_df):
    ref_to_desc = {}
    for _, row in csv_df.iterrows():
        desc = row["Description"]
        for ref in expand_reference(row["Reference"]):
            ref_to_desc[ref] = desc
    return ref_to_desc

def get_description(components, ref_to_desc):
    if pd.isna(components):
        return ""
    parts = []
    for part in str(components).replace(',', '/').replace(' ', '').split('/'):
        part = part.strip().upper()
        if part:
            part_main = part.split('_')[0].split('-')[0]
            parts.append(part_main)
    for part in parts:
        if part in ref_to_desc:
            return ref_to_desc[part]
    return ""

def gen_remark(row):
    if row.get("Testable") == "N":
        return "No test point"
    elif row.get("Testable") == "L":
        comps = str(row.get("Components", ""))
        if "/" in comps:
            return f"it is in parallel with {comps.split('/')[-1].strip()}"
        elif "," in comps:
            return f"it is in parallel with {comps.split(',')[-1].strip()}"
        else:
            return "it is in parallel with"
    else:
        return row.get("Remark", "")

def extract_main_comp(comp):
    if pd.isna(comp):
        return ""
    comp = str(comp).replace(' ', '').upper()
    main = comp.split('/')[0].split(',')[0]
    main = main.split('_')[0].split('-')[0]
    return main

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Report Auto-generated Tool")
        self.resize(1200, 900)
        self.init_ui()
        self.reset_state()

    def init_ui(self):
        widget = QWidget()
        self.setCentralWidget(widget)
        layout = QVBoxLayout(widget)

        # 文件选择（顺序修正：dat, dcl/csv, csv, 模板）
        file_layout = QHBoxLayout()
        self.dat_btn = QPushButton("选择 .dat 文件")
        self.dcl_btn = QPushButton("选择 .dcl/.csv 文件")
        self.csv_btn = QPushButton("选择 .csv 文件")
        self.template_btn = QPushButton("选择 Excel 模板")
        file_layout.addWidget(self.dat_btn)
        file_layout.addWidget(self.dcl_btn)
        file_layout.addWidget(self.csv_btn)
        file_layout.addWidget(self.template_btn)
        layout.addLayout(file_layout)

        # 文件路径显示
        self.file_labels = [QLabel("未选择") for _ in range(4)]
        file_label_layout = QHBoxLayout()
        for label in self.file_labels:
            file_label_layout.addWidget(label)
        layout.addLayout(file_label_layout)

        # 进度条
        self.progress = QProgressBar()
        layout.addWidget(self.progress)

        # 操作按钮
        self.process_btn = QPushButton("🚀 开始处理文件")
        layout.addWidget(self.process_btn)

        # 主表格
        layout.addWidget(QLabel("唯一主编号（主表）"))
        self.table = QTableWidget()
        layout.addWidget(self.table)
        self.export_btn = QPushButton("📥 导出主表Excel")
        self.export_btn.setEnabled(False)
        layout.addWidget(self.export_btn)

        # /NC表格
        layout.addWidget(QLabel("/NC Components 信息（仅显示/可导出）"))
        self.nc_table = QTableWidget()
        self.nc_table.setMaximumHeight(150)
        layout.addWidget(self.nc_table)
        self.export_nc_btn = QPushButton("导出/NC Excel")
        self.export_nc_btn.setEnabled(False)
        layout.addWidget(self.export_nc_btn)

        # 重复项表格
        layout.addWidget(QLabel("Components 重复项（仅显示/可导出）"))
        self.dup_table = QTableWidget()
        self.dup_table.setMaximumHeight(150)
        layout.addWidget(self.dup_table)
        self.export_dup_btn = QPushButton("导出重复项Excel")
        self.export_dup_btn.setEnabled(False)
        layout.addWidget(self.export_dup_btn)

        # 无Description表格
        layout.addWidget(QLabel("Description 为空的 Components 信息（仅显示/可导出）"))
        self.no_desc_table = QTableWidget()
        self.no_desc_table.setMaximumHeight(150)
        layout.addWidget(self.no_desc_table)
        self.export_no_desc_btn = QPushButton("导出无Description Excel")
        self.export_no_desc_btn.setEnabled(False)
        layout.addWidget(self.export_no_desc_btn)

        # 信号绑定（顺序修正）
        self.dat_btn.clicked.connect(lambda: self.select_file(0, "DAT 文件 (*.dat)"))
        self.dcl_btn.clicked.connect(lambda: self.select_file(1, "DCL/CSV 文件 (*.dcl *.csv)"))
        self.csv_btn.clicked.connect(lambda: self.select_file(2, "CSV 文件 (*.csv)"))
        self.template_btn.clicked.connect(lambda: self.select_file(3, "Excel 文件 (*.xlsx)"))
        self.process_btn.clicked.connect(self.process_files)
        self.export_btn.clicked.connect(self.export_excel)
        self.export_nc_btn.clicked.connect(lambda: self.export_df(self.nc_data, "nc_data.xlsx"))
        self.export_dup_btn.clicked.connect(lambda: self.export_df(self.df_dup, "dup_data.xlsx"))
        self.export_no_desc_btn.clicked.connect(lambda: self.export_df(self.df_no_desc, "no_desc_data.xlsx"))

    def reset_state(self):
        self.file_paths = [None, None, None, None]
        self.header_data = None
        self.component_df = None
        self.dat_data = None
        self.csv_df = None
        self.ref_to_desc = {}
        self.df_unique = None
        self.df_dup = pd.DataFrame()
        self.df_no_desc = pd.DataFrame()
        self.nc_data = pd.DataFrame()
        self.excel_bytes = None

    def select_file(self, idx, filter_str):
        path, _ = QFileDialog.getOpenFileName(self, "选择文件", "", filter_str)
        if path:
            self.file_paths[idx] = path
            self.file_labels[idx].setText(path.split("/")[-1])

    def process_files(self):
        self.progress.setValue(5)
        dat_path, dcl_path, csv_path, template_path = self.file_paths
        if not dcl_path or not template_path:
            QMessageBox.warning(self, "提示", "请上传 .dcl 文件和 Excel 模板。")
            self.progress.setValue(0)
            return

        # dcl
        self.progress.setValue(20)
        with open(dcl_path, "rb") as f:
            self.header_data, self.component_df = DclHandler(f).process_dcl()

        # dat
        self.progress.setValue(40)
        self.dat_data = None
        if dat_path:
            with open(dat_path, "rb") as f:
                self.dat_data = DatHandler(f).process_dat()

        # csv
        self.progress.setValue(60)
        self.csv_df = CsvHandler().process_csv(csv_path) if csv_path else pd.DataFrame()
        print("csv_df columns:", self.csv_df.columns)
        print("csv_df head:\n", self.csv_df.head())
        if csv_path and self.csv_df.empty:
            QMessageBox.critical(self, "错误", "CSV文件解析失败，请将当前CSV文件另存为CSV UTF8格式。")
            return
        # 列名标准化，防止有空格或大小写问题
        self.csv_df.columns = [str(col).strip() for col in self.csv_df.columns]
        if not self.csv_df.empty and "Reference" in self.csv_df.columns and "Description" in self.csv_df.columns:
            self.ref_to_desc = build_ref_to_desc(self.csv_df)
        else:
            self.ref_to_desc = {}
            QMessageBox.warning(self, "提示", "CSV文件缺少 Reference 或 Description 列，或内容为空。")
            return

        self.progress.setValue(70)

        # 数据处理与去重
        self.df_unique = pd.DataFrame()
        self.df_dup = pd.DataFrame()
        self.df_no_desc = pd.DataFrame()
        self.nc_data = pd.DataFrame()
        if self.dat_data is not None and not self.dat_data["data"].empty:
            print("ref_to_desc keys:", list(self.ref_to_desc.keys()))
            print("dat_data['data'] head:\n", self.dat_data["data"].head())

            self.dat_data["data"]["Description"] = self.dat_data["data"]["Components"].apply(
                lambda x: get_description(x, self.ref_to_desc)
            )
            print("dat_data['data'] with Description head:\n", self.dat_data["data"].head())

            df_with_desc = self.dat_data["data"][self.dat_data["data"]["Description"].notna() & (self.dat_data["data"]["Description"] != "")].copy()
            print("df_with_desc head:\n", df_with_desc.head())

            self.df_no_desc = self.dat_data["data"][self.dat_data["data"]["Description"].isna() | (self.dat_data["data"]["Description"] == "")].copy()
            self.nc_data = self.dat_data["nc_data"].copy() if "nc_data" in self.dat_data else pd.DataFrame()
            df_with_desc["Remark"] = df_with_desc.apply(gen_remark, axis=1)

            # 去重逻辑
            if not df_with_desc.empty:
                df_with_desc["main_comp"] = df_with_desc["Components"].apply(extract_main_comp)
                df_with_desc["testable_notna"] = df_with_desc["Testable"].apply(lambda x: pd.notna(x) and str(x).strip() != "")
                df_with_desc["has_np"] = df_with_desc["Components"].apply(lambda x: "/NP" in str(x).upper())
                df_with_desc["has_slash"] = df_with_desc["Components"].apply(lambda x: "/" in str(x))
                df_with_desc["has_dash"] = df_with_desc["Components"].apply(lambda x: "-" in str(x))
                df_with_desc["has_underscore"] = df_with_desc["Components"].apply(lambda x: "_" in str(x))
                df_with_desc["comp_len"] = df_with_desc["Components"].apply(lambda x: len(str(x)))
                idx = (
                    df_with_desc
                    .sort_values(
                        ["main_comp", "testable_notna", "has_np", "has_slash", "has_dash", "has_underscore", "comp_len"],
                        ascending=[True, False, False, False, False, False, True]
                    )
                    .groupby("main_comp", as_index=False)
                    .head(1)
                    .index
                )
                self.df_unique = df_with_desc.loc[idx].copy()
                self.df_dup = df_with_desc.drop(idx).copy()
                self.df_unique = self.df_unique.drop(columns=["main_comp", "testable_notna", "has_np", "has_slash", "has_underscore", "has_dash", "comp_len"])
                self.df_dup = self.df_dup.drop(columns=["main_comp", "testable_notna", "has_np", "has_slash", "has_underscore", "has_dash", "comp_len"])
            else:
                self.df_unique = df_with_desc
                self.df_dup = pd.DataFrame()

            # 显示表格
            self.show_table(self.df_unique, self.table)
            self.show_table(self.nc_data, self.nc_table)
            self.show_table(self.df_dup, self.dup_table)
            self.show_table(self.df_no_desc, self.no_desc_table)
            self.export_btn.setEnabled(not self.df_unique.empty)
            self.export_nc_btn.setEnabled(not self.nc_data.empty)
            self.export_dup_btn.setEnabled(not self.df_dup.empty)
            self.export_no_desc_btn.setEnabled(not self.df_no_desc.empty)
            self.progress.setValue(85)
            QMessageBox.information(self, "完成", "处理完成！可下载Excel文件。")
        else:
            QMessageBox.warning(self, "提示", "未检测到dat文件内容")
            self.df_unique = pd.DataFrame()
            self.df_dup = pd.DataFrame()
            self.df_no_desc = pd.DataFrame()
            self.nc_data = pd.DataFrame()
            self.show_table(self.df_unique, self.table)
            self.show_table(self.nc_data, self.nc_table)
            self.show_table(self.df_dup, self.dup_table)
            self.show_table(self.df_no_desc, self.no_desc_table)
            self.export_btn.setEnabled(False)
            self.export_nc_btn.setEnabled(False)
            self.export_dup_btn.setEnabled(False)
            self.export_no_desc_btn.setEnabled(False)
            self.progress.setValue(0)

    def show_table(self, df, table_widget):
        table_widget.clearContents()
        if df is None or df.empty:
            table_widget.setRowCount(0)
            table_widget.setColumnCount(0)
            return
        columns = [str(col) for col in df.columns]
        table_widget.setColumnCount(len(columns))
        table_widget.setRowCount(len(df))
        table_widget.setHorizontalHeaderLabels(columns)
        for i, row in enumerate(df.itertuples(index=False)):
            for j, value in enumerate(row):
                item = QTableWidgetItem(str(value))
                table_widget.setItem(i, j, item)
        table_widget.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)

    def export_excel(self):
        if self.df_unique is None or self.df_unique.empty:
            QMessageBox.warning(self, "提示", "没有可导出的数据")
            return
        save_path, _ = QFileDialog.getSaveFileName(self, "保存Excel文件", "processed_data.xlsx", "Excel 文件 (*.xlsx)")
        if not save_path:
            return
        try:
            # 计算 Parts Testable Percentage
            if not self.df_unique.empty and "Testable" in self.df_unique.columns:
                total = len(self.df_unique)
                testable = (self.df_unique["Testable"] == "Y").sum()
                coverage = round(testable / total * 100, 2) if total > 0 else 0
            else:
                coverage = 0

            # 生成Excel
            excel_bytes = fill_template_excel(
                self.file_paths[3],  # template_file
                self.component_df,
                self.csv_df,
                self.header_data,
                {
                    "data": self.df_unique,
                    "nc_data": self.nc_data,
                    "board_name": self.dat_data.get("board_name", "") if self.dat_data else "",
                    "test_time": self.dat_data.get("test_time", "") if self.dat_data else "",
                    "coverage": coverage
                }
            )
            with open(save_path, "wb") as f:
                f.write(excel_bytes.getvalue())
            QMessageBox.information(self, "导出成功", f"文件已保存到：{save_path}")
            self.progress.setValue(100)
        except Exception as e:
            QMessageBox.critical(self, "导出失败", f"导出异常：{e}")

    def export_df(self, df, default_name):
        if df is None or df.empty:
            QMessageBox.warning(self, "提示", "没有可导出的数据")
            return
        save_path, _ = QFileDialog.getSaveFileName(self, "导出Excel", default_name, "Excel 文件 (*.xlsx)")
        if not save_path:
            return
        try:
            df.to_excel(save_path, index=False)
            QMessageBox.information(self, "导出成功", f"文件已保存到：{save_path}")
        except Exception as e:
            QMessageBox.critical(self, "导出失败", f"导出异常：{e}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
