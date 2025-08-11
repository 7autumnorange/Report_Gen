import pandas as pd
import io

class DclHandler:
    def __init__(self, file):
        self.file = file

    def process_dcl(self):
        header_fields = []
        header_values = []
        in_header = False
        component_fields = []
        component_rows = []
        in_component = False

        content = self.file.read().decode('utf-8')
        f = io.StringIO(content)

        for line in f:
            line = line.strip()
            # Header_Data 区块
            if line.startswith("// Header_Data"):
                in_header = True
                in_component = False
                continue
            if in_header and line.startswith("//"):
                header_fields = [x.strip() for x in line[2:].split(',')]
                continue
            if in_header and line and not line.startswith("//"):
                header_values = [x.strip() for x in line.split(',')]
                in_header = False  # 只取第一条
                continue
            # Component_Data 区块
            if line.startswith("// Component_Data"):
                in_component = True
                continue
            if in_component and line.startswith("//"):
                component_fields = [x.strip() for x in line[2:].split(',')]
                continue
            if in_component and line and not line.startswith("//"):
                row = [x.strip() for x in line.split(',')]
                if len(row) == len(component_fields):
                    component_rows.append(row)
        # Header_Data
        field_map = {name: idx for idx, name in enumerate(header_fields)}
        board_name = header_values[field_map.get("BoardName", -1)] if "BoardName" in field_map else ""
        # 取PASS/FAIL字段
        passfail = header_values[field_map.get("PASS/FAIL", -1)] if "PASS/FAIL" in field_map else ""
        result = header_values[field_map.get("Result", -1)] if "Result" in field_map else ""
        date = header_values[field_map.get("Date", -1)] if "Date" in field_map else ""
        time = header_values[field_map.get("Time", -1)] if "Time" in field_map else ""
        test_time = ""
        if date and time and len(date) == 8 and len(time) == 6:
            test_time = f"{date[:4]}-{date[4:6]}-{date[6:]} {time[:2]}:{time[2:4]}:{time[4:]}"
        header_data = {
            "passfail": passfail,
            "result": result,
            "test_time": test_time,
            "board_name": board_name
        }
        # Component_Data 映射
        comp_map = {
            "No.": "StepNum",
            "Component Name": "PartName",
            "Testing Type": "Type",
            "Testing Point No.": ("HPin", "LPin"),
            "Reference Value": "Std_V",
            "Lower Limit(%)": "HLim",
            "Upper Limit(%)": "LLim",
            "Measured Value": "Msr_V",
            "Result": "Result"
        }
        comp_field_idx = {name: idx for idx, name in enumerate(component_fields)}
        data = []
        for row in component_rows:
            item = {}
            for k, v in comp_map.items():
                if k == "Testing Point No.":
                    hpin = row[comp_field_idx.get("HPin", -1)] if "HPin" in comp_field_idx else ""
                    lpin = row[comp_field_idx.get("LPin", -1)] if "LPin" in comp_field_idx else ""
                    item[k] = f"{hpin},{lpin}" if hpin and lpin else hpin or lpin
                else:
                    idx = comp_field_idx.get(v, -1) if isinstance(v, str) else -1
                    item[k] = row[idx] if idx >= 0 else ""
            # 处理Result字段
            item["Result"] = "PASS" if item["Result"] == "0" else "FAIL"
            data.append(item)
        component_df = pd.DataFrame(data)
        return header_data, component_df