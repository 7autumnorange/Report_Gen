import pandas as pd
import re

class DatHandler:
    def __init__(self, file):
        self.file = file

    def process_dat(self):
        board_name = ""
        test_time = ""
        test_rows = []
        nc_rows = []
        content = self.file.read().decode('utf-8')
        lines = content.splitlines()
        # 提取板名和时间
        for line in lines:
            if line.startswith("! Board Name:"):
                parts = line.split("Time:")
                if len(parts) == 2:
                    board_name = parts[0].replace("! Board Name:", "").strip()
                    test_time = parts[1].strip()

        # 查找step为1的数据行作为数据区第一行
        data_start = 0
        for idx, line in enumerate(lines):
            l = line.strip()
            if not l or l.startswith("!"):
                continue
            m = re.match(r'^(\d+)', l)
            if m and m.group(1) == "1":
                data_start = idx
                break

        i = data_start
        while i < len(lines):
            line = lines[i].strip()
            if not line or line.startswith("!"):
                i += 1
                continue
            match = re.match(r'^(\d+)\s+([^\s]+)', line)
            if match:
                step = match.group(1)
                parts_n = match.group(2)
                tokens = line.split()
                # 校验skip字段，只能是0或1，且其后一列为字母
                skip = ""
                for idx_token in range(len(tokens) - 1):
                    if tokens[idx_token] in ("0", "1") and tokens[idx_token + 1].isalpha():
                        skip = tokens[idx_token]
                        break
                # /NC 特殊处理
                if "/NC" in parts_n.upper():
                    nc_rows.append({
                        "Step": step,
                        "No.": None,
                        "Components": parts_n,
                        "Testable": "NC",
                        "Skip": skip
                    })
                    i += 1
                    continue

                # 新增特殊规则
                testable = ""
                parts_n_upper = parts_n.upper()
                group_y = ["R", "IC", "U", "C", "Q", "D"]
                group_n = ["SG", "L", "NP", "RM", "VM", "PCB"]
                group_l = ["TVS", "PCB"]

                def match_any(comp, group):
                    for g in group:
                        if comp == g:
                            return True
                    return False

                # 优先用/后缀判断
                if "/" in parts_n_upper:
                    suffix = parts_n_upper.split("/")[-1]
                    prefix_match = re.match(r"^([A-Z]+)\d+", parts_n_upper)
                    prefix = prefix_match.group(1) if prefix_match else ""
                    if skip == "0":
                        if match_any(suffix, group_n):
                            testable = "N"
                        elif match_any(suffix, group_y):
                            testable = "Y"
                        elif match_any(suffix, group_l):
                            testable = "L"
                        else:
                            # 如果/后面没命中，再用前缀判断
                            if prefix and match_any(prefix, group_y):
                                testable = "Y"
                            elif prefix and match_any(prefix, group_n):
                                testable = "N"
                            elif prefix and match_any(prefix, group_l):
                                testable = "L"
                            else:
                                testable = ""
                    elif skip == "1":
                        if match_any(suffix, group_n):
                            testable = "N"
                        elif match_any(suffix, group_y):
                            testable = "L"
                        elif match_any(suffix, group_l):
                            testable = "L"
                        else:
                            if prefix and match_any(prefix, group_n):
                                testable = "N"
                            elif prefix and match_any(prefix, group_y):
                                testable = "L"
                            elif prefix and match_any(prefix, group_l):
                                testable = "L"
                            else:
                                testable = ""
                    else:
                        testable = ""
                else:
                    # 没有/，直接判断
                    if skip == "0":
                        testable = "Y"
                    elif skip == "1":
                        testable = ""
                    else:
                        testable = ""

                test_rows.append({
                    "Step": step,
                    "No.": None,
                    "Components": parts_n,
                    "Testable": testable,
                    "Skip": skip
                })
            i += 1

        df = pd.DataFrame(test_rows)
        nc_df = pd.DataFrame(nc_rows)
        # No.为自然序号，Step为原始编号
        if not df.empty:
            df["No."] = range(1, len(df) + 1)
        if not nc_df.empty:
            nc_df["No."] = range(1, len(nc_df) + 1)

        # 计算覆盖率
        y_count = df["Testable"].value_counts().get("Y", 0)
        n_count = df["Testable"].value_counts().get("N", 0)
        l_count = df["Testable"].value_counts().get("L", 0)
        total = y_count + n_count + l_count
        coverage = (y_count + l_count) / total if total > 0 else 0

        return {
            "board_name": board_name,
            "test_time": test_time,
            "data": df,
            "nc_data": nc_df,
            "coverage": coverage
        }
