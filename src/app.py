import streamlit as st
import pandas as pd
import re
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
            # 先按下划线分割，再按短横线分割，只取第一个编号
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
    # 取第一个/前的编号，再去掉下划线及后缀
    if pd.isna(comp):
        return ""
    comp = str(comp).replace(' ', '').upper()
    main = comp.split('/')[0].split(',')[0]
    main = main.split('_')[0].split('-')[0]
    return main

def main():
    st.set_page_config(page_title="Report Auto-generated Tool", page_icon="📊", layout="wide")
    st.markdown(
        "<h1 style='text-align: center; color: #4F8BF9;'>📊 Report Auto-generated Tool</h1>",
        unsafe_allow_html=True
    )
    st.markdown("---")

    if "template_file" not in st.session_state:
        st.session_state["template_file"] = None

    # 步骤1：上传数据文件
    with st.container():
        st.markdown("<h2 style='font-family:微软雅黑,Arial,sans-serif;font-weight:600;color:#4F8BF9;'>📁 步骤1：上传数据文件</h2>", unsafe_allow_html=True)
        col1, col2, col3, col4 = st.columns([1, 1, 1, 1])
        with col1:
            dat_file = st.file_uploader("上传 .dat 文件", type=["dat"])
            st.caption("DAT文件（测试覆盖数据）")
        with col2:
            csv_file = st.file_uploader("上传 .csv 文件", type=["csv"])
            st.caption("CSV文件（原始BOM/对照）")
        with col3:
            dcl_file = st.file_uploader("上传 .dcl 或 .csv 文件", type=["dcl", "csv"])
            st.caption("DCL文件或CSV文件（测试结果）")
        with col4:
            template_file = st.file_uploader("上传 Excel 模板（.xlsx）", type=["xlsx"])
            st.caption("模板文件（导出格式）")
            if template_file is not None:
                st.session_state["template_file"] = template_file

    st.markdown("---")
    progress = st.progress(0, text="等待开始...")

    if st.button("🚀 开始处理文件", use_container_width=True):
        progress.progress(5, text="校验文件...")
        if not dcl_file or not st.session_state["template_file"]:
            st.warning("请上传 .dcl 文件和 Excel 模板。")
            progress.progress(0, text="等待开始...")
            return

        # dcl
        progress.progress(20, text="解析 .dcl 文件...")
        header_data, component_df = DclHandler(dcl_file).process_dcl()

        # dat
        progress.progress(40, text="解析 .dat 文件...")
        dat_data = DatHandler(dat_file).process_dat() if dat_file else None

        # csv
        progress.progress(60, text="解析 .csv 文件...")
        csv_df = CsvHandler().process_csv(csv_file) if csv_file else pd.DataFrame()
        if csv_file and csv_df.empty:
            st.error("CSV文件解析失败，请将当前CSV文件另存为CSV UTF8格式。")
            return
        if not csv_df.empty and "Reference" in csv_df.columns and "Description" in csv_df.columns:
            ref_to_desc = build_ref_to_desc(csv_df)
        else:
            ref_to_desc = {}

        st.success("所有文件处理成功！")
        progress.progress(70, text="数据预览与编辑...")

        # 步骤2：Test Result 预览与编辑
        with st.expander("📝 步骤2：Test Result 预览与编辑", expanded=True):
            st.markdown("<h2 style='font-family:微软雅黑,Arial,sans-serif;font-weight:600;color:#4F8BF9;'>📝 步骤2：Test Result 预览与编辑</h2>", unsafe_allow_html=True)
            edited_dcl = st.data_editor(
                component_df,
                num_rows="dynamic",
                key="dcl_editor",
                use_container_width=True
            )

        # 步骤3：PartsCoverage 预览与编辑
        with st.expander("📝 步骤3：PartsCoverage 预览与编辑", expanded=True):
            st.markdown("<h2 style='font-family:微软雅黑,Arial,sans-serif;font-weight:600;color:#4F8BF9;'>📝 步骤3：PartsCoverage 预览与编辑</h2>", unsafe_allow_html=True)
            if dat_data and not dat_data["data"].empty:
                # 添加Description列（支持区间和多编号Reference）
                dat_data["data"]["Description"] = dat_data["data"]["Components"].apply(lambda x: get_description(x, ref_to_desc))
                # 分离Description为空和非空的行
                df_with_desc = dat_data["data"][dat_data["data"]["Description"].notna() & (dat_data["data"]["Description"] != "")]
                df_no_desc = dat_data["data"][dat_data["data"]["Description"].isna() | (dat_data["data"]["Description"] == "")]
                
                # 生成Remark列
                df_with_desc["Remark"] = df_with_desc.apply(gen_remark, axis=1)

                # 去重逻辑：同主编号优先保留没有/的，如果都有/，优先保留没有-的，都没有-的选字符串长度最短的
                if not df_with_desc.empty:
                    df_with_desc["main_comp"] = df_with_desc["Components"].apply(extract_main_comp)
                    df_with_desc["has_slash"] = df_with_desc["Components"].apply(lambda x: "/" in str(x))
                    df_with_desc["has_dash"] = df_with_desc["Components"].apply(lambda x: "-" in str(x))
                    df_with_desc["comp_len"] = df_with_desc["Components"].apply(lambda x: len(str(x)))
                    idx = (
                        df_with_desc
                        .sort_values(
                            ["main_comp", "has_slash", "has_dash", "comp_len"],
                            ascending=[True, True, True, True]
                        )
                        .groupby("main_comp", as_index=False)
                        .head(1)
                        .index
                    )
                    # 保留唯一主编号的行
                    df_unique = df_with_desc.loc[idx].copy()
                    # 其余为重复项
                    df_dup = df_with_desc.drop(idx)
                    # 清理辅助列
                    df_unique = df_unique.drop(columns=["main_comp", "has_slash", "has_dash", "comp_len"])
                    df_dup = df_dup.drop(columns=["main_comp", "has_slash", "has_dash", "comp_len"])
                else:
                    df_unique = df_with_desc
                    df_dup = pd.DataFrame()
               
                # 保证No.列自上而下递增
                if "No." in df_unique.columns:
                    df_unique["No."] = range(1, len(df_unique) + 1)

                # 只显示有Description且唯一的行在PartsCoverage
                show_cols = list(df_unique.columns)
                if "Description" not in show_cols:
                    show_cols.append("Description")
                if "Remark" not in show_cols:
                    show_cols.append("Remark")
                edited_dat = st.data_editor(
                    df_unique[show_cols],
                    num_rows="dynamic",
                    key="dat_editor",
                    use_container_width=True
                )
                # 统计Testable字段（只统计唯一主编号的行，不含重复项）
                testable_counts = df_unique["Testable"].value_counts()
                y_count = testable_counts.get("Y", 0)
                n_count = testable_counts.get("N", 0)
                l_count = testable_counts.get("L", 0)
                total = y_count + n_count + l_count
                coverage = (y_count + l_count) / total if total > 0 else 0

                st.info(
                    f"**Testable统计：** Y = {y_count}，N = {n_count}，L = {l_count}  \n"
                    f"**覆盖率 (Y+L)/(Y+L+N)：** {coverage:.2%}"
                )
                # 显示重复项
                if not df_dup.empty:
                    st.markdown("**Components重复项（仅显示不导出）**")
                    def highlight_dup(s):
                        return ['background-color: #d9ead3; text-align: center'] * len(s)
                    st.dataframe(
                        df_dup.style.apply(highlight_dup, axis=1).set_properties(**{'text-align': 'center'}),
                        use_container_width=True,
                        height=200
                    )

                
            else:
                edited_dat = pd.DataFrame()
                df_no_desc = pd.DataFrame()
                st.info("未检测到dat文件内容")
                coverage = 0

            # 显示/NC行并标红且居中
            if dat_data and not dat_data["nc_data"].empty:
                st.markdown("**/NC Components 信息（仅显示不导出）**")
                def highlight_nc(s):
                    return ['background-color: #ffcccc; text-align: center'] * len(s)
                st.dataframe(
                    dat_data["nc_data"].style.apply(highlight_nc, axis=1).set_properties(**{'text-align': 'center'}),
                    use_container_width=True,
                    height=200
                )

            # 显示Description为空的行，样式与/NC一致
            if not df_no_desc.empty:
                st.markdown("**Description为空的 Components 信息（仅显示不导出）**")
                def highlight_no_desc(s):
                    return ['background-color: #fff2cc; text-align: center'] * len(s)
                st.dataframe(
                    df_no_desc.style.apply(highlight_no_desc, axis=1).set_properties(**{'text-align': 'center'}),
                    use_container_width=True,
                    height=200
                )

        progress.progress(85, text="生成Excel文件...")

        # 生成Excel
        excel_bytes = fill_template_excel(
            st.session_state["template_file"],
            edited_dcl,
            csv_df,
            header_data,
            {
                "data": edited_dat,  # 只导出有Description且唯一的部分
                "nc_data": dat_data["nc_data"] if dat_data else pd.DataFrame(),
                "board_name": dat_data.get("board_name", "") if dat_data else "",
                "test_time": dat_data.get("test_time", "") if dat_data else "",
                "coverage": coverage
            }
        )

        progress.progress(100, text="处理完成！可下载Excel文件。")
        st.balloons()

        st.markdown("---")
        st.markdown("<h2 style='font-family:微软雅黑,Arial,sans-serif;font-weight:600;color:#4F8BF9;'>📥 步骤4：下载处理后的Excel</h2>", unsafe_allow_html=True)
        st.download_button(
            "下载Excel",
            excel_bytes,
            file_name="processed_data.xlsx",
            use_container_width=True
        )

if __name__ == "__main__":
    main()
