import streamlit as st
import pandas as pd
from core import process_all

st.set_page_config(page_title="PRER数据分析工具", layout="wide")

st.title("📊 PRER 文献数据自动分析系统")

# =========================
# 文件上传
# =========================
eff_file = st.file_uploader("上传有效性Excel", type=["xlsx"])
saf_file = st.file_uploader("上传安全性Excel", type=["xlsx"])

# =========================
# 配置项
# =========================
st.sidebar.header("⚙️ 分析配置")

enable_merge_eff = st.sidebar.checkbox("启用有效性合并计算", value=True)
enable_merge_saf = st.sidebar.checkbox("启用安全性汇总", value=True)

st.sidebar.markdown("### 输出控制")
download_format = st.sidebar.selectbox("下载格式", ["xlsx"])

# =========================
# 执行按钮
# =========================
if st.button("🚀 开始分析"):

    if not eff_file or not saf_file:
        st.error("请上传两个Excel文件")
        st.stop()

    with st.spinner("正在计算中..."):

        eff_results, saf_results = process_all(
            eff_file,
            saf_file,
            enable_merge_eff,
            enable_merge_saf
        )

    st.success("计算完成！")

    # =========================
    # 下载区
    # =========================
    st.subheader("📥 下载结果")

    for name, df in eff_results.items():
        st.download_button(
            label=f"下载有效性-{name}",
            data=df.to_excel(index=False, engine="openpyxl"),
            file_name=f"eff_{name}.xlsx"
        )

    for name, df in saf_results.items():
        st.download_button(
            label=f"下载安全性-{name}",
            data=df.to_excel(index=False, engine="openpyxl"),
            file_name=f"saf_{name}.xlsx"
        )
