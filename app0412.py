import streamlit as st
import pandas as pd
from io import BytesIO
from core import process_all

st.set_page_config(page_title="PRER数据分析系统", layout="wide")

# =========================
# 初始化状态
# =========================
if "eff_results" not in st.session_state:
    st.session_state.eff_results = None

if "saf_results" not in st.session_state:
    st.session_state.saf_results = None

if "has_result" not in st.session_state:
    st.session_state.has_result = False


# =========================
# 页面标题
# =========================
st.title("📊 PRER 文献数据自动分析系统")

# =========================
# 左下角版权
# =========================
st.markdown(
    """
    <style>
    .footer {
        position: fixed;
        left: 10px;
        bottom: 10px;
        color: gray;
        font-size: 12px;
    }
    </style>
    <div class="footer">
        CER中心——数据分析工具
    </div>
    """,
    unsafe_allow_html=True
)

# =========================
# 左右布局
# =========================
left, right = st.columns([1, 2])


# =========================
# Excel转Bytes
# =========================
def to_excel_bytes(results_dict):
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for name, df in results_dict.items():
            df.to_excel(writer, sheet_name=str(name)[:31], index=False)
    buffer.seek(0)
    return buffer


# =========================
# 左侧：输入区
# =========================
with left:

    st.header("📂 数据上传")

    eff_file = st.file_uploader("有效性数据", type=["xlsx"])
    saf_file = st.file_uploader("安全性数据", type=["xlsx"])

    st.divider()

    # =========================
    # 按钮区
    # =========================
    if st.button("🚀 开始分析", use_container_width=True):

        if eff_file is None or saf_file is None:
            st.error("请上传两个Excel文件")
        else:
            try:
                with st.spinner("正在分析中..."):

                    eff_results, saf_results = process_all(
                        eff_file,
                        saf_file
                    )

                    st.session_state.eff_results = eff_results
                    st.session_state.saf_results = saf_results
                    st.session_state.has_result = True

                st.success("✅ 分析完成")

            except Exception as e:
                st.error("❌ 运行失败")
                st.exception(e)

    # =========================
    # 重置按钮（新增）
    # =========================
    if st.session_state.has_result:
        if st.button("🔄 重新分析", use_container_width=True):
            st.session_state.eff_results = None
            st.session_state.saf_results = None
            st.session_state.has_result = False
            st.rerun()


# =========================
# 右侧：结果区
# =========================
with right:

    st.header("📊 分析结果")

    if not st.session_state.has_result:
        st.info("请上传数据并点击“开始分析”")
    else:
        st.success("📌 当前结果已生成")

        eff_results = st.session_state.eff_results
        saf_results = st.session_state.saf_results

        # =========================
        # 下载区
        # =========================
        st.subheader("📥 下载结果")

        eff_excel = to_excel_bytes(eff_results)
        saf_excel = to_excel_bytes(saf_results)

        col1, col2 = st.columns(2)

        with col1:
            st.markdown("**有效性结果** ✅")
            st.download_button(
                "⬇️ 下载",
                data=eff_excel,
                file_name="有效性_结果.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

        with col2:
            st.markdown("**安全性结果** ✅")
            st.download_button(
                "⬇️ 下载",
                data=saf_excel,
                file_name="安全性_结果.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

        st.divider()

        # =========================
        # 预览区
        # =========================
        st.subheader("📊 数据预览")

        tab1, tab2 = st.tabs(["有效性", "安全性"])

        with tab1:
            st.dataframe(list(eff_results.values())[0], use_container_width=True)

        with tab2:
            st.dataframe(list(saf_results.values())[0], use_container_width=True)
