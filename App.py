import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
import tempfile
import os
import base64

# --- 页面配置 ---
st.set_page_config(page_title="PDF 高亮 Pro 版", page_icon="🖍️", layout="wide")

# --- 初始化 Session State (用于保存历史记录) ---
if 'history' not in st.session_state:
    st.session_state['history'] = []  # 存储格式: [{'name': '时间戳/文件名', 'words': ['word1', 'word2']}]

if 'current_keywords' not in st.session_state:
    st.session_state['current_keywords'] = ""


# --- 辅助函数：PDF 预览生成器 ---
def display_pdf(file_path):
    """读取 PDF 文件并转换为 HTML iframe 以便在浏览器中预览"""
    with open(file_path, "rb") as f:
        base64_pdf = base64.b64encode(f.read()).decode('utf-8')
    pdf_display = f'<iframe src="data:application/pdf;base64,{base64_pdf}" width="100%" height="800" type="application/pdf"></iframe>'
    st.markdown(pdf_display, unsafe_allow_html=True)


# --- 辅助函数：颜色转换 ---
def hex_to_rgb(hex_color):
    hex_color = hex_color.lstrip('#')
    return tuple(int(hex_color[i:i + 2], 16) / 255.0 for i in (0, 2, 4))


# --- 侧边栏 UI ---
with st.sidebar:
    st.title("🛠️ 设置面板")

    st.subheader("1. 文件上传")
    uploaded_pdf = st.file_uploader("上传 PDF 论文", type=["pdf"])

    st.subheader("2. 词库来源")

    # 创建标签页：手动输入 vs Excel导入 vs 历史记录
    tab1, tab2, tab3 = st.tabs(["📝 手动", "📊 Excel", "clock 历史"])

    keywords_to_process = []

    # --- Tab 1: 手动输入 ---
    with tab1:
        text_input = st.text_area("输入单词 (逗号/换行分隔)",
                                  value=st.session_state['current_keywords'],
                                  height=150,
                                  key="text_area_input")
        if text_input:
            keywords_to_process = [w.strip() for w in text_input.replace('\n', ',').split(',') if w.strip()]

    # --- Tab 2: Excel 导入 ---
    with tab2:
        uploaded_excel = st.file_uploader("上传 Excel (.xlsx)", type=['xlsx'])
        if uploaded_excel:
            try:
                # 读取 Excel 第一列
                df = pd.read_excel(uploaded_excel)
                # 假设单词在第一列，转为字符串并去重
                excel_words = df.iloc[:, 0].dropna().astype(str).unique().tolist()
                st.info(f"成功读取 {len(excel_words)} 个单词")

                # 这里的按钮用于确认将 Excel 内容覆盖到当前处理列表
                if st.button("使用此 Excel 词库"):
                    st.session_state['current_keywords'] = ", ".join(excel_words)
                    keywords_to_process = excel_words
                    # 自动存入历史
                    st.session_state['history'].append({
                        'name': f"Excel: {uploaded_excel.name}",
                        'words': excel_words
                    })
                    st.rerun()  # 刷新页面以更新手动输入框
            except Exception as e:
                st.error(f"Excel 读取失败: {e}")

    # --- Tab 3: 历史记录 (本次会话) ---
    with tab3:
        if not st.session_state['history']:
            st.caption("暂无历史记录")
        else:
            # 下拉框选择历史
            history_names = [h['name'] for h in st.session_state['history'][::-1]]  # 倒序显示最新的
            selected_history = st.selectbox("选择历史词库", history_names)

            if st.button("加载历史词库"):
                # 找到对应的数据
                for h in st.session_state['history']:
                    if h['name'] == selected_history:
                        st.session_state['current_keywords'] = ", ".join(h['words'])
                        st.rerun()

    st.subheader("3. 选项")
    highlight_color = st.color_picker("高亮颜色", "#FFFF00")

    # 确认最终使用的关键词列表
    # 优先使用 text_input 的内容 (因为它可能被 Excel 或 历史记录 填充了)
    final_keywords = [w.strip() for w in text_input.replace('\n', ',').split(',') if w.strip()]

    st.markdown("---")
    process_btn = st.button("🚀 开始高亮处理", type="primary", use_container_width=True)

# --- 主界面 UI ---
st.title("🖍️ PDF 论文关键词高亮 Pro")

if not uploaded_pdf:
    st.info("👈 请先在左侧侧边栏上传 PDF 文件并设置词库。")
    # 展示一个空的占位符或说明
    st.markdown("""
    **功能更新说明：**
    - ✅ 支持 Excel 批量导入单词
    - ✅ 支持 PDF 在线预览
    - ✅ 支持会话级历史记录回溯
    """)

if process_btn and uploaded_pdf and final_keywords:

    # 将当前使用的词库也存入历史 (如果还没存过)
    current_combo_name = f"手动输入 ({len(final_keywords)}词)"
    # 简单的去重判断
    if not any(h['name'] == current_combo_name for h in st.session_state['history']):
        st.session_state['history'].append({'name': current_combo_name, 'words': final_keywords})

    col1, col2 = st.columns([1, 1])

    with st.spinner("正在逐页扫描文档..."):
        try:
            # 1. 保存上传的文件
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_input:
                tmp_input.write(uploaded_pdf.getvalue())
                tmp_input_path = tmp_input.name

            # 2. 打开 PDF
            doc = fitz.open(tmp_input_path)
            total_matches = 0
            rgb_color = hex_to_rgb(highlight_color)

            # 3. 处理每一页
            progress_bar = st.progress(0)
            for i, page in enumerate(doc):
                progress_bar.progress((i + 1) / len(doc))
                for word in final_keywords:
                    quads = page.search_for(word, quads=True)
                    for quad in quads:
                        annot = page.add_highlight_annot(quad)
                        annot.set_colors(stroke=rgb_color)
                        annot.update()
                        total_matches += 1

            # 4. 保存结果
            output_path = tmp_input_path.replace(".pdf", "_highlighted.pdf")
            doc.save(output_path)
            doc.close()

            # 5. 结果展示区域
            st.success(f"✅ 处理完成！共发现 **{total_matches}** 处高亮。")

            # 下载按钮
            with open(output_path, "rb") as file:
                pdf_bytes = file.read()
                st.download_button(
                    label="📥 下载已标注 PDF",
                    data=pdf_bytes,
                    file_name=f"highlighted_{uploaded_pdf.name}",
                    mime="application/pdf"
                )

            st.markdown("---")
            st.subheader("📄 文件预览")
            # 调用预览函数
            display_pdf(output_path)

            # 清理
            os.unlink(tmp_input_path)
            os.unlink(output_path)

        except Exception as e:
            st.error(f"处理出错: {e}")

elif process_btn:
    if not uploaded_pdf:
        st.error("请上传 PDF！")
    elif not final_keywords:
        st.error("关键词列表不能为空！")