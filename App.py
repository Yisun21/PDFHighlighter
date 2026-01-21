import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
import tempfile
import os
import base64

# --- 页面配置 ---
st.set_page_config(page_title="PDF 多色高亮 Pro Max", page_icon="🎨", layout="wide")

# --- 初始化 Session State (核心数据存储) ---
# word_libraries 结构: {'词库名': {'words': ['word1', 'word2'], 'default_color': '#FFFF00'}}
if 'word_libraries' not in st.session_state:
    st.session_state['word_libraries'] = {}


# --- 辅助函数 ---
def display_pdf(file_path):
    """生成 PDF 预览"""
    with open(file_path, "rb") as f:
        base64_pdf = base64.b64encode(f.read()).decode('utf-8')
    pdf_display = f'<iframe src="data:application/pdf;base64,{base64_pdf}" width="100%" height="800" type="application/pdf"></iframe>'
    st.markdown(pdf_display, unsafe_allow_html=True)


def hex_to_rgb(hex_color):
    """Hex 颜色转 RGB (0-1)"""
    hex_color = hex_color.lstrip('#')
    return tuple(int(hex_color[i:i + 2], 16) / 255.0 for i in (0, 2, 4))


# --- 侧边栏 UI ---
with st.sidebar:
    st.title("🛠️ 设置面板")

    # 1. 文件上传
    st.subheader("1. 上传 PDF")
    uploaded_pdf = st.file_uploader("选择论文文件", type=["pdf"], label_visibility="collapsed")

    st.divider()

    # 2. 词库管理 (支持多文件上传)
    st.subheader("2. 导入词库 (Excel)")
    # accept_multiple_files=True 允许一次选多个文件
    uploaded_excels = st.file_uploader(
        "上传多个 Excel (.xlsx)",
        type=['xlsx'],
        accept_multiple_files=True
    )

    # 处理上传的 Excel
    if uploaded_excels:
        for excel_file in uploaded_excels:
            # 如果这个文件还没被加载过，才去读取
            if excel_file.name not in st.session_state['word_libraries']:
                try:
                    df = pd.read_excel(excel_file)
                    # 默认读取第一列，去重，转字符串
                    words = df.iloc[:, 0].dropna().astype(str).unique().tolist()
                    # 存入 Session State
                    st.session_state['word_libraries'][excel_file.name] = {
                        'words': words,
                        'default_color': '#FFFF00'  # 默认黄色
                    }
                    st.toast(f"✅ 已加载: {excel_file.name} ({len(words)}词)")
                except Exception as e:
                    st.error(f"{excel_file.name} 读取失败: {e}")

    # 手动添加词库的功能
    with st.expander("➕ 手动添加临时词库"):
        manual_name = st.text_input("给词库起个名", placeholder="例如: 重点词汇")
        manual_text = st.text_area("输入单词 (逗号或换行分隔)", height=100)
        if st.button("添加手动词库"):
            if manual_name and manual_text:
                words = [w.strip() for w in manual_text.replace('\n', ',').split(',') if w.strip()]
                st.session_state['word_libraries'][manual_name] = {
                    'words': words,
                    'default_color': '#00FF00'  # 手动默认绿色
                }
                st.success(f"已添加 {manual_name}")
                st.rerun()

    st.divider()

    # 3. 词库配置与颜色选择
    st.subheader("3. 启用与配色")

    if not st.session_state['word_libraries']:
        st.info("👈 请先上传 Excel 或手动添加词库")
        final_configs = {}
    else:
        # 多选框：选择要使用哪些词库
        all_libs = list(st.session_state['word_libraries'].keys())
        selected_lib_names = st.multiselect(
            "选择要使用的高亮词库",
            all_libs,
            default=all_libs
        )

        # 动态生成颜色选择器
        final_configs = {}  # 存储最终的配置: {'词库名': {'words': [], 'rgb': (1,1,0)}}

        if selected_lib_names:
            st.write("🎨 为每个词库设置颜色:")
            for name in selected_lib_names:
                col1, col2 = st.columns([3, 1])
                with col1:
                    st.caption(f"**{name}** ({len(st.session_state['word_libraries'][name]['words'])} 词)")
                with col2:
                    # 获取该词库之前的颜色设置，如果没有则用默认
                    current_hex = st.color_picker(
                        f"颜色-{name}",
                        st.session_state['word_libraries'][name]['default_color'],
                        key=f"picker_{name}",
                        label_visibility="collapsed"
                    )

                # 保存配置
                final_configs[name] = {
                    'words': st.session_state['word_libraries'][name]['words'],
                    'rgb': hex_to_rgb(current_hex)
                }

    st.divider()
    process_btn = st.button("🚀 开始多色高亮", type="primary", use_container_width=True)

    # 清空历史按钮
    if st.button("🗑️ 清空所有词库缓存"):
        st.session_state['word_libraries'] = {}
        st.rerun()

# --- 主界面 ---
st.title("🎨 PDF 多源词库高亮工具")

if not uploaded_pdf:
    st.info("请在左侧上传 PDF 并配置词库。")

if process_btn and uploaded_pdf and final_configs:
    col1, col2 = st.columns([1, 1])

    with st.spinner("正在进行多色图层渲染..."):
        try:
            # 1. 准备文件
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_input:
                tmp_input.write(uploaded_pdf.getvalue())
                tmp_input_path = tmp_input.name

            doc = fitz.open(tmp_input_path)
            total_stats = {name: 0 for name in final_configs}  # 统计每个词库高亮了多少个

            # 2. 核心处理循环
            progress_bar = st.progress(0)

            for i, page in enumerate(doc):
                progress_bar.progress((i + 1) / len(doc))

                # 针对每一页，遍历所有选中的词库
                for lib_name, config in final_configs.items():
                    words = config['words']
                    color_rgb = config['rgb']

                    for word in words:
                        # 搜索单词
                        quads = page.search_for(word, quads=True)

                        # 应用高亮
                        for quad in quads:
                            annot = page.add_highlight_annot(quad)
                            annot.set_colors(stroke=color_rgb)
                            annot.update()
                            total_stats[lib_name] += 1

            # 3. 保存与展示
            output_path = tmp_input_path.replace(".pdf", "_highlighted.pdf")
            doc.save(output_path)
            doc.close()

            # 显示统计信息
            st.success("✅ 处理完成！统计如下：")
            stat_cols = st.columns(len(total_stats))
            for idx, (name, count) in enumerate(total_stats.items()):
                # 为了防止列太多挤压，这里简单的用 container
                st.write(f"**{name}**: {count} 处")

            # 下载按钮
            with open(output_path, "rb") as file:
                st.download_button(
                    label="📥 下载多色标注版 PDF",
                    data=file,
                    file_name=f"MultiColor_{uploaded_pdf.name}",
                    mime="application/pdf"
                )

            st.divider()
            st.subheader("📄 效果预览")
            display_pdf(output_path)

            # 清理
            os.unlink(tmp_input_path)
            os.unlink(output_path)

        except Exception as e:
            st.error(f"处理过程中出错: {e}")

elif process_btn:
    if not uploaded_pdf:
        st.error("请先上传 PDF 文件！")
    elif not final_configs:
        st.error("请至少启用一个词库！")