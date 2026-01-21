import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
import tempfile
import os
import gc  # 垃圾回收

# --- 页面配置 ---
st.set_page_config(page_title="PDF 全词匹配高亮工具", page_icon="🎯", layout="wide")


# --- 缓存函数 ---
@st.cache_data(ttl=3600)
def load_excel_data(file):
    try:
        df = pd.read_excel(file)
        # 读取第一列，去重，转字符串，去除首尾空格
        return df.iloc[:, 0].dropna().astype(str).str.strip().unique().tolist()
    except Exception:
        return []


def hex_to_rgb(hex_color):
    hex_color = hex_color.lstrip('#')
    return tuple(int(hex_color[i:i + 2], 16) / 255.0 for i in (0, 2, 4))


# --- 初始化 Session State ---
if 'word_libraries' not in st.session_state:
    st.session_state['word_libraries'] = {}

# --- 侧边栏 UI ---
with st.sidebar:
    st.title("🎯 精准设置")

    st.subheader("1. 文件")
    uploaded_pdf = st.file_uploader("上传 PDF", type=["pdf"], label_visibility="collapsed")

    st.divider()

    st.subheader("2. 词库 (Excel)")
    uploaded_excels = st.file_uploader(
        "上传词库（单词放Excel表格第一列） (.xlsx)",
        type=['xlsx'],
        accept_multiple_files=True
    )

    if uploaded_excels:
        for excel_file in uploaded_excels:
            if excel_file.name not in st.session_state['word_libraries']:
                words = load_excel_data(excel_file)
                if words:
                    st.session_state['word_libraries'][excel_file.name] = {
                        'words': words,
                        'default_color': '#FFFF00'
                    }
                    st.toast(f"✅ 已缓存: {excel_file.name} (共 {len(words)} 词)")

    with st.expander("➕ 手动添加"):
        manual_name = st.text_input("词库名")
        manual_text = st.text_area("单词列表")
        if st.button("添加"):
            if manual_name and manual_text:
                words = [w.strip() for w in manual_text.replace('\n', ',').split(',') if w.strip()]
                st.session_state['word_libraries'][manual_name] = {
                    'words': words,
                    'default_color': '#00FF00'
                }
                st.rerun()

    st.divider()

    st.subheader("3. 颜色配置")
    final_configs = {}

    if st.session_state['word_libraries']:
        all_libs = list(st.session_state['word_libraries'].keys())
        selected = st.multiselect("选择词库", all_libs, default=all_libs)

        if selected:
            for name in selected:
                col1, col2 = st.columns([3, 1])
                with col1:
                    count = len(st.session_state['word_libraries'][name]['words'])
                    st.caption(f"**{name}** ({count} 词)")
                with col2:
                    c = st.color_picker(f"C-{name}", st.session_state['word_libraries'][name]['default_color'],
                                        key=f"c_{name}")

                final_configs[name] = {
                    'words': st.session_state['word_libraries'][name]['words'],
                    'rgb': hex_to_rgb(c)
                }

    st.divider()
    process_btn = st.button("🚀 开始精准匹配", type="primary", use_container_width=True)
    if st.button("🗑️ 清除缓存"):
        st.session_state['word_libraries'] = {}
        st.cache_data.clear()
        st.rerun()

# --- 主界面 ---
st.title("🎯 PDF 全词匹配高亮工具")
st.markdown("已启用 **Whole Word Matching** 模式：精确匹配单词，拒绝部分匹配。")

if process_btn and uploaded_pdf and final_configs:

    progress_bar = st.progress(0)
    status_text = st.empty()

    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_input:
            tmp_input.write(uploaded_pdf.getvalue())
            tmp_input_path = tmp_input.name

        doc = fitz.open(tmp_input_path)
        total_pages = len(doc)
        total_stats = {name: 0 for name in final_configs}

        status_text.text("🔍 正在初始化精准匹配引擎...")

        # --- 预处理词库：区分单词和短语 ---
        # 单词：用 get_text("words") 做全等匹配 (解决 cat 匹配 scatter)
        # 短语：用 search_for 做搜索匹配 (解决 Deep Learning 带空格问题)
        processed_configs = {}
        for name, config in final_configs.items():
            words_list = config['words']
            single_words = set()  # 用集合加速查找
            phrases = []

            for w in words_list:
                clean_w = w.strip()
                if " " in clean_w:  # 如果包含空格，视为短语
                    phrases.append(clean_w)
                else:
                    single_words.add(clean_w.lower())  # 转小写存入集合

            processed_configs[name] = {
                'singles': single_words,
                'phrases': phrases,
                'color': config['rgb']
            }

        # --- 核心循环 ---
        for i, page in enumerate(doc):
            if i % 5 == 0:
                progress_bar.progress((i + 1) / total_pages)
                status_text.text(f"正在分析第 {i + 1} / {total_pages} 页...")

            # 1. 处理所有“单个单词” (全词匹配逻辑)
            # 获取页面所有单词: (x0, y0, x1, y1, "word_string", ...)
            page_words = page.get_text("words")

            for w_info in page_words:
                # w_info[4] 是单词文本
                current_word_text = w_info[4].lower()
                current_word_rect = fitz.Rect(w_info[0], w_info[1], w_info[2], w_info[3])

                # 检查这个单词是否在我们的任何一个词库里
                for lib_name, p_cfg in processed_configs.items():
                    if current_word_text in p_cfg['singles']:
                        # 只有完全相等才高亮
                        annot = page.add_highlight_annot(current_word_rect)
                        annot.set_colors(stroke=p_cfg['color'])
                        annot.update()
                        total_stats[lib_name] += 1

            # 2. 处理“短语” (传统搜索逻辑，因为 get_text("words") 会把短语拆散)
            for lib_name, p_cfg in processed_configs.items():
                for phrase in p_cfg['phrases']:
                    # 短语依然使用 search_for，但通常短语不太容易出现误匹配
                    quads = page.search_for(phrase, quads=True)
                    if quads:
                        for quad in quads:
                            annot = page.add_highlight_annot(quad)
                            annot.set_colors(stroke=p_cfg['color'])
                            annot.update()
                            total_stats[lib_name] += 1

        # 保存
        status_text.text("💾 正在保存...")
        output_path = tmp_input_path.replace(".pdf", "_highlighted.pdf")
        doc.save(output_path, garbage=4, deflate=True)
        doc.close()

        progress_bar.progress(100)
        status_text.text("✅ 完成！")

        # 统计
        cols = st.columns(len(total_stats))
        for idx, (name, count) in enumerate(total_stats.items()):
            cols[idx].metric(label=name, value=count)

        # 仅显示下载按钮，无预览
        with open(output_path, "rb") as file:
            st.download_button(
                "📥 下载结果 PDF",
                data=file,
                file_name=f"WholeWord_{uploaded_pdf.name}",
                mime="application/pdf",
                type="primary"  # 醒目的按钮
            )

        os.unlink(tmp_input_path)
        os.unlink(output_path)
        gc.collect()

    except Exception as e:
        st.error(f"出错: {e}")

elif process_btn:
    st.error("请检查配置。")