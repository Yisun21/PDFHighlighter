import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
import tempfile
import os
import gc
import nltk
from nltk.stem import SnowballStemmer

# --- 页面配置 ---
st.set_page_config(page_title="PDF 智能词库匹配高亮工具", page_icon="📚", layout="wide")

# --- NLTK 初始化 ---
# 初始化英语词干提取器
stemmer = SnowballStemmer("english")


# --- 缓存函数 ---
@st.cache_data(ttl=3600)
def load_excel_data(file):
    try:
        df = pd.read_excel(file)
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
    st.title("📚 智能设置")

    st.subheader("1. 文件")
    uploaded_pdf = st.file_uploader("上传 PDF", type=["pdf"], label_visibility="collapsed")

    st.divider()

    st.subheader("2. 词库 (Excel)")
    uploaded_excels = st.file_uploader(
        "上传词库（单词放在Excel表格第一列） (.xlsx)",
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

    st.subheader("3. 匹配与颜色")

    # 新增：模糊匹配开关
    use_stemming = st.checkbox("启用智能词形匹配 (Stemming)", value=True,
                               help="勾选后，'run' 可以匹配 'running', 'ran', 'runner' 等")

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
    process_btn = st.button("🚀 开始智能处理", type="primary", use_container_width=True)
    if st.button("🗑️ 清除缓存"):
        st.session_state['word_libraries'] = {}
        st.cache_data.clear()
        st.rerun()

# --- 主界面 ---
st.title("📚 PDF 智能词库匹配高亮工具")
if use_stemming:
    st.success("✨ 智能模式已开启：将自动忽略单词的时态、复数和变形。")
else:
    st.info("🔒 精确模式：仅匹配完全一致的单词。")

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

        status_text.text("🔍 正在构建词根索引...")

        # --- 预处理：构建匹配字典 ---
        processed_configs = {}
        for name, config in final_configs.items():
            words_list = config['words']

            # 我们需要存储两个集合：
            # 1. singles_stems: 单个单词的词根集合 (用于智能匹配)
            # 2. singles_exact: 单个单词的原词集合 (用于精确匹配)
            # 3. phrases: 短语 (短语很难做词根匹配，通常保持原样搜索)

            singles_stems = set()
            singles_exact = set()
            phrases = []

            for w in words_list:
                clean_w = w.strip()
                if " " in clean_w:
                    phrases.append(clean_w)  # 短语走传统搜索
                else:
                    lower_w = clean_w.lower()
                    singles_exact.add(lower_w)
                    if use_stemming:
                        # 计算词根，例如 'computing' -> 'comput'
                        stem_w = stemmer.stem(lower_w)
                        singles_stems.add(stem_w)

            processed_configs[name] = {
                'singles_stems': singles_stems,
                'singles_exact': singles_exact,
                'phrases': phrases,
                'color': config['rgb']
            }

        # --- 核心循环 ---
        for i, page in enumerate(doc):
            if i % 5 == 0:
                progress_bar.progress((i + 1) / total_pages)
                status_text.text(f"正在分析第 {i + 1} / {total_pages} 页...")

            # 1. 处理单个单词 (智能/精确逻辑)
            page_words = page.get_text("words")  # 获取页面所有单词信息

            for w_info in page_words:
                current_text = w_info[4].lower()  # PDF中的单词
                current_rect = fitz.Rect(w_info[0], w_info[1], w_info[2], w_info[3])

                # 如果开启了智能匹配，我们计算当前单词的词根
                current_stem = stemmer.stem(current_text) if use_stemming else None

                for lib_name, p_cfg in processed_configs.items():
                    matched = False

                    if use_stemming:
                        # 智能模式：比较词根
                        if current_stem in p_cfg['singles_stems']:
                            matched = True
                    else:
                        # 精确模式：比较原词
                        if current_text in p_cfg['singles_exact']:
                            matched = True

                    if matched:
                        annot = page.add_highlight_annot(current_rect)
                        annot.set_colors(stroke=p_cfg['color'])
                        annot.update()
                        total_stats[lib_name] += 1

            # 2. 处理短语 (依然使用 search_for，短语通常不需要词形变化)
            for lib_name, p_cfg in processed_configs.items():
                for phrase in p_cfg['phrases']:
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

        with open(output_path, "rb") as file:
            st.download_button(
                "📥 下载结果 PDF",
                data=file,
                file_name=f"SmartMatch_{uploaded_pdf.name}",
                mime="application/pdf",
                type="primary"
            )

        os.unlink(tmp_input_path)
        os.unlink(output_path)
        gc.collect()

    except Exception as e:
        st.error(f"出错: {e}")

elif process_btn:
    st.error("请检查配置。")