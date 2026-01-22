import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
import tempfile
import os
import gc
import nltk
from nltk.stem import SnowballStemmer

# --- 页面配置 ---
st.set_page_config(page_title="PDF 智能词库高亮工具", page_icon="📚", layout="wide")

# --- NLTK 初始化 ---
try:
    stemmer = SnowballStemmer("english")
except:
    nltk.download('snowball_data')
    stemmer = SnowballStemmer("english")


# --- 缓存函数 ---
@st.cache_data(ttl=3600)
def load_excel_data(file):
    try:
        df = pd.read_excel(file)
        return df.iloc[:, 0].dropna().astype(str).str.strip().unique().tolist()
    except Exception:
        return []


# --- 颜色处理函数 ---
def hex_to_rgb(hex_color):
    hex_color = hex_color.lstrip('#')
    return tuple(int(hex_color[i:i + 2], 16) / 255.0 for i in (0, 2, 4))


def get_lighter_color(rgb, factor):
    r, g, b = rgb
    new_r = r + (1 - r) * factor
    new_g = g + (1 - g) * factor
    new_b = b + (1 - b) * factor
    return (new_r, new_g, new_b)


# --- 初始化 Session State ---
if 'word_libraries' not in st.session_state:
    st.session_state['word_libraries'] = {}

if 'opacity_value' not in st.session_state:
    st.session_state['opacity_value'] = 0.20


# --- 回调函数 ---
def update_opacity_from_slider():
    st.session_state['opacity_value'] = st.session_state['slider_widget']


def update_opacity_from_input():
    st.session_state['opacity_value'] = st.session_state['input_widget']


# --- 侧边栏 UI ---
with st.sidebar:
    st.title("🌟 效果设置")

    st.subheader("1. 文件")
    uploaded_pdf = st.file_uploader("上传 PDF", type=["pdf"], label_visibility="collapsed")

    st.divider()

    st.subheader("2. 词库（Excel）")
    uploaded_excels = st.file_uploader("上传词库", type=['xlsx'], accept_multiple_files=True)

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
                    'default_color': '#FFFF00'
                }
                st.rerun()

    st.divider()

    st.subheader("3. 匹配与视觉")
    use_stemming = st.checkbox("启用智能词形匹配 (Stemming)", value=True)

    # 索引页选项
    generate_index = st.checkbox("在文末附上索引页 (3栏排版)", value=True)

    st.write("重复单词高亮透明度 (1.0=原色, 0.0=透明)")

    col_input, col_slider = st.columns([1, 2.5])

    with col_input:
        st.number_input(
            label="数值输入",
            label_visibility="collapsed",
            min_value=0.0,
            max_value=1.0,
            step=0.01,
            value=st.session_state['opacity_value'],
            key='input_widget',
            on_change=update_opacity_from_input,
            format="%.2f"
        )

    with col_slider:
        st.slider(
            label="滑块调节",
            label_visibility="collapsed",
            min_value=0.0,
            max_value=1.0,
            step=0.01,
            value=st.session_state['opacity_value'],
            key='slider_widget',
            on_change=update_opacity_from_slider,
            help="1.00 表示保持最深的原色，0.00 表示完全透明（白色）。"
        )

    repeat_opacity = st.session_state['opacity_value']

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
    process_btn = st.button("🚀 生成高亮文件", type="primary", use_container_width=True)
    if st.button("🗑️ 清除缓存"):
        st.session_state['word_libraries'] = {}
        st.cache_data.clear()
        st.rerun()

# --- 主界面 ---
st.title("📚 PDF 智能词库高亮工具")

if use_stemming:
    st.success("✨ 智能模式已开启：将自动忽略单词的时态、复数和变形。")
else:
    st.info("🔒 精确模式：仅匹配完全一致的单词。")

st.markdown("Tip：**首次**出现的单词使用**深色**，**重复**出现的单词自动按**透明度**变浅。")

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

        status_text.text("🔍 正在初始化...")

        # --- 预处理配置 ---
        processed_configs = {}
        whiteness_factor = 1.0 - repeat_opacity

        for name, config in final_configs.items():
            words_list = config['words']
            singles_stems = set()
            singles_exact = set()
            phrases = []

            # 反向映射字典
            stem_to_origin_map = {}
            exact_to_origin_map = {}

            for w in words_list:
                clean_w = w.strip()
                if " " in clean_w:
                    phrases.append(clean_w)
                else:
                    lower_w = clean_w.lower()
                    singles_exact.add(lower_w)
                    exact_to_origin_map[lower_w] = clean_w

                    if use_stemming:
                        stem_w = stemmer.stem(lower_w)
                        singles_stems.add(stem_w)
                        stem_to_origin_map[stem_w] = clean_w

            base_rgb = config['rgb']
            light_rgb = get_lighter_color(base_rgb, factor=whiteness_factor)

            processed_configs[name] = {
                'singles_stems': singles_stems,
                'singles_exact': singles_exact,
                'phrases': phrases,
                'base_color': base_rgb,
                'light_color': light_rgb,
                'stem_map': stem_to_origin_map,
                'exact_map': exact_to_origin_map
            }

        # --- 追踪记录器 ---
        global_seen_items = {name: set() for name in final_configs}

        # 按词库分类收集
        index_data_by_lib = {name: set() for name in final_configs}

        # --- 核心循环 ---
        for i, page in enumerate(doc):
            if i % 5 == 0:
                progress_bar.progress((i + 1) / total_pages)
                status_text.text(f"正在分析第 {i + 1} / {total_pages} 页...")

            # 1. 处理单个单词
            page_words = page.get_text("words")

            for w_info in page_words:
                current_text = w_info[4].lower()
                current_rect = fitz.Rect(w_info[0], w_info[1], w_info[2], w_info[3])
                current_stem = stemmer.stem(current_text) if use_stemming else None

                for lib_name, p_cfg in processed_configs.items():
                    matched = False
                    match_key = None
                    origin_word = None  # 用于索引

                    if use_stemming:
                        if current_stem in p_cfg['singles_stems']:
                            matched = True
                            match_key = current_stem
                            origin_word = p_cfg['stem_map'].get(current_stem)
                    else:
                        if current_text in p_cfg['singles_exact']:
                            matched = True
                            match_key = current_text
                            origin_word = p_cfg['exact_map'].get(current_text)

                    if matched:
                        if match_key not in global_seen_items[lib_name]:
                            use_color = p_cfg['base_color']
                            global_seen_items[lib_name].add(match_key)
                        else:
                            use_color = p_cfg['light_color']

                        if origin_word:
                            index_data_by_lib[lib_name].add(origin_word)

                        annot = page.add_highlight_annot(current_rect)
                        annot.set_colors(stroke=use_color)
                        annot.update()
                        total_stats[lib_name] += 1

            # 2. 处理短语
            for lib_name, p_cfg in processed_configs.items():
                for phrase in p_cfg['phrases']:
                    quads_list = page.search_for(phrase, quads=True)
                    if quads_list:
                        for quad in quads_list:
                            match_key = phrase.lower()

                            if match_key not in global_seen_items[lib_name]:
                                use_color = p_cfg['base_color']
                                global_seen_items[lib_name].add(match_key)
                            else:
                                use_color = p_cfg['light_color']

                            index_data_by_lib[lib_name].add(phrase)

                            annot = page.add_highlight_annot(quad)
                            annot.set_colors(stroke=use_color)
                            annot.update()
                            total_stats[lib_name] += 1

        # --- 生成3栏分类索引页逻辑 ---
        if generate_index:
            has_any_words = any(len(words) > 0 for words in index_data_by_lib.values())

            if has_any_words:
                status_text.text("📄 正在排版索引页 (3栏模式)...")

                # 创建新页面
                idx_page = doc.new_page()
                page_width = idx_page.rect.width
                page_height = idx_page.rect.height

                # 排版参数
                margin_x = 40
                margin_y = 50
                col_gap = 20
                col_count = 3
                col_width = (page_width - 2 * margin_x - (col_count - 1) * col_gap) / col_count

                current_col = 0
                current_y = margin_y
                line_height = 14
                header_height = 20

                idx_page.insert_text((margin_x, 30), "Index of Found Words", fontsize=16, color=(0, 0, 0))

                for lib_name, words_set in index_data_by_lib.items():
                    if not words_set:
                        continue

                    sorted_words = sorted(list(words_set), key=str.lower)

                    # --- 【修复点】这里改成了 ['rgb'] ---
                    lib_color = final_configs[lib_name]['rgb']

                    needed_height = header_height + line_height
                    if current_y + needed_height > page_height - margin_y:
                        current_col += 1
                        current_y = margin_y
                        if current_col >= col_count:
                            idx_page = doc.new_page()
                            current_col = 0

                    current_x = margin_x + current_col * (col_width + col_gap)

                    idx_page.insert_text((current_x, current_y), f"■ {lib_name}", fontsize=10, color=lib_color)
                    current_y += header_height

                    for word in sorted_words:
                        if current_y > page_height - margin_y:
                            current_col += 1
                            current_y = margin_y
                            if current_col >= col_count:
                                idx_page = doc.new_page()
                                current_col = 0

                            current_x = margin_x + current_col * (col_width + col_gap)

                        display_word = word if len(word) < 25 else word[:22] + "..."
                        idx_page.insert_text((current_x, current_y), f"  {display_word}", fontsize=8,
                                             color=(0.2, 0.2, 0.2))
                        current_y += line_height

                    current_y += line_height / 2

        # 保存与结束
        status_text.text("💾 正在渲染最终文件...")
        output_path = tmp_input_path.replace(".pdf", "_highlighted_index.pdf")
        doc.save(output_path, garbage=4, deflate=True)
        doc.close()

        progress_bar.progress(100)
        status_text.text("✅ 完成！")

        cols = st.columns(len(total_stats))
        for idx, (name, count) in enumerate(total_stats.items()):
            cols[idx].metric(label=name, value=count)

        with open(output_path, "rb") as file:
            st.download_button(
                "📥 下载结果 PDF",
                data=file,
                file_name=f"Highlight_Index_{uploaded_pdf.name}",
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