import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
import tempfile
import os
import gc
import nltk
import base64
from nltk.stem import SnowballStemmer

# --- 引入专用的 PDF 预览库 ---
try:
    from streamlit_pdf_viewer import pdf_viewer
except ImportError:
    st.error("请先安装预览库：pip install streamlit-pdf-viewer")

# --- 页面配置 ---
st.set_page_config(page_title="PDF 智能词库高亮工具", page_icon="📚", layout="wide")

# --- NLTK 初始化 ---
try:
    stemmer = SnowballStemmer("english")
except:
    nltk.download('snowball_data')
    stemmer = SnowballStemmer("english")

# --- Session State 初始化 ---
if 'word_libraries' not in st.session_state:
    st.session_state['word_libraries'] = {}
if 'opacity_value' not in st.session_state:
    st.session_state['opacity_value'] = 0.20
# 存储生成结果的状态
if 'processed_pdf_data' not in st.session_state:
    st.session_state['processed_pdf_data'] = None
if 'processed_file_name' not in st.session_state:
    st.session_state['processed_file_name'] = ""

# 页码控制的状态变量初始化
if 'p_start' not in st.session_state: st.session_state['p_start'] = 1
if 'p_end' not in st.session_state: st.session_state['p_end'] = 1
if 'p_all' not in st.session_state: st.session_state['p_all'] = True


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

    st.subheader("2. 词库管理")

    # 2.1 Excel 上传
    uploaded_excels = st.file_uploader("📂 上传 Excel 词库（单词放在表格第一列）", type=['xlsx'], accept_multiple_files=True)
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

    # 2.2 【核心修改】手动词库管理器 (新建/编辑/查看)
    with st.expander("✍️ 手动添加/编辑词库", expanded=True):
        # 获取当前所有词库列表
        current_lib_names = list(st.session_state['word_libraries'].keys())
        # 下拉菜单：选择模式（新建 或 编辑现有）
        edit_mode_selection = st.selectbox(
            "选择操作",
            ["➕ 新建词库"] + current_lib_names,
            help="选择“新建”来创建新词库，或选择下方已有的词库名进行查看和修改。"
        )

        # 初始化编辑器变量
        target_lib_name = ""
        target_words_str = ""
        target_color = "#FFFF00"
        is_editing_existing = False

        if edit_mode_selection == "➕ 新建词库":
            target_lib_name = st.text_input("设置词库名称", placeholder="例如: 重点动词")
        else:
            is_editing_existing = True
            target_lib_name = edit_mode_selection  # 锁定名称
            st.info(f"正在编辑: **{target_lib_name}**")

            # 从 Session State 获取现有数据
            existing_data = st.session_state['word_libraries'][target_lib_name]
            # 将列表转回字符串显示
            target_words_str = ", ".join(existing_data['words'])
            target_color = existing_data.get('default_color', '#FFFF00')

        # 单词输入说明
        st.caption("📝 **单词输入格式说明**：")
        st.caption("请在下方文本框输入单词。支持**多行输入**或用**英文逗号**分隔。")
        st.caption("示例：\n`apple, banana` \n或\n`apple`\n`banana`")

        # 单词编辑区
        words_input = st.text_area(
            "单词列表内容",
            value=target_words_str,
            height=150,
            help="在这里编辑你的单词列表"
        )

        # 颜色选择
        color_input = st.color_picker("设置高亮颜色", value=target_color)

        # 按钮区域
        col_save, col_del = st.columns([1, 1])

        with col_save:
            if st.button("💾 保存/更新", type="primary"):
                if target_lib_name and words_input:
                    # 处理单词数据：支持换行和逗号混用
                    raw_words = words_input.replace('\n', ',').split(',')
                    # 去除空白和空项并去重
                    clean_words = sorted(list(set([w.strip() for w in raw_words if w.strip()])))

                    if clean_words:
                        st.session_state['word_libraries'][target_lib_name] = {
                            'words': clean_words,
                            'default_color': color_input
                        }
                        st.success(f"已保存! 共 {len(clean_words)} 词")
                        st.rerun()  # 刷新页面以更新选择列表
                    else:
                        st.warning("词库不能为空")
                else:
                    st.error("名称和内容不能为空")

        with col_del:
            if is_editing_existing:
                if st.button("🗑️ 删除此库"):
                    del st.session_state['word_libraries'][target_lib_name]
                    st.rerun()

    st.divider()

    st.subheader("3. 匹配与视觉")
    use_stemming = st.checkbox("启用智能词形匹配 (Stemming)", value=True)

    # --- 索引页高级设置 ---
    generate_index = st.checkbox("生成文末单词索引 (Index Page)", value=True)

    idx_col_count = 4
    idx_font_size = 10
    index_target_libs = []
    show_variants = False

    if generate_index:
        if use_stemming:
            show_variants = st.checkbox("在索引中显示文内单词变体 (例如: run -> running, ran)", value=True)
        else:
            show_variants = False

        default_col_index = 1 if show_variants else 3

        col1, col2 = st.columns(2)
        with col1:
            idx_col_count = st.selectbox("排版列数", [1, 2, 3, 4], index=default_col_index)
        with col2:
            idx_font_size = st.number_input("索引字号", min_value=8, max_value=16, value=10, step=1)

        available_libs = list(st.session_state['word_libraries'].keys())
        st.caption("选择要包含在索引页中的词库：")
        index_target_libs = st.multiselect(
            "索引词库选择",
            options=available_libs,
            default=available_libs,
            label_visibility="collapsed"
        )

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
        selected = st.multiselect("选择高亮词库", all_libs, default=all_libs)

        if selected:
            for name in selected:
                col1, col2 = st.columns([3, 1])
                with col1:
                    count = len(st.session_state['word_libraries'][name]['words'])
                    st.caption(f"**{name}** ({count} 词)")
                with col2:
                    # 这里的颜色选择器会显示该词库当前的颜色（可能是手动编辑时保存的）
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
        st.session_state['processed_pdf_data'] = None
        st.session_state['p_start'] = 1
        st.session_state['p_end'] = 1
        st.session_state['p_all'] = True
        st.cache_data.clear()
        st.rerun()

# --- 主界面 ---
st.title("📚 PDF 智能词库高亮工具")

if use_stemming:
    st.success("✨ 智能模式已开启：将自动忽略单词的时态、复数和变形。")
else:
    st.info("🔒 精确模式：仅匹配完全一致的单词。")

st.markdown(
    "Tip：**首次**出现的单词使用**深色**，**重复**出现的单词自动按**透明度**变浅；选择生成文末单词索引，将在文末附上**高亮单词列表**。")

# --- 处理逻辑 ---
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

        global_seen_items = {name: set() for name in final_configs}
        index_data_by_lib = {name: {} for name in final_configs}

        # --- 核心循环 ---
        for i, page in enumerate(doc):
            if i % 5 == 0:
                progress_bar.progress((i + 1) / total_pages)
                status_text.text(f"正在分析第 {i + 1} / {total_pages} 页...")

            page_words = page.get_text("words")

            for w_info in page_words:
                current_text = w_info[4]
                current_text_lower = current_text.lower()
                current_rect = fitz.Rect(w_info[0], w_info[1], w_info[2], w_info[3])
                current_stem = stemmer.stem(current_text_lower) if use_stemming else None

                for lib_name, p_cfg in processed_configs.items():
                    matched = False
                    match_key = None
                    origin_word = None

                    if use_stemming:
                        if current_stem in p_cfg['singles_stems']:
                            matched = True
                            match_key = current_stem
                            origin_word = p_cfg['stem_map'].get(current_stem)
                    else:
                        if current_text_lower in p_cfg['singles_exact']:
                            matched = True
                            match_key = current_text_lower
                            origin_word = p_cfg['exact_map'].get(current_text_lower)

                    if matched:
                        if match_key not in global_seen_items[lib_name]:
                            use_color = p_cfg['base_color']
                            global_seen_items[lib_name].add(match_key)
                        else:
                            use_color = p_cfg['light_color']

                        if origin_word:
                            if origin_word not in index_data_by_lib[lib_name]:
                                index_data_by_lib[lib_name][origin_word] = set()
                            index_data_by_lib[lib_name][origin_word].add(current_text)

                        annot = page.add_highlight_annot(current_rect)
                        annot.set_colors(stroke=use_color)
                        annot.update()
                        total_stats[lib_name] += 1

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

                            if phrase not in index_data_by_lib[lib_name]:
                                index_data_by_lib[lib_name][phrase] = set()
                            index_data_by_lib[lib_name][phrase].add(phrase)

                            annot = page.add_highlight_annot(quad)
                            annot.set_colors(stroke=use_color)
                            annot.update()
                            total_stats[lib_name] += 1

        # --- 索引生成 ---
        if generate_index:
            final_index_data = {k: v for k, v in index_data_by_lib.items() if k in index_target_libs}
            has_any_words = any(len(words_dict) > 0 for words_dict in final_index_data.values())

            if has_any_words:
                status_text.text(f"📄 正在排版索引页...")
                idx_page = doc.new_page()
                page_width = idx_page.rect.width
                page_height = idx_page.rect.height

                margin_x, margin_y = 40, 50
                col_gap = 15
                col_count = idx_col_count
                col_width = (page_width - 2 * margin_x - (col_count - 1) * col_gap) / col_count

                line_height = idx_font_size * 1.5
                header_height = idx_font_size * 2.0
                title_font_size = idx_font_size + 8
                lib_title_font_size = idx_font_size + 2
                var_font_size = max(6, idx_font_size - 2)

                avg_char_width = idx_font_size * 0.55
                truncation_limit = int(col_width / avg_char_width) - 2
                if truncation_limit < 5: truncation_limit = 5

                var_avg_char_width = var_font_size * 0.55
                var_truncation_limit = int(col_width / var_avg_char_width) - 4

                current_col = 0
                current_y = margin_y

                idx_page.insert_text((margin_x, 30), "Index of Words", fontsize=title_font_size, color=(0, 0, 0))

                for lib_name, words_dict in final_index_data.items():
                    if not words_dict: continue
                    sorted_origins = sorted(list(words_dict.keys()), key=str.lower)
                    lib_color = final_configs[lib_name]['rgb']

                    needed_height = header_height + line_height
                    if current_y + needed_height > page_height - margin_y:
                        current_col += 1
                        current_y = margin_y
                        if current_col >= col_count:
                            idx_page = doc.new_page()
                            current_col = 0
                    current_x = margin_x + current_col * (col_width + col_gap)

                    idx_page.insert_text((current_x, current_y), f"■ {lib_name}", fontsize=lib_title_font_size,
                                         color=lib_color)
                    current_y += header_height

                    for origin_word in sorted_origins:
                        display_variations = []
                        if show_variants:
                            found_variations = words_dict[origin_word]
                            display_variations = [v for v in found_variations if v.lower() != origin_word.lower()]
                            display_variations = sorted(list(set(display_variations)))

                        var_lines = []
                        if display_variations:
                            current_var_line = "("
                            for i, var in enumerate(display_variations):
                                separator = ", " if i > 0 else ""
                                if len(current_var_line + separator + var) > var_truncation_limit:
                                    if i > 0: current_var_line += ","
                                    var_lines.append(current_var_line)
                                    current_var_line = "  " + var
                                else:
                                    current_var_line += separator + var
                            current_var_line += ")"
                            var_lines.append(current_var_line)

                        item_height = line_height
                        if var_lines:
                            item_height += len(var_lines) * line_height

                        if current_y + item_height > page_height - margin_y:
                            current_col += 1
                            current_y = margin_y
                            if current_col >= col_count:
                                idx_page = doc.new_page()
                                current_col = 0
                            current_x = margin_x + current_col * (col_width + col_gap)

                        display_word = origin_word if len(origin_word) < truncation_limit else origin_word[
                                                                                               :truncation_limit] + "..."
                        idx_page.insert_text((current_x, current_y), f"  {display_word}", fontsize=idx_font_size,
                                             color=(0.2, 0.2, 0.2))
                        current_y += line_height

                        for v_line in var_lines:
                            idx_page.insert_text((current_x + 10, current_y), v_line, fontsize=var_font_size,
                                                 color=(0.5, 0.5, 0.5))
                            current_y += line_height

                    current_y += line_height / 2

        status_text.text("💾 正在保存结果...")
        output_path = tmp_input_path.replace(".pdf", "_highlighted_index.pdf")
        doc.save(output_path, garbage=4, deflate=True)
        doc.close()

        # 将结果存入 Session State
        with open(output_path, "rb") as file:
            st.session_state['processed_pdf_data'] = file.read()
            st.session_state['processed_file_name'] = f"Highlight_{uploaded_pdf.name}"

        # 重置页码状态
        temp_doc = fitz.open(stream=st.session_state['processed_pdf_data'], filetype="pdf")
        new_total_pages = len(temp_doc)
        temp_doc.close()

        st.session_state['p_start'] = 1
        st.session_state['p_end'] = new_total_pages
        st.session_state['p_all'] = True

        progress_bar.progress(100)
        status_text.text("✅ 完成！")

        os.unlink(tmp_input_path)
        os.unlink(output_path)
        gc.collect()

    except Exception as e:
        st.error(f"出错: {e}")

elif process_btn:
    st.error("请检查配置。")

# --- 结果显示区域 (独立渲染) ---
if st.session_state['processed_pdf_data'] is not None:
    st.divider()
    st.subheader("📂 结果区域")

    doc_result = fitz.open(stream=st.session_state['processed_pdf_data'], filetype="pdf")
    total_result_pages = len(doc_result)


    # 回调函数
    def on_toggle_all():
        if st.session_state['p_all']:
            st.session_state['p_start'] = 1
            st.session_state['p_end'] = total_result_pages


    def on_page_change():
        if st.session_state['p_start'] == 1 and st.session_state['p_end'] == total_result_pages:
            st.session_state['p_all'] = True
        else:
            st.session_state['p_all'] = False


    st.caption("选择预览和下载的页面范围：")
    col_p1, col_p2, col_opt = st.columns([1, 1, 2])

    with col_opt:
        st.write("")
        st.checkbox("🔄 全部预览 (默认所有页)", key='p_all', on_change=on_toggle_all)

        only_dl_preview = False
        if not st.session_state['p_all']:
            only_dl_preview = st.checkbox("⬇️ 仅下载上方选中的预览页数", value=False)

    with col_p1:
        st.number_input(
            "起始页",
            min_value=1,
            max_value=total_result_pages,
            step=1,
            key='p_start',
            on_change=on_page_change
        )
    with col_p2:
        st.number_input(
            "结束页",
            min_value=st.session_state['p_start'],
            max_value=total_result_pages,
            step=1,
            key='p_end',
            on_change=on_page_change
        )

    st.divider()

    # 动态切片逻辑
    target_pdf_data = st.session_state['processed_pdf_data']
    start_page_val = st.session_state['p_start']
    end_page_val = st.session_state['p_end']

    if start_page_val != 1 or end_page_val != total_result_pages:
        doc_slice = fitz.open()
        doc_slice.insert_pdf(doc_result, from_page=start_page_val - 1, to_page=end_page_val - 1)
        target_pdf_data = doc_slice.tobytes()
        doc_slice.close()

    doc_result.close()

    if only_dl_preview and not st.session_state['p_all']:
        download_data = target_pdf_data
        download_name = "Highlight_preview_" + uploaded_pdf.name
    else:
        download_data = st.session_state['processed_pdf_data']
        download_name = st.session_state['processed_file_name']

    col_dl, col_preview = st.columns([1, 4])

    with col_dl:
        st.download_button(
            "📥 下载结果 PDF",
            data=download_data,
            file_name=download_name,
            mime="application/pdf",
            type="primary"
        )

    with col_preview:
        if st.checkbox("👀 在线预览结果 PDF (展开/收起)", value=False):
            try:
                pdf_viewer(input=target_pdf_data, width=800)
            except Exception as e:
                st.error(f"预览加载失败: {e}")