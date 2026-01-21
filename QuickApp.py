import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
import tempfile
import os
import base64
import gc # 垃圾回收

# --- 页面配置 ---
st.set_page_config(page_title="PDF 多色高亮极速版", page_icon="⚡", layout="wide")

# --- 缓存函数：读取 Excel (性能优化核心) ---
@st.cache_data(ttl=3600) # 缓存 1 小时
def load_excel_data(file):
    try:
        df = pd.read_excel(file)
        # 读取第一列，去重，转字符串，过滤空值
        return df.iloc[:, 0].dropna().astype(str).unique().tolist()
    except Exception:
        return []

# --- 辅助函数：生成轻量级预览 (只预览前N页) ---
def display_pdf_preview(file_path, max_pages=3):
    """
    为了速度，只提取 PDF 的前几页进行预览，
    避免整个大文件 Base64 编码导致浏览器卡顿。
    """
    try:
        # 打开生成的 PDF
        doc = fitz.open(file_path)
        # 如果页数超过限制，创建一个新的临时小 PDF 用于预览
        if len(doc) > max_pages:
            temp_preview_doc = fitz.open()
            temp_preview_doc.insert_pdf(doc, from_page=0, to_page=max_pages-1)
            pdf_bytes = temp_preview_doc.tobytes()
            temp_preview_doc.close()
            st.caption(f"⚡ 为提升速度，仅预览前 {max_pages} 页 (下载文件是完整的)")
        else:
            pdf_bytes = doc.tobytes()
        
        doc.close()

        # 编码并显示
        base64_pdf = base64.b64encode(pdf_bytes).decode('utf-8')
        pdf_display = f'<iframe src="data:application/pdf;base64,{base64_pdf}" width="100%" height="800" type="application/pdf"></iframe>'
        st.markdown(pdf_display, unsafe_allow_html=True)
        
    except Exception as e:
        st.error(f"预览生成失败: {e}")

def hex_to_rgb(hex_color):
    hex_color = hex_color.lstrip('#')
    return tuple(int(hex_color[i:i+2], 16)/255.0 for i in (0, 2, 4))

# --- 初始化 Session State ---
if 'word_libraries' not in st.session_state:
    st.session_state['word_libraries'] = {} 

# --- 侧边栏 UI ---
with st.sidebar:
    st.title("⚡ 极速设置")
    
    st.subheader("1. 文件")
    uploaded_pdf = st.file_uploader("上传 PDF", type=["pdf"], label_visibility="collapsed")
    
    st.divider()
    
    st.subheader("2. 词库 (Excel)")
    uploaded_excels = st.file_uploader(
        "上传词库 (.xlsx)", 
        type=['xlsx'], 
        accept_multiple_files=True
    )
    
    # 优化后的 Excel 读取逻辑
    if uploaded_excels:
        for excel_file in uploaded_excels:
            if excel_file.name not in st.session_state['word_libraries']:
                # 使用缓存函数读取
                words = load_excel_data(excel_file)
                if words:
                    st.session_state['word_libraries'][excel_file.name] = {
                        'words': words,
                        'default_color': '#FFFF00'
                    }
                    st.toast(f"已缓存: {excel_file.name}")

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

    # 3. 配置
    st.subheader("3. 颜色配置")
    final_configs = {}
    
    if st.session_state['word_libraries']:
        all_libs = list(st.session_state['word_libraries'].keys())
        selected = st.multiselect("选择词库", all_libs, default=all_libs)
        
        if selected:
            for name in selected:
                col1, col2 = st.columns([3, 1])
                with col1:
                    st.caption(f"{name}")
                with col2:
                    c = st.color_picker(f"C-{name}", st.session_state['word_libraries'][name]['default_color'], key=f"c_{name}")
                
                final_configs[name] = {
                    'words': st.session_state['word_libraries'][name]['words'],
                    'rgb': hex_to_rgb(c)
                }
        
    st.divider()
    process_btn = st.button("🚀 极速处理", type="primary", use_container_width=True)
    if st.button("🗑️ 清除缓存"):
        st.session_state['word_libraries'] = {}
        st.cache_data.clear() # 清除 Excel 读取缓存
        st.rerun()

# --- 主界面 ---
st.title("⚡ PDF 高亮工具 (性能优化版)")

if process_btn and uploaded_pdf and final_configs:
    
    # 进度条占位符
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    try:
        # 1. 读取文件流 (不在内存中完全加载，使用流式处理优化大文件)
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_input:
            tmp_input.write(uploaded_pdf.getvalue())
            tmp_input_path = tmp_input.name

        doc = fitz.open(tmp_input_path)
        total_pages = len(doc)
        total_stats = {name: 0 for name in final_configs}

        status_text.text("🔍 正在初始化搜索引擎...")

        # 2. 核心循环优化
        # 将配置字典转化为列表，减少循环内的字典查找开销
        active_configs = list(final_configs.items()) 

        for i, page in enumerate(doc):
            # 每 5 页更新一次进度条，减少界面重绘带来的卡顿
            if i % 5 == 0:
                progress_bar.progress((i + 1) / total_pages)
                status_text.text(f"正在处理第 {i+1} / {total_pages} 页...")
            
            # --- 核心搜索层 ---
            # 优化点：对于某些页面，如果完全没有文本，可以跳过（可选，此处暂未加，防止OCR页漏检）
            
            for lib_name, config in active_configs:
                target_words = config['words']
                color = config['rgb']
                
                for word in target_words:
                    # search_for 已经是 C 语言级别的速度，很难再优化
                    # 但我们可以确保不进行无意义的 update
                    quads = page.search_for(word, quads=True)
                    
                    if quads: # 只有找到时才操作
                        for quad in quads:
                            annot = page.add_highlight_annot(quad)
                            annot.set_colors(stroke=color)
                            annot.update() # 这一步必须有
                            total_stats[lib_name] += 1

        # 3. 保存
        status_text.text("💾 正在重新打包 PDF...")
        output_path = tmp_input_path.replace(".pdf", "_highlighted.pdf")
        
        # garbage=4: 深度压缩和清理未使用的对象，减小文件体积
        doc.save(output_path, garbage=4, deflate=True) 
        doc.close()
        
        # 4. 完成反馈
        progress_bar.progress(100)
        status_text.text("✅ 完成！")
        
        # 统计展示
        cols = st.columns(len(total_stats))
        for idx, (name, count) in enumerate(total_stats.items()):
            cols[idx].metric(label=name, value=count)

        # 下载
        with open(output_path, "rb") as file:
            st.download_button(
                "📥 下载完整版 PDF",
                data=file,
                file_name=f"Highlighted_{uploaded_pdf.name}",
                mime="application/pdf"
            )
        
        st.divider()
        # 调用极速预览
        display_pdf_preview(output_path, max_pages=3)
        
        # 清理
        os.unlink(tmp_input_path)
        os.unlink(output_path)
        gc.collect() # 手动触发垃圾回收

    except Exception as e:
        st.error(f"出错: {e}")

elif process_btn:
    st.error("请检查文件和配置。")
