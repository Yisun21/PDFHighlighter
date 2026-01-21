import streamlit as st
import fitz  # PyMuPDF
import tempfile
import os

# --- 页面配置 ---
st.set_page_config(page_title="PDF 自动高亮工具", page_icon="🖍️", layout="wide")

st.title("🖍️ 论文关键词自动高亮助手")
st.markdown("上传 PDF，输入关键词，一键生成高亮版文档。")

# --- 侧边栏配置 ---
with st.sidebar:
    st.header("🛠️ 配置面板")

    # 1. 上传文件
    uploaded_pdf = st.file_uploader("1. 上传 PDF 文件", type=["pdf"])

    # 2. 输入关键词
    st.subheader("2. 关键词设置")
    word_input = st.text_area(
        "输入单词库 (支持换行或逗号分隔)",
        height=150,
        placeholder="例如：\ndeep learning\nattention mechanism\ntransformer"
    )

    # 3. 选项
    st.subheader("3. 选项")
    highlight_color = st.color_picker("选择高亮颜色", "#FFFF00")

    process_btn = st.button("🚀 开始处理", type="primary")


# --- 辅助函数：将 Hex 颜色转为 RGB ---
def hex_to_rgb(hex_color):
    hex_color = hex_color.lstrip('#')
    return tuple(int(hex_color[i:i + 2], 16) / 255.0 for i in (0, 2, 4))


# --- 主逻辑 ---
if process_btn:
    if not uploaded_pdf:
        st.error("请先上传一个 PDF 文件！")
    elif not word_input.strip():
        st.error("请输入至少一个关键词！")
    else:
        # 处理关键词
        raw_words = word_input.replace('\n', ',').split(',')
        keywords = [w.strip() for w in raw_words if w.strip()]

        if not keywords:
            st.error("关键词列表为空。")
        else:
            try:
                with st.spinner(f"正在扫描 {len(keywords)} 个关键词..."):

                    # 保存上传文件
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_input:
                        tmp_input.write(uploaded_pdf.getvalue())
                        tmp_input_path = tmp_input.name

                    # 打开 PDF
                    doc = fitz.open(tmp_input_path)
                    total_matches = 0
                    rgb_color = hex_to_rgb(highlight_color)

                    # 逐页搜索
                    progress_bar = st.progress(0)
                    for page_num, page in enumerate(doc):
                        progress_bar.progress((page_num + 1) / len(doc))

                        for word in keywords:
                            # 搜索单词坐标
                            quads = page.search_for(word, quads=True)
                            for quad in quads:
                                annot = page.add_highlight_annot(quad)
                                annot.set_colors(stroke=rgb_color)
                                annot.update()
                                total_matches += 1

                    # 保存结果
                    output_path = tmp_input_path.replace(".pdf", "_highlighted.pdf")
                    doc.save(output_path)
                    doc.close()

                    st.success(f"✅ 处理完成！共高亮 **{total_matches}** 处。")

                    with open(output_path, "rb") as file:
                        st.download_button(
                            label="📥 下载处理后的 PDF",
                            data=file,
                            file_name=f"highlighted_{uploaded_pdf.name}",
                            mime="application/pdf"
                        )

                    os.unlink(tmp_input_path)

            except Exception as e:
                st.error(f"发生错误: {e}")