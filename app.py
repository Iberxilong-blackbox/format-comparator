import streamlit as st
from pathlib import Path
import tempfile
import os
import shutil # For copying uploaded file if needed
import json # Added for debugging print

# Import modules from the win32com package
from docx_reader_win32 import DocxReaderWin32
from template_manager_win32 import TemplateManagerWin32
from format_comparator_win32 import FormatComparatorWin32, add_comments_to_document_static
from report_generator import (
    ReportGenerator,
    display_report_summary,
    display_report_charts, 
    display_report_details_table,
    DEFAULT_ERROR_WEIGHTS,
    DEFAULT_PENALTY_TIERS,
    DEFAULT_ACCELERATION_THRESHOLDS
)

# --- Page Configuration ---
st.set_page_config(page_title="Word Format Checker (win32com)", layout="wide", initial_sidebar_state="collapsed")

# --- Paths Configuration ---
# Assuming app.py is in win32com/, so user_files is a sibling to app.py's parent or at a predefined location
# For simplicity, let's assume user_files is at the same level as the win32com directory,
# or the TemplateManagerWin32 handles its path relative to its own location.
WIN32COM_DIR = Path(__file__).parent
USER_FILES_DIR = WIN32COM_DIR / "user_files"
USER_FILES_DIR.mkdir(parents=True, exist_ok=True)

COMMENTED_DOCS_DIR = USER_FILES_DIR / "commented_docs"
COMMENTED_DOCS_DIR.mkdir(parents=True, exist_ok=True)

TOLERANCE_CONFIG_PATH = USER_FILES_DIR / "tolerance_config.json"

# --- Initialize Managers ---
# TemplateManagerWin32 expects base_user_dir to be where 'templates_map.db' and 'templates/' subdir reside
template_manager = TemplateManagerWin32(base_user_dir=USER_FILES_DIR)


# --- Helper Functions ---
def ensure_tolerance_config_exists():
    """Creates a default tolerance config if it doesn't exist."""
    if not TOLERANCE_CONFIG_PATH.exists():
        default_tolerance = {
            "pt_tolerance": 0.1,
            "multiple_tolerance": 0.05,
            "specific_tolerances": {
                "段落.首行缩进.pt": 1.0, # Example: Allow 1pt tolerance for first line indent
                "字体.大小.pt": 0.5
            }
        }
        try:
            with open(TOLERANCE_CONFIG_PATH, 'w', encoding='utf-8') as f:
                import json
                json.dump(default_tolerance, f, ensure_ascii=False, indent=2)
            st.toast(f"默认容差配置文件已创建于: {TOLERANCE_CONFIG_PATH}", icon="📄")
        except Exception as e:
            st.error(f"创建默认容差配置文件失败: {e}")

# --- Main Application UI ---
st.title("📝 Word 文档格式规范检查工具 ")
st.markdown("---")

# Ensure default tolerance config exists
ensure_tolerance_config_exists()

# Sidebar for controls (optional, can be in main page)
# st.sidebar.header("操作面板")
uploaded_file = st.file_uploader("1. 上传 Word 文档 (.docx)", type=["docx"], key="file_uploader")

available_templates = template_manager.list_selectable_templates()
if not available_templates:
    st.warning("系统中暂无可用模板。请先通过“create template”页面添加模板。")
    # Link to create_template page if it exists
    # st.page_link("pages/create_template.py", label="前往创建模板", icon="➕")
    st.stop()

template_options = {tpl['name']: tpl['id'] for tpl in available_templates}
selected_template_name = st.selectbox(
    "2. 选择一个样式模板",
    options=template_options.keys(),
    key="template_selector"
)

# Add input for first_chapter_title
first_chapter_title_input = st.text_input(
    "3. （可选）输入正文起始章节标题",
    value="绪论",
    help="用于识别文档正文部分的起始标题，例如“绪论”、“引言”等。如果留空或不匹配，将尝试从文档开头处理。"
)

col1, col2 = st.columns([1,5]) # Adjust column width ratio if needed
with col1:
    generate_report_button = st.button("生成检查报告", type="primary", use_container_width=True, key="generate_report_btn")

# Placeholders for report display
report_display_area = st.container()

if "report_generated" not in st.session_state:
    st.session_state.report_generated = False
if "differences" not in st.session_state:
    st.session_state.differences = []
if "original_doc_path" not in st.session_state:
    st.session_state.original_doc_path = None
if "doc_metadata" not in st.session_state:
    st.session_state.doc_metadata = {}


if generate_report_button and uploaded_file and selected_template_name:
    st.session_state.report_generated = False # Reset on new generation
    
    temp_file_path = None
    docx_reader = None
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_file:
            shutil.copyfileobj(uploaded_file, tmp_file)
            temp_file_path = tmp_file.name
        
        st.session_state.original_doc_path = temp_file_path # Save for commented doc later

        with st.spinner("正在打开并解析 Word 文档，请稍候..."):
            docx_reader = DocxReaderWin32(word_visible=False, word_display_alerts=False)
            docx_reader.open_document(temp_file_path)
            
            # Unpack the tuple returned by get_paragraph_data_df
            # Pass the user-provided first_chapter_title
            doc_df, middle_start_index, back_start_index = docx_reader.get_paragraph_data_df(first_chapter_title=first_chapter_title_input)
            
            # For now, document_properties might include page_setup, default_fonts etc.
            # Let's assume a simple dict for metadata for now.
            doc_metadata = {
                "total_paragraphs": len(doc_df) if doc_df is not None and not doc_df.empty else 0,
                "filename": uploaded_file.name,
                "middle_start_index": middle_start_index, # Optionally store these
                "back_start_index": back_start_index      # Optionally store these
                # Add other relevant metadata from docx_reader if available
            }
            st.session_state.doc_metadata = doc_metadata

        if doc_df is None or doc_df.empty:
            st.error("无法从文档中提取有效的段落数据。请检查文档内容。")
        else:
            selected_template_id = template_options[selected_template_name]
            template_data = template_manager.load_template_json(template_id=selected_template_id)
            print(f"[DEBUG app.py] Type of template_data from manager: {type(template_data)}") # Added for debugging
            print(f"[DEBUG app.py] Content of template_data from manager:\n{json.dumps(template_data, ensure_ascii=False, indent=2)}") # Added for debugging

            if not template_data:
                st.error(f"无法加载所选模板 '{selected_template_name}'。")
            else:
                with st.spinner("正在比较文档与模板格式..."):
                    comparator = FormatComparatorWin32(template_data, str(TOLERANCE_CONFIG_PATH))
                    # Pass middle_start_index and back_start_index to the comparator
                    differences = comparator.compare_document_formats(
                        doc_df=doc_df,
                        middle_start_index=middle_start_index,
                        back_start_index=back_start_index,
                        document_properties=None # document_properties is currently not used by comparator
                    )
                    st.session_state.differences = differences
                    st.session_state.report_generated = True
                    st.toast("格式比较完成！", icon="✅")

    except Exception as e:
        st.error(f"处理文档时发生错误: {e}")
        st.exception(e) # Show full traceback for debugging
        st.session_state.report_generated = False
    finally:
        if docx_reader:
            docx_reader.close_document()
            docx_reader.quit_word()
        # temp_file_path is handled by NamedTemporaryFile's delete=False,
        # but we might want to clean it up explicitly if it's not needed after commenting.
        # For now, keep it for the commenting step.

# --- Display Report ---
if st.session_state.report_generated and st.session_state.differences is not None:
    with report_display_area:
        st.markdown("---")
        st.subheader("📊 格式检查报告")

        report_gen = ReportGenerator(st.session_state.differences, st.session_state.doc_metadata)
        score, comment = report_gen.calculate_score_and_comment(
            DEFAULT_ERROR_WEIGHTS, DEFAULT_PENALTY_TIERS, DEFAULT_ACCELERATION_THRESHOLDS
        )
        
        display_report_summary(st, score, comment, report_gen.get_summary_stats())
        display_report_charts(st, report_gen, score)
        display_report_details_table(st, report_gen.df_diff)

        st.markdown("---")
        st.subheader("生成带批注的报告")
        st.info("此操作将直接调用 Word 应用在后台处理文档并添加批注，根据文档大小可能需要一些时间，请耐心等待。期间请勿关闭本页面。")
        
        if st.button("生成带批注的 Word 文档", key="generate_commented_doc_btn"):
            if st.session_state.original_doc_path and st.session_state.differences:
                with st.spinner("正在生成带批注的 Word 文档... 请稍候，此过程可能较慢。"):
                    commented_doc_file_path = add_comments_to_document_static(
                        original_docx_path=st.session_state.original_doc_path, # Path to the temporary file content
                        differences=st.session_state.differences,
                        output_dir=str(COMMENTED_DOCS_DIR),
                        original_file_basename=uploaded_file.name # Pass the original uploaded filename
                    )
                
                if commented_doc_file_path and os.path.exists(commented_doc_file_path):
                    st.success(f"带批注的文档已成功生成！")
                    with open(commented_doc_file_path, "rb") as fp:
                        st.download_button(
                            label="下载带批注的 Word 文档",
                            data=fp,
                            file_name=os.path.basename(commented_doc_file_path),
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )
                    # Clean up the temporary uploaded file after successful commenting and download prep
                    if st.session_state.original_doc_path and os.path.exists(st.session_state.original_doc_path):
                        try:
                            os.remove(st.session_state.original_doc_path)
                            st.session_state.original_doc_path = None # Clear path after deletion
                        except Exception as e_del:
                            st.warning(f"无法删除临时上传文件 {st.session_state.original_doc_path}: {e_del}")
                else:
                    st.error("生成带批注的文档失败。请查看应用日志获取更多信息。")
            else:
                st.warning("请先生成格式检查报告，或报告中没有差异。")

elif generate_report_button and (not uploaded_file or not selected_template_name):
    st.warning("请先上传 Word 文档并选择一个样式模板。")

st.markdown("---")
st.caption("win32com Word Format Checker v0.1.0")