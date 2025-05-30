import streamlit as st
from pathlib import Path

# Assuming these modules are in the parent directory 'win32com'
# This relative import works if 'pages' is a sub-package of 'win32com'
# or if the streamlit multipage app structure handles paths correctly.
# When running `streamlit run win32com/app.py` from project root,
# win32com directory should be in sys.path.
try:
    from ui_components import (
        render_basic_info_form,
        render_style_section,
        render_toc_section, # Simplified version
        render_numbering_section, # Simplified version
        WIN32COM_DEFAULT_STYLES_STRUCTURE,
        WIN32COM_DEFAULT_TOC_STRUCTURE,
        WIN32COM_DEFAULT_NUMBERING_STRUCTURE
    )
    from ui_helpers import form_data_to_json_win32
    from template_manager_win32 import TemplateManagerWin32
except ImportError as e:
    st.error(f"无法导入必要的 UI 或管理模块: {e}。请确保项目结构正确，并且从项目根目录运行。")
    st.stop()

# --- Page Configuration ---
st.set_page_config(page_title="创建/编辑样式模板", layout="wide")
st.title("📝 创建或编辑样式模板 (win32com 版)")
st.markdown("---")

# --- Initialize Managers ---
# Determine base path for user_files relative to this script's location
# create_template.py is in win32com/pages/
# user_files should be in win32com/user_files/
PAGE_DIR = Path(__file__).parent # win32com/pages/
WIN32COM_ROOT_DIR = PAGE_DIR.parent # win32com/
USER_FILES_BASE_DIR = WIN32COM_ROOT_DIR / "user_files"

template_manager = TemplateManagerWin32(base_user_dir=USER_FILES_BASE_DIR)

# --- Define Styles to Configure ---
# These keys should match the keys expected by form_data_to_json_win32
# and correspond to the default structures in ui_components.
STYLES_TO_CONFIGURE = {
    "正文": "正文样式",
    "标题1": "一级标题样式 (标题1)",
    "标题2": "二级标题样式 (标题2)",
    "标题3": "三级标题样式 (标题3)", # Example, add more if needed
    "图题": "图片标题样式 (图题)",
    "表题": "表格标题样式 (表题)",
    # Add other common styles like "参考文献", "摘要" etc.
}

# --- Form Rendering ---

# 1. Basic Information
# For a new template, defaults are empty or minimal.
# If editing, these would be pre-filled. For now, focus on creation.
basic_info_defaults = {} 
basic_info_data = render_basic_info_form(defaults=basic_info_defaults)

st.markdown("---")
st.subheader("2. 主要样式配置")
st.caption("为文档中的主要元素（如正文、各级标题、图/表题等）配置格式。")

styles_form_data = {}
# Use default structures from ui_components for each style section
for internal_name, display_name in STYLES_TO_CONFIGURE.items():
    default_config_for_style = WIN32COM_DEFAULT_STYLES_STRUCTURE.get(internal_name, {})
    styles_form_data[internal_name] = render_style_section(
        style_internal_name=internal_name, 
        display_name=display_name,
        default_style_config=default_config_for_style
    )

st.markdown("---")
st.subheader("3. 目录 (TOC) 样式配置 (简化版)")
toc_form_data = render_toc_section(default_toc_config=WIN32COM_DEFAULT_TOC_STRUCTURE)

st.markdown("---")
st.subheader("4. 编号模板配置 (简化版)")
numbering_form_data = render_numbering_section(default_numbering_config=WIN32COM_DEFAULT_NUMBERING_STRUCTURE)

st.markdown("---")

# --- Save Button and Logic ---
if st.button("💾 保存模板", type="primary", use_container_width=True):
    if not basic_info_data.get('template_name'):
        st.error("模板名称是必填项，请输入模板名称。")
    else:
        with st.spinner("正在处理并保存模板..."):
            # Convert form data to the final JSON structure
            # The form_data_to_json_win32 function expects the collected data
            # from the rendering functions.
            final_template_json = form_data_to_json_win32(
                basic_info=basic_info_data,
                styles_data=styles_form_data, # This contains all configured main styles
                toc_data=toc_form_data,
                numbering_data=numbering_form_data
            )
            
            # The '样式' part for save_template is within final_template_json
            style_rules_to_save = final_template_json.get("样式", {})
            
            # TemplateManagerWin32's save_template expects the "样式" part for style_rules_dict.
            # Other metadata like school, major, creator_id are not part of its current signature.
            
            success, message = template_manager.save_template(
                name=basic_info_data['template_name'],
                style_rules_dict=style_rules_to_save # Pass only the style rules
            )

            if success:
                st.success(f"模板 '{basic_info_data['template_name']}' 保存成功！ {message}")
                st.balloons()
            else:
                st.error(f"保存模板失败: {message}")

st.markdown("---")
st.caption("Win32com Template Creator v0.1.0")