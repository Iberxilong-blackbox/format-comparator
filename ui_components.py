import streamlit as st
import json
from pathlib import Path

# --- Constants (copied and simplified from original ui_components.py) ---
COMMON_FONTS_FAREAST = ["宋体", "黑体", "楷体", "仿宋", "微软雅黑", "华文仿宋"]
COMMON_FONTS_ASCII = ["Times New Roman", "Arial", "Calibri", "Cambria Math"]
FONTS_FOR_ASCII_DROPDOWN = sorted(list(set(COMMON_FONTS_ASCII + COMMON_FONTS_FAREAST)))
ALIGNMENT_OPTIONS = {"left": "左对齐", "center": "居中", "right": "右对齐", "justify": "两端对齐", "distribute": "分散对齐"}
LINE_SPACING_RULES = {
    "single": "单倍行距 (倍)", "1.5 lines": "1.5 倍行距 (倍)", "double": "2 倍行距 (倍)",
    "exactly": "固定值 (磅)", "multiple": "多倍行距 (倍)"
}
INDENT_UNITS = {"磅": "磅", "字符": "字符", "厘米": "厘米", "行": "行"} # Added 行

FONT_SIZE_MAP_DISPLAY_TO_PT = {
    "六号": 7.5, "小五": 9.0, "五号": 10.5, "小四": 12.0, "四号": 14.0,
    "小三": 15.0, "三号": 16.0, "小二": 18.0, "二号": 22.0, "小一": 24.0,
    "一号": 26.0, "小初": 36.0, "初号": 42.0
}
FONT_SIZE_MAP_PT_TO_DISPLAY = {v: k for k, v in FONT_SIZE_MAP_DISPLAY_TO_PT.items()}

# --- UI Rendering Functions (Simplified for win32com version) ---

def render_basic_info_form(defaults: dict = None):
    """Renders the basic information form for a template."""
    if defaults is None: defaults = {}
    basic_info = {}
    with st.container(border=True):
        st.subheader("1. 模板基本信息")
        basic_info['template_name'] = st.text_input(
            "模板名称*", 
            value=defaults.get('template_name', ""), 
            help="为您的模板起一个独特的名称", 
            key="ui_template_name"
        )
    return basic_info

def render_font_options(key_prefix: str, default_font_config: dict):
    """Renders font configuration options for a style."""
    font_config = {}
    if default_font_config is None: default_font_config = {}
    
    cols = st.columns([2, 2, 1, 1, 1]) # Adjusted for potentially fewer options
    
    with cols[0]:
        default_fareast = default_font_config.get('中文字体', "宋体")
        if default_fareast not in COMMON_FONTS_FAREAST: default_fareast = "宋体"
        font_config['中文字体'] = st.selectbox(
            "中文字体", COMMON_FONTS_FAREAST, 
            index=COMMON_FONTS_FAREAST.index(default_fareast), 
            key=f"{key_prefix}_font_fareast"
        )
        font_config['名称'] = font_config['中文字体'] # Main font is East Asian

    with cols[1]:
        default_ascii = default_font_config.get('西文字体', "Times New Roman")
        if default_ascii not in FONTS_FOR_ASCII_DROPDOWN: default_ascii = "Times New Roman"
        font_config['西文字体'] = st.selectbox(
            "西文字体", FONTS_FOR_ASCII_DROPDOWN, 
            index=FONTS_FOR_ASCII_DROPDOWN.index(default_ascii), 
            key=f"{key_prefix}_font_ascii"
        )

    with cols[2]:
        default_pt_size = float(default_font_config.get('大小', 12.0))
        default_display_size = FONT_SIZE_MAP_PT_TO_DISPLAY.get(default_pt_size, '小四')
        if default_display_size not in FONT_SIZE_MAP_DISPLAY_TO_PT: default_display_size = '小四'
        
        selected_display_size = st.selectbox(
            "大小", options=list(FONT_SIZE_MAP_DISPLAY_TO_PT.keys()),
            index=list(FONT_SIZE_MAP_DISPLAY_TO_PT.keys()).index(default_display_size),
            key=f"{key_prefix}_font_size_display"
        )
        font_config['大小'] = selected_display_size # Store display name, helper will convert

    with cols[3]:
        font_config['颜色'] = st.color_picker(
            "颜色", value=default_font_config.get('颜色', "#000000"), 
            key=f"{key_prefix}_font_color"
        )
    with cols[4]:
        font_config['粗体'] = st.checkbox(
            "粗体", value=default_font_config.get('粗体', False), 
            key=f"{key_prefix}_font_bold"
        )
        font_config['斜体'] = st.checkbox(
            "斜体", value=default_font_config.get('斜体', False), 
            key=f"{key_prefix}_font_italic"
        )
        # Underline can be added if needed, for now keep it simple
        font_config['下划线'] = default_font_config.get('下划线', False) 

    return font_config

def render_paragraph_options(key_prefix: str, default_para_config: dict):
    """Renders paragraph configuration options for a style."""
    para_config = {'行间距': {}, '段前': {}, '段后': {}, '首行缩进': {}}
    if default_para_config is None: default_para_config = {}

    cols = st.columns(4)
    with cols[0]: # Line Spacing
        st.write("行间距:")
        default_ls_rule_key = default_para_config.get('行间距', {}).get('规则key', 'single') # Assume '规则key' if loading existing
        if default_ls_rule_key not in LINE_SPACING_RULES: default_ls_rule_key = 'single'
        
        selected_ls_rule_key = st.selectbox(
            "规则", options=list(LINE_SPACING_RULES.keys()),
            format_func=lambda k: LINE_SPACING_RULES[k],
            index=list(LINE_SPACING_RULES.keys()).index(default_ls_rule_key),
            key=f"{key_prefix}_para_ls_rule"
        )
        para_config['行间距']['规则key'] = selected_ls_rule_key

        default_ls_value = float(default_para_config.get('行间距', {}).get('值', 1.0))
        default_ls_unit = default_para_config.get('行间距', {}).get('单位', '倍')

        current_value = default_ls_value
        current_unit = default_ls_unit
        step = 0.1

        if selected_ls_rule_key in ["exactly"]: # "at least" removed
            current_unit = "磅"
            step = 0.5
            if default_ls_unit != "磅": current_value = 12.0 # Default to 12pt if unit mismatch
        elif selected_ls_rule_key == "multiple":
            current_unit = "倍"
            step = 0.05
            if default_ls_unit != "倍": current_value = 1.15
        else: # single, 1.5 lines, double
            current_unit = "倍"
            fixed_values = {"single": 1.0, "1.5 lines": 1.5, "double": 2.0}
            current_value = fixed_values[selected_ls_rule_key]
        
        para_config['行间距']['值'] = st.number_input(
            "值", min_value=0.0, value=current_value, step=step,
            key=f"{key_prefix}_para_ls_value",
            disabled=(selected_ls_rule_key in ["single", "1.5 lines", "double"]) # Disable for fixed rules
        )
        para_config['行间距']['单位'] = current_unit
        st.caption(f"单位: {current_unit}")

    with cols[1]: # Space Before
        st.write("段前间距:")
        default_sb_unit = default_para_config.get('段前', {}).get('单位', '行')
        if default_sb_unit not in INDENT_UNITS: default_sb_unit = '行'
        para_config['段前']['单位'] = st.selectbox(
            "单位 ", options=list(INDENT_UNITS.keys()),
            index=list(INDENT_UNITS.keys()).index(default_sb_unit),
            key=f"{key_prefix}_para_before_unit"
        )
        para_config['段前']['值'] = st.number_input(
            "值 ", min_value=0.0, 
            value=float(default_para_config.get('段前', {}).get('值', 0.0)), 
            step=0.1, key=f"{key_prefix}_para_before_value"
        )

    with cols[2]: # Space After
        st.write("段后间距:")
        default_sa_unit = default_para_config.get('段后', {}).get('单位', '行')
        if default_sa_unit not in INDENT_UNITS: default_sa_unit = '行'
        para_config['段后']['单位'] = st.selectbox(
            "单位  ", options=list(INDENT_UNITS.keys()),
            index=list(INDENT_UNITS.keys()).index(default_sa_unit),
            key=f"{key_prefix}_para_after_unit"
        )
        para_config['段后']['值'] = st.number_input(
            "值  ", min_value=0.0, 
            value=float(default_para_config.get('段后', {}).get('值', 0.0)), 
            step=0.1, key=f"{key_prefix}_para_after_value"
        )

    with cols[3]: # First Line Indent
        st.write("首行缩进:")
        default_fli_unit = default_para_config.get('首行缩进', {}).get('单位', '字符')
        if default_fli_unit not in INDENT_UNITS: default_fli_unit = '字符'
        para_config['首行缩进']['单位'] = st.selectbox(
            "单位   ", options=list(INDENT_UNITS.keys()),
            index=list(INDENT_UNITS.keys()).index(default_fli_unit),
            key=f"{key_prefix}_para_indent_unit"
        )
        para_config['首行缩进']['值'] = st.number_input(
            "值   ", min_value=0.0, 
            value=float(default_para_config.get('首行缩进', {}).get('值', 2.0)), 
            step=0.1, key=f"{key_prefix}_para_indent_value"
        )

    default_align = default_para_config.get('对齐', 'left')
    if default_align not in ALIGNMENT_OPTIONS: default_align = 'left'
    para_config['对齐'] = st.selectbox(
        "对齐方式", options=list(ALIGNMENT_OPTIONS.keys()),
        format_func=lambda k: ALIGNMENT_OPTIONS[k],
        index=list(ALIGNMENT_OPTIONS.keys()).index(default_align),
        key=f"{key_prefix}_para_align"
    )
    return para_config

def render_style_section(style_internal_name: str, display_name: str, default_style_config: dict = None):
    """Renders a configuration section for a single style."""
    if default_style_config is None: default_style_config = {}
    style_config = {}
    
    with st.expander(f"**{display_name}** 配置", expanded=False):
        st.markdown("##### 字体")
        default_font = default_style_config.get('字体', {})
        style_config['字体'] = render_font_options(style_internal_name, default_font)

        st.markdown("##### 段落")
        default_para = default_style_config.get('段落', {})
        # Merge '对齐' from top level into default_para for render_paragraph_options
        default_para_merged = {**default_para, '对齐': default_style_config.get('对齐', 'left')}
        paragraph_settings = render_paragraph_options(style_internal_name, default_para_merged)
        
        style_config['对齐'] = paragraph_settings.pop('对齐') # Move alignment back to top level of style
        style_config['段落'] = paragraph_settings
        
        # Keep other properties like 大纲级别, 样式类型 with defaults, not user-editable for now
        style_config['大纲级别'] = default_style_config.get('大纲级别', 9)
        style_config['样式类型'] = default_style_config.get('样式类型', 'paragraph')
        
    return style_config

# --- Stubs for other sections (can be expanded later) ---
def render_toc_section(default_toc_config: dict = None):
    """Placeholder for TOC styles configuration."""
    if default_toc_config is None: default_toc_config = {}
    toc_config = {}
    with st.expander("**目录 (TOC) 样式** 配置 ", expanded=False):
        st.info("目录标题和条目样式的详细配置正在施工中，comming soon（大概）。")
        # For now, return a basic structure or defaults
        toc_config['toc_title_style'] = default_toc_config.get('toc_title_style', {
            "font": {"name_fareast": "黑体", "name_ascii": "Times New Roman", "size": 18.0, "bold": True},
            "paragraph": {"alignment": "center", "space_before_pt": 9.0, "space_after_pt": 9.0, "line_spacing_rule": "single", "line_spacing": 1.0}
        })
        toc_config['toc_styles'] = default_toc_config.get('toc_styles', {}) # Keep existing structure for levels
        # prefix is being removed from the system
        # toc_config['prefix'] = default_toc_config.get('prefix', '自定义')
    return toc_config

def render_numbering_section(default_numbering_config: dict = None):
    """Placeholder for numbering templates configuration."""
    if default_numbering_config is None: default_numbering_config = {}
    numbering_data = {}
    with st.expander("**编号模板** 配置 ", expanded=False):
        st.info("多级列表编号模板的定义和样式链接正在施工中，comming soon（大概）。")
        numbering_data['numbering_templates'] = default_numbering_config.get('numbering_templates', {})
        numbering_data['style_numbering_links'] = default_numbering_config.get('style_numbering_links', {})
    return numbering_data

# Example of how to define default styles for the create_template page
# This would typically be loaded from a base JSON or defined here.
WIN32COM_DEFAULT_STYLES_STRUCTURE = {
    "正文": {"字体": {}, "段落": {}, "对齐": "left", "大纲级别": 9, "样式类型": "paragraph"},
    "标题1": {"字体": {}, "段落": {}, "对齐": "left", "大纲级别": 1, "样式类型": "paragraph"},
    "标题2": {"字体": {}, "段落": {}, "对齐": "left", "大纲级别": 2, "样式类型": "paragraph"},
    "图题": {"字体": {}, "段落": {}, "对齐": "center", "大纲级别": 9, "样式类型": "paragraph"},
    "表题": {"字体": {}, "段落": {}, "对齐": "center", "大纲级别": 9, "样式类型": "paragraph"},
    # Add other common styles as needed
}

WIN32COM_DEFAULT_TOC_STRUCTURE = {
    "toc_title_style": {"font": {}, "paragraph": {}},
    "toc_styles": { # Example for TOC 1, 2, 3
        "自定义TOC 1": {"font": {}, "paragraph": {}, "tabs": [{"position_cm": 16, "align": "right", "leader": "dot"}]},
        "自定义TOC 2": {"font": {}, "paragraph": {}, "tabs": [{"position_cm": 16, "align": "right", "leader": "dot"}]},
        "自定义TOC 3": {"font": {}, "paragraph": {}, "tabs": [{"position_cm": 16, "align": "right", "leader": "dot"}]}
    }
    # "prefix": "自定义" # prefix is being removed
}

WIN32COM_DEFAULT_NUMBERING_STRUCTURE = {
    "numbering_templates": {},
    "style_numbering_links": {}
}