import json
from pathlib import Path
import streamlit as st # For type hinting if needed, not for direct use here

# Attempt to import from the local ui_components within the win32com package
try:
    from ui_components import FONT_SIZE_MAP_DISPLAY_TO_PT, LINE_SPACING_RULES
except ImportError:
    # Fallback if run directly or import fails, though in app context this should work.
    print("警告 (ui_helpers.py): 无法从 ui_components 导入常量，将使用备用定义。")
    FONT_SIZE_MAP_DISPLAY_TO_PT = {
        "六号": 7.5, "小五": 9.0, "五号": 10.5, "小四": 12.0, "四号": 14.0,
        "小三": 15.0, "三号": 16.0, "小二": 18.0, "二号": 22.0, "小一": 24.0,
        "一号": 26.0, "小初": 36.0, "初号": 42.0
    }
    LINE_SPACING_RULES = { # Ensure this matches ui_components
        "single": "单倍行距 (倍)", "1.5 lines": "1.5 倍行距 (倍)", "double": "2 倍行距 (倍)",
        "exactly": "固定值 (磅)", "multiple": "多倍行距 (倍)"
    }

def _to_float(value, default=0.0):
    try:
        return float(value)
    except (ValueError, TypeError):
        return default

def _to_int(value, default=0):
    try:
        return int(value)
    except (ValueError, TypeError):
        return default

def form_data_to_json_win32(basic_info: dict, styles_data: dict, toc_data: dict, numbering_data: dict) -> dict:
    """
    Converts form data from the win32com UI components into a template JSON structure.

    Args:
        basic_info (dict): Basic template metadata (name).
        styles_data (dict): Configuration for main styles (e.g., "正文", "标题1").
                            Keys are internal style names, values are dicts from render_style_section.
        toc_data (dict): Configuration for TOC styles.
        numbering_data (dict): Configuration for numbering templates and links.

    Returns:
        dict: A dictionary representing the complete template JSON.
    """
    output_json = {}

    # 1. Basic Info (becomes top-level metadata in the JSON saved by TemplateManagerWin32)
    # TemplateManagerWin32's save_template will handle incorporating these into the final file.
    # Here, we are primarily concerned with the "样式" and other specific config blocks.
    # However, the plan for TemplateManagerWin32.save_template is:
    # "The method will construct a full JSON object including the provided metadata ... and the style_rules_dict"
    # So, this helper should return the *entire* structure that TemplateManagerWin32 saves.
    
    output_json['name'] = basic_info.get('template_name', '未命名模板')
    # Add a version or type marker if desired
    output_json['template_type'] = 'win32com_manual' 

    # 2. Main Styles ("样式")
    processed_styles = {}
    for style_key, style_config_from_ui in styles_data.items():
        processed_style = {}
        
        # Font processing
        font_ui = style_config_from_ui.get('字体', {})
        font_json = {
            '中文字体': font_ui.get('中文字体', '宋体'),
            '西文字体': font_ui.get('西文字体', 'Times New Roman'),
            '颜色': font_ui.get('颜色', '#000000'),
            '粗体': bool(font_ui.get('粗体', False)),
            '斜体': bool(font_ui.get('斜体', False)),
            '下划线': bool(font_ui.get('下划线', False)) # Assuming boolean, adjust if complex
        }
        display_size = font_ui.get('大小', '小四') # This is display name e.g. "小四"
        font_json['大小'] = _to_float(FONT_SIZE_MAP_DISPLAY_TO_PT.get(display_size, 12.0))
        font_json['名称'] = font_json['中文字体'] # As per ui_components logic
        processed_style['字体'] = font_json

        # Paragraph processing
        para_ui = style_config_from_ui.get('段落', {})
        para_json = {}

        # Alignment (comes from top level of style_config_from_ui)
        processed_style['对齐'] = style_config_from_ui.get('对齐', 'left')

        # Line Spacing
        ls_ui = para_ui.get('行间距', {})
        ls_rule_key = ls_ui.get('规则key', 'single')
        ls_value_input = _to_float(ls_ui.get('值'), 1.0) # User input or calculated default
        ls_unit_from_ui = ls_ui.get('单位', '倍') # Unit determined by UI based on rule

        ls_json = {"单位": ls_unit_from_ui}
        if ls_rule_key in ["exactly", "multiple"]: # "at least" removed
            ls_json["值"] = ls_value_input
            ls_json["规则"] = LINE_SPACING_RULES.get(ls_rule_key, ls_rule_key)
        else: # single, 1.5 lines, double - value is fixed by rule
            fixed_values = {"single": 1.0, "1.5 lines": 1.5, "double": 2.0}
            ls_json["值"] = fixed_values.get(ls_rule_key, 1.0)
            # "规则" field is not strictly needed for these as value implies rule
        para_json['行间距'] = ls_json
        
        # Spacing (段前, 段后)
        for space_key_zh, space_key_en in [("段前", "段前"), ("段后", "段后")]:
            space_ui = para_ui.get(space_key_zh, {})
            para_json[space_key_zh] = {
                "值": _to_float(space_ui.get('值'), 0.0),
                "单位": space_ui.get('单位', '行')
            }
        
        # Indentation (首行缩进)
        indent_ui = para_ui.get('首行缩进', {})
        para_json['首行缩进'] = {
            "值": _to_float(indent_ui.get('值'), 2.0),
            "单位": indent_ui.get('单位', '字符')
        }
        # Add left/right indent if they are part of the UI and data structure
        # For now, assuming they are not, as per simplified plan.

        processed_style['段落'] = para_json
        
        # Other properties (大纲级别, 样式类型) - usually fixed or from defaults
        processed_style['大纲级别'] = _to_int(style_config_from_ui.get('大纲级别', 9))
        processed_style['样式类型'] = style_config_from_ui.get('样式类型', 'paragraph')
        
        processed_styles[style_key] = processed_style
    
    output_json['样式'] = processed_styles

    # 3. TOC Data (Simplified - pass through what ui_components.render_toc_section collects)
    # Ensure numeric types are correct.
    cleaned_toc_data = {}
    ui_toc_title = toc_data.get('toc_title_style', {})
    if ui_toc_title.get('font'):
        ui_toc_title['font']['size'] = _to_float(ui_toc_title['font'].get('size'), 18.0)
    if ui_toc_title.get('paragraph'):
        ui_toc_title['paragraph']['space_before_pt'] = _to_float(ui_toc_title['paragraph'].get('space_before_pt'), 9.0)
        ui_toc_title['paragraph']['space_after_pt'] = _to_float(ui_toc_title['paragraph'].get('space_after_pt'), 9.0)
        ui_toc_title['paragraph']['line_spacing'] = _to_float(ui_toc_title['paragraph'].get('line_spacing'), 1.0)
    cleaned_toc_data['toc_title_style'] = ui_toc_title

    cleaned_toc_styles = {}
    for key, entry_style in toc_data.get('toc_styles', {}).items():
        cleaned_entry = entry_style.copy()
        if cleaned_entry.get('font'):
            cleaned_entry['font']['size'] = _to_float(cleaned_entry['font'].get('size'), 12.0)
        if cleaned_entry.get('paragraph'):
            cleaned_entry['paragraph']['line_spacing'] = _to_float(cleaned_entry['paragraph'].get('line_spacing'), 22.0) # Example default
        if cleaned_entry.get('tabs') and isinstance(cleaned_entry['tabs'], list) and cleaned_entry['tabs']:
            cleaned_entry['tabs'][0]['position_cm'] = _to_float(cleaned_entry['tabs'][0].get('position_cm'), 16.0)
        cleaned_toc_styles[key] = cleaned_entry
    cleaned_toc_data['toc_styles'] = cleaned_toc_styles
    # cleaned_toc_data['prefix'] = toc_data.get('prefix', '自定义') # prefix is being removed
    
    output_json.update(cleaned_toc_data) # Merge TOC data into main JSON

    # 4. Numbering Data (Simplified - pass through)
    output_json['numbering_templates'] = numbering_data.get('numbering_templates', {})
    output_json['style_numbering_links'] = numbering_data.get('style_numbering_links', {})

    return output_json

if __name__ == '__main__':
    # Example usage for testing
    # This requires mock data that would come from Streamlit's session_state
    # populated by win32com.ui_components.
    
    print("--- Testing ui_helpers_win32.form_data_to_json_win32 ---")

    mock_basic_info = {
        'template_name': 'Win32 测试模板',
    }

    # Data structure from win32com.ui_components.render_style_section
    mock_styles_data = {
        "正文": {
            "字体": {'中文字体': '宋体', '西文字体': 'Arial', '大小': '小四', '颜色': '#111111', '粗体': False, '斜体': True, '下划线': False, '名称': '宋体'},
            "对齐": "justify",
            "段落": {
                '行间距': {'规则key': 'multiple', '值': 1.25, '单位': '倍'},
                '段前': {'值': 6, '单位': '磅'},
                '段后': {'值': 6, '单位': '磅'},
                '首行缩进': {'值': 2, '单位': '字符'}
            },
            "大纲级别": 9, "样式类型": "paragraph"
        },
        "标题1": {
            "字体": {'中文字体': '黑体', '西文字体': 'Arial Black', '大小': '三号', '颜色': '#0000FF', '粗体': True, '斜体': False, '下划线': False, '名称': '黑体'},
            "对齐": "center",
            "段落": {
                '行间距': {'规则key': 'single', '值': 1.0, '单位': '倍'}, # Value will be ignored by helper for 'single'
                '段前': {'值': 12, '单位': '磅'},
                '段后': {'值': 12, '单位': '磅'},
                '首行缩进': {'值': 0, '单位': '字符'}
            },
            "大纲级别": 1, "样式类型": "paragraph"
        }
    }
    
    # Simplified mock TOC data from win32com.ui_components.render_toc_section
    mock_toc_data = {
        'toc_title_style': {
            "font": {"name_fareast": "楷体", "name_ascii": "Calibri", "size": 20.0, "bold": False},
            "paragraph": {"alignment": "left", "space_before_pt": 6.0, "space_after_pt": 6.0, "line_spacing_rule": "single", "line_spacing": 1.0}
        },
        'toc_styles': {
            "自定义TOC 1": {"font": {"size": 10.0}, "paragraph": {"line_spacing": 18.0}, "tabs": [{"position_cm": 15.0, "align": "right", "leader": "dot"}]}
        },
        'prefix': '自定义'
    }

    mock_numbering_data = {
        'numbering_templates': {"num_list_1": {"levels": [{"level": 1, "format_string": "%1."}]}},
        'style_numbering_links': {"标题1": "num_list_1"}
    }

    final_json = form_data_to_json_win32(mock_basic_info, mock_styles_data, mock_toc_data, mock_numbering_data)
    
    print("\nGenerated JSON:")
    print(json.dumps(final_json, indent=2, ensure_ascii=False))

    # Verification
    print("\nVerifications:")
    assert final_json['样式']['正文']['字体']['大小'] == 12.0 # 小四 -> 12.0
    assert final_json['样式']['正文']['段落']['行间距']['值'] == 1.25
    assert final_json['样式']['正文']['段落']['行间距']['单位'] == '倍'
    assert final_json['样式']['正文']['段落']['行间距']['规则'] == LINE_SPACING_RULES['multiple']
    
    assert final_json['样式']['标题1']['字体']['大小'] == 16.0 # 三号 -> 16.0
    assert final_json['样式']['标题1']['段落']['行间距']['值'] == 1.0 # single rule fixed value
    assert final_json['样式']['标题1']['段落']['行间距']['单位'] == '倍'
    assert '规则' not in final_json['样式']['标题1']['段落']['行间距'] # No "规则" for single/1.5/double

    print("Basic verifications passed.")