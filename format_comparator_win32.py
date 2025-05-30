import os
import json
import re
import platform
import traceback
import difflib
from typing import Optional, Dict, Any, List, Tuple
from collections import defaultdict
import tempfile
import shutil
import pandas as pd
from pathlib import Path # Added
from llm_mapper import LLMStyleMapper, create_llm_client # Added
from template_manager_win32 import TemplateManagerWin32 # Added

# win32com import for add_comments_to_document_static
_HAS_WIN32COM = False
if platform.system() == "Windows":
    try:
        import win32com.client
        import pythoncom # Import pythoncom
        _HAS_WIN32COM = True
    except ImportError:
        print("警告：当前为 Windows 系统，但未找到 'pywin32' 库。无法使用 win32com 添加批注。请运行 'pip install pywin32' 安装。")

# Assuming these modules will be in the same 'win32com' directory
from unit_converter import UnitConverter
from utils import _number_to_chinese, _is_primarily_east_asian, normalize_text


class FormatComparatorWin32:
    """
    负责比较 DOCX 文档格式与模板配置 (基于 win32com 提取的数据)，并记录差异。
    """
    TEMPLATE_TO_RUNINFO_KEY_MAP = {
        "名称": "name",
        "大小": "size",
        "粗体": "bold",
        "斜体": "italic",
        "下划线": "underline_type",
        "颜色": "color_hex",
        "中文字体": "font_eastasia",
        "西文字体": "font_ascii"
    }
    RUNINFO_TO_TEMPLATE_KEY_MAP = {v: k for k, v in TEMPLATE_TO_RUNINFO_KEY_MAP.items()}

    RUN_PROPERTY_DEFAULTS = {
        'bold': False,
        'italic': False,
        'underline_type': 'none', # Default from WdUnderline map in DocxReaderWin32
        'color_hex': '#000000',   # Default to black (DocxReaderWin32 maps auto to #000000)
    }
    FONT_NAME_KEYS = {"font_eastasia", "font_ascii", "name"}

    def __init__(self, template_data: Dict[str, Any], tolerance_config_path: str):
        print(f"[DEBUG FormatComparatorWin32.__init__] Received template_data type: {type(template_data)}")
        # print(f"[DEBUG FormatComparatorWin32.__init__] Received template_data content:\n{json.dumps(template_data, ensure_ascii=False, indent=2)}")
        """
        初始化 FormatComparatorWin32。

        Args:
            template_data (Dict[str, Any]): 已加载的样式模板数据 (JSON 对象)。
            tolerance_config_path (str): 容差配置文件 JSON 路径。
        """
        self.unit_converter = UnitConverter()
        self.template_data = template_data # Store template_data for LLM mapper

        if not self.template_data:
            raise ValueError("错误：模板数据不能为空。")

        styles_definitions = self.template_data.get('样式')
        if not isinstance(styles_definitions, dict):
            raise ValueError("错误：模板数据中 '样式' 键不存在或其值不是一个字典。")
        
        # The check for nested '样式' key is removed as it was incorrect.
        # Actual style definitions are directly under the first '样式' key.
        self.target_styles = styles_definitions

        # Prefix functionality is being removed.
        # self.prefix = self.template_data.get("prefix")
        self.prefix = None # Explicitly set to None as it's being removed
        
        # Template name should also be sourced from the top level of template_data.
        self.template_name = self.template_data.get("name", "未命名模板")

        # Since prefix is removed, unprefixed_target_styles is a direct copy of target_styles.
        # Or, subsequent code could be refactored to use self.target_styles directly.
        # For now, maintain the structure by copying.
        self.unprefixed_target_styles = self.target_styles.copy()
        
        self.tolerance_config = self._load_tolerance_config(tolerance_config_path)
        self.differences: List[Dict[str, Any]] = []
        self.llm_style_map: Dict[int, str] = {}
        self.doc_df: Optional[pd.DataFrame] = None

        # Initialize LLM components
        self.llm_client = create_llm_client() # Name kept as create_llm_client
        
        # Determine base_user_dir for TemplateManagerWin32 relative to this file's location
        # Assuming this file (format_comparator_win32.py) is in 'win32com/'
        # and 'user_files' is a sibling to 'win32com/' or at a known relative path.
        current_file_dir = Path(__file__).parent
        user_files_base_dir = current_file_dir / "user_files"

        self.template_manager_for_llm = TemplateManagerWin32(base_user_dir=user_files_base_dir)
        self.mapper_generator: Optional[LLMStyleMapper] = None
        if self.llm_client:
            self.mapper_generator = LLMStyleMapper(
                template_manager=self.template_manager_for_llm,
                llm_client=self.llm_client,
                template_data=self.template_data # Pass the already loaded template_data
            )
            print("LLMStyleMapper (win32com) initialized.")
        else:
            print("LLM client for win32com could not be initialized. LLM mapping will be skipped.")

    def _load_tolerance_config(self, path: str) -> Dict[str, Any]:
        default_tolerance = {
            "pt_tolerance": 0.1,
            "multiple_tolerance": 0.05,
            "specific_tolerances": {}
        }
        if not os.path.exists(path):
            print(f"警告：容差配置文件 {path} 不存在，将使用默认容差。")
            return default_tolerance
        try:
            with open(path, 'r', encoding='utf-8') as f:
                config = json.load(f)
                return config
        except Exception as e:
            print(f"警告：加载容差配置文件 {path} 失败: {e}。将使用默认容差。")
            return default_tolerance

    def _get_tolerance(self, property_key: str, unit: str) -> float:
        specific_key = f"{property_key}.{unit}"
        if specific_key in self.tolerance_config.get("specific_tolerances", {}):
            return self.tolerance_config["specific_tolerances"][specific_key]
        elif unit == "pt":
            return self.tolerance_config.get("pt_tolerance", 0.1)
        elif unit == "multiple":
            return self.tolerance_config.get("multiple_tolerance", 0.05)
        else:
            return 0.0

    def _compare_values(self,
                        expected_template_value: Any,
                        actual_doc_value: Any,
                        property_key: str,
                        template_font_size_pt: Optional[float],
                        actual_context: Optional[Dict[str, Any]] = None
                       ) -> Tuple[bool, Optional[str], Optional[str]]:
        # This is the full _compare_values method from the original format_comparator.py
        # It's assumed that its logic is largely applicable.
        # Minor adaptations might be needed based on exact DocxReaderWin32 output formats.
        def format_value(val):
            if val is None:
                return "未定义/默认"
            return str(val)

        formatted_expected = format_value(expected_template_value)
        formatted_actual = format_value(actual_doc_value)
        # print(f"  [DEBUG _compare_values] Start: property='{property_key}', expected='{expected_template_value}' (type: {type(expected_template_value)}), actual='{actual_doc_value}' (type: {type(actual_doc_value)}), font_size_pt='{template_font_size_pt}'")

        if property_key in ["段落.段前", "段落.段后"]:
            is_expected_zero_lines = False
            if isinstance(expected_template_value, dict) and \
               expected_template_value.get('单位', '').lower() in ['行', 'line'] and \
               expected_template_value.get('值') == 0:
                is_expected_zero_lines = True
            elif isinstance(expected_template_value, str):
                val_exp_parsed, unit_exp_parsed = self.unit_converter.parse_value(expected_template_value)
                if val_exp_parsed == 0 and unit_exp_parsed == 'line':
                    is_expected_zero_lines = True
            
            if is_expected_zero_lines:
                actual_is_zero_pt = False
                if isinstance(actual_doc_value, (int, float)) and abs(actual_doc_value) < 0.01:
                    actual_is_zero_pt = True
                if actual_is_zero_pt:
                    fmt_act_debug = f"{actual_doc_value:.2f} 磅" if actual_doc_value is not None else "0.00 磅"
                    return True, "0 行", fmt_act_debug

        if expected_template_value is None and actual_doc_value is None:
            return True, formatted_expected, formatted_actual
        if expected_template_value is None or actual_doc_value is None:
            # print(f"  比较差异 (None/缺失): 属性='{property_key}', 预期='{formatted_expected}', 实际='{formatted_actual}'")
            return False, formatted_expected, formatted_actual

        target_unit = None
        comparison_type = 'strict'

        is_expected_dict_format = isinstance(expected_template_value, dict) and '值' in expected_template_value and '单位' in expected_template_value

        if is_expected_dict_format:
            comparison_type = 'numeric_convert'
            unit_from_dict = expected_template_value.get('单位', '').lower()
            if unit_from_dict in ['倍', 'multiple']:
                 target_unit = 'multiple'
            elif unit_from_dict in ['磅', 'pt', '厘米', 'cm', '英寸', 'inch', '字符', 'char', '行', 'line']:
                 target_unit = 'pt'
            else:
                 target_unit = 'pt'
        elif isinstance(expected_template_value, str) and any(u in expected_template_value for u in ['pt', 'cm', '倍', '行', '字符']):
             comparison_type = 'numeric_convert'
             if '行间距' in property_key and ('倍' in expected_template_value or '行' in expected_template_value):
                 target_unit = 'multiple'
             else:
                 target_unit = 'pt'
        elif isinstance(expected_template_value, bool) or isinstance(actual_doc_value, bool):
            comparison_type = 'boolean'
            try:
                expected_b = bool(expected_template_value)
                
                actual_b: bool
                if property_key == "字体.下划线" and isinstance(actual_doc_value, str):
                    # For underline, "none" (string) means False, other underline type strings mean True.
                    actual_b = actual_doc_value.lower() != 'none'
                else:
                    actual_b = bool(actual_doc_value)
                    
            except Exception as e:
                 # print(f"  比较错误: 无法将 '{property_key}' 的值转换为布尔值. 预期: {expected_template_value}, 实际: {actual_doc_value}, 错误: {e}")
                 return False, formatted_expected, formatted_actual # Use initial formatted values

            # Display boolean values consistently
            if expected_b == actual_b:
                 return True, str(expected_b), str(actual_b)
            else:
                 # print(f"  比较差异 (布尔): 属性='{property_key}', 预期='{expected_b}', 实际='{actual_b}'")
                 return False, str(expected_b), str(actual_b)
        elif isinstance(expected_template_value, (int, float)) or isinstance(actual_doc_value, (int, float)):
             comparison_type = 'numeric'
             target_unit = 'pt'
             if '行间距' in property_key and isinstance(actual_doc_value, (float, int)) and actual_doc_value < 10: # Heuristic
                 target_unit = 'multiple'
             elif '大小' in property_key:
                  target_unit = 'pt'
        elif isinstance(expected_template_value, str) or isinstance(actual_doc_value, str):
             comparison_type = 'string'
             expected_str = str(expected_template_value)
             actual_str = str(actual_doc_value)
             if 'color' in property_key.lower() or '颜色' in property_key: # Handles color_hex
                 norm_expected = expected_str.lstrip('#').upper()
                 norm_actual = actual_str.lstrip('#').upper()
                 is_match = (norm_expected == norm_actual)
                 # if not is_match: print(f"  比较差异 (颜色): 属性='{property_key}', 预期='{expected_str}', 实际='{actual_str}'")
                 return is_match, expected_str, actual_str
             else: # Default string comparison (e.g. for underline_type, font names)
                 is_match = (expected_str == actual_str)
                 # if not is_match: print(f"  比较差异 (字符串): 属性='{property_key}', 预期='{expected_str}', 实际='{actual_str}'")
                 return is_match, expected_str, actual_str
        
        converted_expected = None
        converted_actual = None
        try:
            if comparison_type in ['numeric', 'numeric_convert']:
                if comparison_type == 'numeric_convert':
                    if isinstance(expected_template_value, dict):
                         val_exp = expected_template_value.get('值')
                         unit_exp = expected_template_value.get('单位')
                    else:
                         val_exp, unit_exp = self.unit_converter.parse_value(expected_template_value)
                    if val_exp is None:
                         return False, format_value(expected_template_value), format_value(actual_doc_value)
                    converted_expected = self.unit_converter.convert_value(val_exp, unit_exp, target_unit, font_size_pt=template_font_size_pt)
                    formatted_expected = f"{converted_expected:.2f} {target_unit}" if converted_expected is not None else "转换失败"
                else: # numeric
                    if isinstance(expected_template_value, (int, float)):
                         converted_expected = float(expected_template_value)
                         formatted_expected = f"{converted_expected:.2f} {target_unit}"
                    else:
                         return False, format_value(expected_template_value), format_value(actual_doc_value)

                actual_val_num, actual_unit_parsed = self.unit_converter.parse_value(actual_doc_value)
                unit_for_conversion = actual_unit_parsed
                if property_key == '段落.行间距' and isinstance(actual_doc_value, (int, float)) and actual_unit_parsed is None:
                    actual_rule = actual_context.get('line_spacing_rule') if actual_context else None
                    # DocxReaderWin32 provides line_spacing_rule as string like "single", "multiple", "atLeast", "exactly"
                    if actual_rule in ['atLeast', 'exactly', 'wdLineSpaceAtLeast', 'wdLineSpaceExactly']:
                         unit_for_conversion = 'pt'
                    elif actual_rule in ['multiple', 'single', 'double', 'wdLineSpaceMultiple', 'wdLineSpaceSingle', 'wdLineSpaceDouble', None]:
                         unit_for_conversion = 'multiple'
                
                if actual_val_num is not None:
                     converted_actual = self.unit_converter.convert_value(actual_val_num, unit_for_conversion, target_unit, font_size_pt=template_font_size_pt)
                elif isinstance(actual_doc_value, (int, float)):
                     actual_unit_implicit = unit_for_conversion if unit_for_conversion else 'pt'
                     if target_unit == actual_unit_implicit:
                         converted_actual = float(actual_doc_value)
                     else:
                         converted_actual = self.unit_converter.convert_value(float(actual_doc_value), actual_unit_implicit, target_unit, font_size_pt=template_font_size_pt)
                else:
                     return False, formatted_expected, format_value(actual_doc_value)
                
                formatted_actual = f"{converted_actual:.2f} {target_unit}" if converted_actual is not None else "转换失败"

                if converted_expected is None or converted_actual is None:
                    return False, formatted_expected, formatted_actual
                
                tolerance = self._get_tolerance(property_key, target_unit)
                is_match = abs(converted_expected - converted_actual) <= tolerance
                # if not is_match: print(f"  比较差异 (数值): 属性='{property_key}', 预期='{formatted_expected}', 实际='{formatted_actual}', 容差='{tolerance}'")
                
                fmt_expected_final = formatted_expected.replace(" pt", " 磅")
                fmt_actual_final = formatted_actual.replace(" pt", " 磅")

                # START Phase Five modification for actual value display
                if property_key == "段落.首行缩进":
                    # Check if actual value is effectively zero after conversion
                    # Use a small epsilon or the defined tolerance for comparison
                    # Using self._get_tolerance for consistency if appropriate, or a fixed small epsilon.
                    # converted_actual is the numeric value in target_unit (usually 'pt')
                    if converted_actual is not None and abs(converted_actual) < 0.01: # Assuming 0.01 pt is effectively zero
                        fmt_actual_final = "未设置"
                # END Phase Five modification

                # For "段落.首行缩进", adjust display of expected and actual if original unit was 'char'
                if property_key == "段落.首行缩进":
                    # Check expected value format
                    expected_val_parsed, expected_unit_parsed = None, None
                    original_expected_char_str = None

                    if isinstance(expected_template_value, dict) and expected_template_value.get('单位') == 'char':
                        expected_val_parsed = expected_template_value.get('值')
                        expected_unit_parsed = 'char'
                        original_expected_char_str = f"{expected_val_parsed}{expected_template_value.get('单位', '')}"
                    elif isinstance(expected_template_value, str):
                        val_check, unit_check = self.unit_converter.parse_value(expected_template_value)
                        if unit_check == 'char':
                            expected_val_parsed = val_check
                            expected_unit_parsed = 'char'
                            original_expected_char_str = expected_template_value
                    
                    # If expected was in 'char', re-format fmt_expected_final and fmt_actual_final
                    if expected_unit_parsed == 'char' and expected_val_parsed is not None:
                        font_size_to_use_for_display = template_font_size_pt # This is actual_para_font_size or template_font_size
                        
                        # Re-format expected value string
                        font_size_info_expected = f"(基于字号 {font_size_to_use_for_display:.2f} 磅)" if font_size_to_use_for_display is not None else "(基于字号)"
                        # fmt_expected_final is already in points, e.g., "24.00 磅"
                        if original_expected_char_str:
                             fmt_expected_final = f"{original_expected_char_str} (约 {fmt_expected_final}) {font_size_info_expected}"
                        else: # Should not happen if original_expected_char_str was derived correctly
                             fmt_expected_final = f"{fmt_expected_final} {font_size_info_expected} (原始单位: 字符)"

                        # Re-format actual value string to include estimated char value
                        if converted_actual is not None and font_size_to_use_for_display is not None and font_size_to_use_for_display > 0:
                            # Check if fmt_actual_final is "未设置" due to previous modification
                            if fmt_actual_final != "未设置":
                                try:
                                    # Estimate char value from actual points and actual font size
                                    # Assuming 1 char width/height is approx. font_size_pt for Chinese fonts,
                                    # or half for some specific cases (this is a simplification).
                                    # For simplicity, let's assume 1 char = font_size_pt for indent.
                                    # More accurate would be: char_in_pt = font_size_pt (for square chars like Chinese)
                                    # For indent, Word often uses the font size of the "Normal" style or current para's font.
                                    # Here, font_size_to_use_for_display is the best guess we have for the paragraph's font size.
                                    actual_chars_estimated = self.unit_converter.convert_value(converted_actual, 'pt', 'char', font_size_pt=font_size_to_use_for_display)
                                    if actual_chars_estimated is not None:
                                        fmt_actual_final = f"{fmt_actual_final} (约 {actual_chars_estimated:.1f} 字符)"
                                except Exception:
                                    pass # Keep original fmt_actual_final if char conversion fails
                return is_match, fmt_expected_final, fmt_actual_final
            
            # Fallback for unhandled types
            # print(f"  比较警告: 使用严格比较处理未知类型 '{property_key}'. 预期: {expected_template_value}, 实际: {actual_doc_value}")
            is_match = (expected_template_value == actual_doc_value)
            fmt_expected_final = format_value(expected_template_value).replace(" pt", " 磅")
            fmt_actual_final = format_value(actual_doc_value).replace(" pt", " 磅")
            return is_match, fmt_expected_final, fmt_actual_final

        except Exception as e:
            # print(f"  比较错误: 属性 '{property_key}' 比较时发生异常: {e}. 预期: {expected_template_value}, 实际: {actual_doc_value}")
            # traceback.print_exc()
            fmt_expected_final_err = format_value(expected_template_value).replace(" pt", " 磅")
            fmt_actual_final_err = format_value(actual_doc_value).replace(" pt", " 磅")
            return False, fmt_expected_final_err, fmt_actual_final_err

    def _find_target_style(self, paragraph_series: pd.Series) -> Tuple[Optional[Dict[str, Any]], Optional[str], Optional[str], Optional[str]]:
        style_name = paragraph_series.get('style_name')
        outline_level = paragraph_series.get('outline_level') # This is an integer from DocxReaderWin32
        text = paragraph_series.get('text', '')
        print(f"[DEBUG _find_target_style] Input - Style Name: '{style_name}', Outline Level: {outline_level}, Text: '{text[:50]}'")
        # self.prefix is removed, no need to print it.
        print(f"[DEBUG _find_target_style] self.unprefixed_target_styles keys: {list(self.unprefixed_target_styles.keys())}")

        target_style_info = None
        target_style_name = None
        original_target_style_name = None # Template style name with prefix
        mapping_method = None

        if style_name: # P1: Explicit style name match
            print(f"[DEBUG P1] style_name: '{style_name}', Attempting match in unprefixed_target_styles.")
            # Since self.prefix is removed, unprefixed_style_name is just style_name.
            # self.unprefixed_target_styles is now a copy of self.target_styles.
            if style_name in self.unprefixed_target_styles:
                target_style_info = self.unprefixed_target_styles[style_name]
                target_style_name = style_name
                original_target_style_name = style_name # As prefix is gone, original name is the matched name
                print(f"[DEBUG P1] Match SUCCESS: target_style_name='{target_style_name}', original_target_style_name='{original_target_style_name}'")
                mapping_method = 'P1'
                print(f"    _find_target_style: P1 匹配: '{style_name}' -> '{target_style_name}' (原始: {original_target_style_name})")
                return target_style_info, target_style_name, original_target_style_name, mapping_method
        
        # P-LLM: LLM 映射建议 (在 P1 失败后)
        # paragraph_series.name should give the paragraph_index (0-based)
        current_para_idx = paragraph_series.name
        if self.llm_style_map and current_para_idx in self.llm_style_map:
            llm_mapped_style_with_prefix = self.llm_style_map[current_para_idx]
            
            # Try to find this style (with prefix) in the original target_styles
            if llm_mapped_style_with_prefix in self.target_styles:
                target_style_info = self.target_styles[llm_mapped_style_with_prefix]
                original_target_style_name = llm_mapped_style_with_prefix
                
                # Derive the unprefixed name for consistency
                # Since self.prefix is removed, target_style_name is original_target_style_name
                target_style_name = original_target_style_name
                
                mapping_method = 'P-LLM'
                print(f"    _find_target_style: P-LLM 匹配: 段落 {current_para_idx} -> '{target_style_name}' (原始: {original_target_style_name})")
                return target_style_info, target_style_name, original_target_style_name, mapping_method
            else:
                print(f"    _find_target_style: P-LLM 警告: LLM 为段落 {current_para_idx} 生成的映射样式 '{llm_mapped_style_with_prefix}' 在模板中未找到。")

        # P2: Outline level for headings
        processed_outline_level_for_p2 = None
        if isinstance(outline_level, (int, float)):
            try:
                if outline_level == int(outline_level): # Ensure it's a whole number if float
                    processed_outline_level_for_p2 = int(outline_level)
            except (ValueError, TypeError):
                pass

        if processed_outline_level_for_p2 is not None and 1 <= processed_outline_level_for_p2 <= 8:
            expected_heading_style = f"标题{_number_to_chinese(processed_outline_level_for_p2)}"
            print(f"[DEBUG P2] outline_level: {outline_level} (processed as {processed_outline_level_for_p2}), calculated expected_heading_style: '{expected_heading_style}', Attempting match.")
            if expected_heading_style in self.unprefixed_target_styles:
                target_style_info = self.unprefixed_target_styles[expected_heading_style]
                target_style_name = expected_heading_style
                # Since self.prefix is removed, original_target_style_name is target_style_name
                original_target_style_name = target_style_name
                mapping_method = 'P2'
                return target_style_info, target_style_name, original_target_style_name, mapping_method

        # P3: Outline level 9 (Body text, captions, formulas)
        # DocxReaderWin32: outline_level for body text is typically 9 (wdOutlineLevelBodyText)
        if outline_level == 9: # wdOutlineLevelBodyText
            FIG_CAPTION_REGEX = re.compile(r"^\s*(?:图|Figure|Fig\.?)\s*\d+(?:[-.]\d+)*") # Regex for Figure Captions
            print(f"[DEBUG P3-FigCaption] FIG_CAPTION_REGEX matched. Checking for '图题'.")
            TAB_CAPTION_REGEX = re.compile(r"^\s*(?:表|Table)\s*\d+(?:[-.]\d+)*")      # Regex for Table Captions
            FORMULA_REGEX = re.compile(r"\(\s*\d+(?:[-.]\d+)\s*\)\s*$") # Formula number at end
            potential_style = None
            temp_mapping_method = None

            if FIG_CAPTION_REGEX.match(text):
                if "图题" in self.unprefixed_target_styles: # Check for "图题" (Figure Caption)
                    potential_style = "图题"
                    temp_mapping_method = 'P3-FigCaption'
                print(f"[DEBUG P3-Formula] FORMULA_REGEX matched. Checking for '公式'.")
            elif TAB_CAPTION_REGEX.match(text):
                if "表题" in self.unprefixed_target_styles: # Check for "表题" (Table Caption)
                    potential_style = "表题"
                    temp_mapping_method = 'P3-TabCaption'
            elif FORMULA_REGEX.search(text):
                if "公式" in self.unprefixed_target_styles:
                    potential_style = "公式"
                    temp_mapping_method = 'P3-Formula'
            
            if potential_style is None: # Default to "正文" for outline level 9 if not caption/formula
                if "正文" in self.unprefixed_target_styles:
                    print(f"[DEBUG P3] Match SUCCESS via {temp_mapping_method}: target_style_name='{target_style_name}', original_target_style_name='{original_target_style_name}'")
                    potential_style = "正文"
                    temp_mapping_method = 'P3-Body'

            if potential_style and potential_style in self.unprefixed_target_styles:
                target_style_info = self.unprefixed_target_styles[potential_style]
                target_style_name = potential_style
                print(f"[DEBUG _find_target_style] All P1, P2, P3 attempts FAILED for Style Name: '{style_name}', Outline Level: {outline_level}, Text: '{text[:50]}'")
                mapping_method = temp_mapping_method
                # Since self.prefix is removed, original_target_style_name is target_style_name
                original_target_style_name = target_style_name
                return target_style_info, target_style_name, original_target_style_name, mapping_method
        
        return None, None, None, None # P5: No match

    def compare_document_formats(self,
                                 doc_df: pd.DataFrame,
                                 middle_start_index: Optional[int] = None, # Added
                                 back_start_index: Optional[int] = None,
                                 document_properties: Optional[Dict[str, Any]] = None
                                 ):
        self.differences = []
        self.doc_df = doc_df # Store doc_df for LLM mapper if needed

        # --- LLM Style Mapping ---
        if self.mapper_generator:
            print("\n开始调用 LLM 生成样式映射 (win32com)...")
            try:
                # Pass the template_name stored in self.template_name
                llm_mappings_list = self.mapper_generator.generate_mapping(
                    doc_df=self.doc_df,
                    middle_start_index=middle_start_index, # Pass this along
                    back_start_index=back_start_index,   # Pass this along
                    template_name=self.template_name
                )
                if llm_mappings_list:
                    for item in llm_mappings_list:
                        if isinstance(item, dict) and 'paragraph_index' in item and 'style' in item:
                            try:
                                para_idx_llm = int(item['paragraph_index'])
                                self.llm_style_map[para_idx_llm] = str(item['style'])
                            except (ValueError, TypeError):
                                print(f"警告：LLM generate_mapping 返回了无效的 paragraph_index 或 style: {item}，已跳过。")
                    print(f"LLM 样式映射生成完成，共获得 {len(self.llm_style_map)} 条有效映射。")
                else:
                    print("LLM 未能生成任何样式映射。")
            except RuntimeError as llm_error:
                print(f"错误：调用 LLM 生成样式映射时发生错误: {llm_error}。将忽略 LLM 映射。")
            except Exception as e_llm:
                print(f"错误：调用 LLM 生成样式映射时发生意外异常: {e_llm}。将忽略 LLM 映射。")
        else:
            print("LLMStyleMapper (win32com) 未初始化，跳过 LLM 样式映射。")
        # --- End LLM Style Mapping ---

        # --- Filter DataFrame for comparison loop ---
        comparison_df = self.doc_df # Start with the full df (self.doc_df was set from input doc_df)
        if 'segment_type' in self.doc_df.columns:
            body_only_df = self.doc_df[self.doc_df['segment_type'] == 'body_matter']
            if not body_only_df.empty:
                print(f"信息：检测到 'segment_type' 列。将只对 'body_matter' 部分（{len(body_only_df)}段）进行格式比较。")
                comparison_df = body_only_df
            elif not self.doc_df.empty: # Original df was not empty, but body_only_df is
                print("警告：'segment_type' 列存在，但未找到 'body_matter' 段落。将比较整个文档以避免无结果。")
                # comparison_df remains self.doc_df (full)
            # If self.doc_df was empty, comparison_df is also empty.
        else:
            print("警告：'segment_type' 列未在文档 DataFrame 中找到。将比较整个文档。")
            # comparison_df remains self.doc_df (full)

        # Iterate over the (potentially filtered) DataFrame for comparison
        for para_idx, para_series in comparison_df.iterrows():
            print(f"[DEBUG compare_document_formats] Para Index: {para_idx}, Style Name: '{para_series.get('style_name')}', Outline Level: {para_series.get('outline_level')}, Text Preview: '{para_series.get('text', '')[:50]}'")
            target_style_info, target_style_name, original_target_style_name, mapping_method = self._find_target_style(para_series)

            if target_style_info is None:
                print(f"[DEBUG compare_document_formats] Fallback FAILED for Para Index: {para_idx}. _find_target_style returned None. Doc Style: '{para_series.get('style_name', '无')}', Outline: {para_series.get('outline_level', '无')}")
                if "正文" in self.unprefixed_target_styles:
                    target_style_name = "正文"
                    target_style_info = self.unprefixed_target_styles[target_style_name]
                    for k_orig, v_orig in self.target_styles.items():
                        if self.prefix and k_orig.startswith(self.prefix) and k_orig[len(self.prefix):] == target_style_name:
                            original_target_style_name = k_orig; break
                        elif not self.prefix and k_orig == target_style_name:
                            original_target_style_name = k_orig; break
                    mapping_method = "P5-FallbackToBody"
                else:
                    self.differences.append(self._format_difference(
                        para_idx, para_series.get('text', ''), "整体样式", "未找到匹配模板样式",
                        f"文档样式: {para_series.get('style_name', '无')}, 大纲级别: {para_series.get('outline_level', '无')}",
                        None, "P5"
                    ))
                    continue
            
            _template_font_size_pt_for_indent = None
            target_font_style_for_size = target_style_info.get("字体", {})
            template_font_size_config = target_font_style_for_size.get("大小")
            if template_font_size_config:
                val_fs, unit_fs = self.unit_converter.parse_value(template_font_size_config)
                if val_fs is not None:
                    _template_font_size_pt_for_indent = self.unit_converter.convert_value(value=val_fs, from_unit=unit_fs, to_unit='pt')

            target_para_template = target_style_info.get("段落", {})
            if target_para_template:
                # Pass the actual paragraph font size if available, otherwise fallback to template font size for indent calc
                actual_para_font_size_pt = para_series.get('paragraph_actual_font_size_pt')
                font_size_for_indent_calc = actual_para_font_size_pt if actual_para_font_size_pt is not None else _template_font_size_pt_for_indent
                self._compare_paragraph_properties(para_idx, para_series, target_para_template, target_style_name, mapping_method, font_size_for_indent_calc)

            target_font_template = target_style_info.get("字体", {})
            actual_runs_info = para_series.get('font_info', [])
            if target_font_template and actual_runs_info:
                 self._compare_run_properties(para_idx, para_series.get('text', ''), actual_runs_info, target_font_template, target_style_name, mapping_method)
        
        return self.differences

    def _compare_paragraph_properties(self, para_idx: int, para_series: pd.Series, target_para_template: Dict[str, Any], target_style_name: str, mapping_method: Optional[str], template_font_size_pt: Optional[float]):
        para_text_preview = para_series.get('text', '')[:30]
        for template_key_zh, expected_val in target_para_template.items():
            actual_val = None
            property_full_key = f"段落.{template_key_zh}"
            actual_context = None
            # Map template_key_zh to DocxReaderWin32 DataFrame column name
            # DocxReaderWin32 provides: alignment (str), left_indent (float, pt), right_indent (float, pt),
            # first_line_indent (float, pt), space_before (float, pt), space_after (float, pt),
            # line_spacing_value (float), line_spacing_rule_str (str), outline_level (int)
            if template_key_zh == "对齐方式": actual_val = para_series.get('alignment') # String like 'left', 'center'
            elif template_key_zh == "左缩进": actual_val = para_series.get('left_indent_pt')
            elif template_key_zh == "右缩进": actual_val = para_series.get('right_indent_pt')
            elif template_key_zh == "首行缩进": actual_val = para_series.get('first_line_indent_pt')
            elif template_key_zh == "段前": actual_val = para_series.get('space_before_pt')
            elif template_key_zh == "段后": actual_val = para_series.get('space_after_pt')
            elif template_key_zh == "行间距":
                actual_val = para_series.get('line_spacing_value')
                actual_context = {'line_spacing_rule': para_series.get('line_spacing_rule')} # e.g. "multiple", "exactly"
            elif template_key_zh == "大纲级别":
                actual_val = para_series.get('outline_level') # Integer
                # Template might store integer or string. _compare_values needs to handle this.
                # For now, assume template stores integer for levels 1-9.
            else: continue

            # Determine which font size to use for 'char' unit conversion in _compare_values
            # For paragraph properties like indent, template_font_size_pt (passed to this method)
            # should already be the actual paragraph font size if available, or template's default.
            font_size_for_conversion = template_font_size_pt

            is_match, fmt_expected, fmt_actual = self._compare_values(
                expected_val, actual_val, property_full_key, font_size_for_conversion, actual_context
            )
            if not is_match:
                self.differences.append(self._format_difference(
                    para_idx, para_text_preview, property_full_key, fmt_expected, fmt_actual, target_style_name, mapping_method
                ))

    def _compare_run_properties(self, para_idx: int, para_text: str, actual_runs_info: List[Dict[str, Any]], target_font_template: Dict[str, Any], target_style_name: str, mapping_method: Optional[str]):
        para_text_preview = para_text[:30]
        
        # Structure to collect differences for this paragraph before adding to self.differences
        # Key: property_full_key (e.g., "字体.大小")
        # Value: {'expected': formatted_expected_value,
        #         'actuals': defaultdict(list_of_run_identifiers)}
        # Example: {'字体.大小': {'expected': '12.0 磅', 'actuals': {'10.5 磅': ['Run 0 (文本: "示例")', 'Run 2 (文本: "文字")']}}}
        collected_diffs_for_paragraph = defaultdict(lambda: {'expected': None, 'actuals': defaultdict(list)})

        for run_idx, actual_run_detail in enumerate(actual_runs_info):
            run_text_content = actual_run_detail.get('text', '').strip()
            # Create a more descriptive identifier for the run
            run_identifier = f"片段 {run_idx + 1}"
            if run_text_content:
                run_identifier = f"“{run_text_content[:15]}” (片段 {run_idx + 1})" # Show up to 15 chars of text

            for template_key_zh, expected_val_from_template in target_font_template.items():
                actual_val_from_run = None
                property_full_key = f"字体.{template_key_zh}"
                run_info_key = self.TEMPLATE_TO_RUNINFO_KEY_MAP.get(template_key_zh)

                if not run_info_key:
                    continue
                
                actual_val_from_run = actual_run_detail.get(run_info_key)
                if actual_val_from_run is None and run_info_key in self.RUN_PROPERTY_DEFAULTS:
                     actual_val_from_run = self.RUN_PROPERTY_DEFAULTS[run_info_key]

                # Previous [DEBUG RunFormat] logs can be removed as Issue #2 is resolved.
                # print(f"[DEBUG RunFormat] ParaIdx: {para_idx}, RunIdx: {run_idx}, RunText: '{actual_run_detail.get('text', '')[:20]}'")
                # ...

                is_match, fmt_expected, fmt_actual = self._compare_values(
                    expected_val_from_template, actual_val_from_run, property_full_key, None
                )

                if not is_match:
                    if collected_diffs_for_paragraph[property_full_key]['expected'] is None:
                        collected_diffs_for_paragraph[property_full_key]['expected'] = fmt_expected
                    collected_diffs_for_paragraph[property_full_key]['actuals'][fmt_actual].append(run_identifier)
        
        # Now, iterate through collected_diffs_for_paragraph and add to self.differences
        for property_key, diff_data in collected_diffs_for_paragraph.items():
            expected_value_str = diff_data['expected']
            if not diff_data['actuals']: # Should not happen if 'expected' was set due to a mismatch
                continue

            for actual_value_str, run_identifiers_list in diff_data['actuals'].items():
                location_detail = ""
                if not run_identifiers_list: # Should not happen
                    location_detail = "特定文本片段 (无详细信息)"
                elif len(run_identifiers_list) == len(actual_runs_info) and len(actual_runs_info) > 0:
                    # Check if all runs in the paragraph have this specific differing actual value for this property
                    all_runs_share_this_actual = True
                    # This check is simplified; a more robust check might involve comparing total number of runs
                    # if actual_runs_info was filtered for non-empty runs, etc.
                    # For now, if all collected identifiers for this actual_value_str cover all runs, assume so.
                    if len(actual_runs_info) > 1 : # only say "all" if there's more than one run
                        location_detail = "段落内所有文本片段"
                    else: # Single run paragraph, or single run with this specific diff
                        location_detail = run_identifiers_list[0] # Show the single run detail
                else:
                    max_examples_to_show = 3 # Show up to 3 examples
                    shown_examples = run_identifiers_list[:max_examples_to_show]
                    location_detail = f"文本: {', '.join(shown_examples)}"
                    if len(run_identifiers_list) > max_examples_to_show:
                        location_detail += f" 等 {len(run_identifiers_list) - max_examples_to_show} 处"
                
                self.differences.append(self._format_difference(
                    para_idx,
                    para_text_preview,
                    property_key,
                    expected_value_str,
                    actual_value_str,
                    target_style_name,
                    mapping_method,
                    location_detail=location_detail
                ))

    def _normalize_font_name(self, font_name: Optional[str]) -> Optional[str]:
        if font_name is None: return None
        return font_name.strip().lower()

    def _is_default_run_property(self, prop_key: str, actual_doc_value: Any) -> bool:
        # Check if actual_doc_value matches the defined default for prop_key
        default_val = self.RUN_PROPERTY_DEFAULTS.get(prop_key)
        if default_val is None and prop_key not in self.RUN_PROPERTY_DEFAULTS: # Key not in defaults means not defaultable by this check
            return False
        return actual_doc_value == default_val

    def _format_difference(self, para_idx: int, para_text_preview: str, property_name: str,
                           expected_value: Any, actual_value: Any,
                           target_style_name: Optional[str],
                           mapping_method: Optional[str],
                           location_detail: Optional[str] = None) -> Dict[str, Any]:
        diff_entry = {
            "paragraph_index": para_idx + 1,
            "paragraph_text_preview": normalize_text(para_text_preview, NORM_FORM='NFKC'),
            "property": property_name,
            "expected_value": str(expected_value),
            "actual_value": str(actual_value),
            "target_style_name": target_style_name or "N/A",
            "mapping_method": mapping_method or "N/A"
        }
        if location_detail:
            diff_entry["location_detail"] = location_detail
        return diff_entry

    def get_comparison_results_df(self) -> pd.DataFrame:
        if not self.differences:
            return pd.DataFrame(columns=[
                "paragraph_index", "paragraph_text_preview", "property",
                "expected_value", "actual_value", "target_style_name",
                "mapping_method", "location_detail"
            ])
        return pd.DataFrame(self.differences)

    def generate_summary_report(self) -> Dict[str, Any]:
        if not self.differences:
            return {"total_differences": 0, "differences_by_property": {}, "differences_by_style": {}}
        summary = {"total_differences": len(self.differences)}
        diff_by_prop = defaultdict(int)
        diff_by_style = defaultdict(lambda: defaultdict(int))
        for diff in self.differences:
            prop = diff["property"]
            style = diff["target_style_name"]
            diff_by_prop[prop] += 1
            diff_by_style[style][prop] += 1
        summary["differences_by_property"] = dict(diff_by_prop)
        summary["differences_by_style"] = {k: dict(v) for k, v in diff_by_style.items()}
        return summary

    def save_report_to_file(self, output_path: str, report_format: str = "json"):
        report_data = {
            "template_name": self.template_name,
            "summary": self.generate_summary_report(),
            "details": self.differences
        }
        try:
            with open(output_path, 'w', encoding='utf-8') as f:
                if report_format == "json":
                    json.dump(report_data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"错误：保存报告到 {output_path} 失败: {e}")


def add_comments_to_document_static(
    original_docx_path: str, # Path to the content (temp file)
    differences: List[Dict[str, Any]], # Differences from FormatComparatorWin32
    output_dir: str = "output_docx", # Should be relative to win32com/ or an absolute path
    suffix: str = "_commented",
    original_file_basename: Optional[str] = None # NEW parameter for original uploaded filename
) -> Optional[str]:
    if not _HAS_WIN32COM:
        print("错误 (add_comments_to_document_static): win32com 未加载，无法添加批注。")
        return None
    if not differences:
        print("信息 (add_comments_to_document_static): 未发现差异，无需生成批注文档。")
        return None

    temp_dir = None
    word_app = None
    doc = None
    coinitialized = False # Flag to track CoInitialize state
    try:
        if _HAS_WIN32COM: # Only attempt CoInitialize if win32com is available
            pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
            coinitialized = True

        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        # Use original_file_basename if provided for naming, otherwise fallback to original_docx_path's basename (temp file name)
        if original_file_basename:
            base_name_for_output = original_file_basename
        else:
            base_name_for_output = os.path.basename(original_docx_path)

        name, ext = os.path.splitext(base_name_for_output)
        commented_docx_name = f"{name}{suffix}{ext}" # e.g., <original_name>_commented.ext
        commented_docx_path = os.path.join(output_dir, commented_docx_name)

        word_app = win32com.client.DispatchEx("Word.Application")
        word_app.Visible = False
        word_app.DisplayAlerts = 0 # wdAlertsNone

        # Open the original_docx_path directly (it's the temp file with uploaded content)
        doc = word_app.Documents.Open(os.path.abspath(original_docx_path))

        comments_by_para_idx = defaultdict(list)
        for diff in differences:
            para_idx_0based = diff["paragraph_index"] - 1 # Convert 1-based from report to 0-based
            # Store the raw diff object instead of a pre-formatted string
            comments_by_para_idx[para_idx_0based].append(diff)
        
        separator = "\n" + "-" * 15 + "\n" # User-defined separator: 15 hyphens

        for para_idx_0based, paragraph_diffs in comments_by_para_idx.items():
            if para_idx_0based >= len(doc.Paragraphs):
                print(f"  警告: 差异中段落索引 {para_idx_0based + 1} 超出文档段落数 {len(doc.Paragraphs)}。跳过此批注。")
                continue
            
            win32_para_idx_1based = para_idx_0based + 1
            
            # Group diffs by property for this paragraph
            grouped_by_prop = defaultdict(lambda: {"expected": None, "actuals_with_locations": []})
            for diff_item in paragraph_diffs:
                prop_name = diff_item["property"]
                if grouped_by_prop[prop_name]["expected"] is None:
                    grouped_by_prop[prop_name]["expected"] = diff_item["expected_value"]
                
                actual_entry = {"actual": diff_item["actual_value"]}
                if diff_item.get("location_detail"): # Check if location_detail exists and is not empty
                    actual_entry["location"] = diff_item["location_detail"]
                grouped_by_prop[prop_name]["actuals_with_locations"].append(actual_entry)

            # Build the comment string
            comment_parts = [f"[格式问题 - 段落 {para_idx_0based + 1}]"]
            
            prop_keys = list(grouped_by_prop.keys())
            for i, prop_name in enumerate(prop_keys):
                prop_data = grouped_by_prop[prop_name]
                if i > 0: # Add separator before the second property onwards
                    comment_parts.append(separator.strip())
                
                comment_parts.append(f"属性: {prop_name}")
                comment_parts.append(f"  预期: {prop_data['expected']}")
                for actual_entry in prop_data["actuals_with_locations"]:
                    location_str = f" ({actual_entry['location']})" if actual_entry.get('location') else ""
                    comment_parts.append(f"  实际: {actual_entry['actual']}{location_str}")
            
            full_comment_text = "\n".join(comment_parts)

            if len(full_comment_text) > 2000: # Increased limit slightly, Word might handle more.
                full_comment_text = full_comment_text[:1990] + "...(截断)"

            try:
                para_obj = doc.Paragraphs(win32_para_idx_1based)
                target_range = para_obj.Range
                
                comment_anchor_range = doc.Range(Start=target_range.Start, End=target_range.Start)
                if target_range.Start == target_range.End:
                     comment_anchor_range = target_range
                elif hasattr(target_range, 'Characters') and target_range.Characters.Count > 0:
                    try:
                        comment_anchor_range = target_range.Characters(1).Range
                    except Exception:
                        comment_anchor_range = doc.Range(Start=target_range.Start, End=target_range.Start +1 if target_range.End > target_range.Start else target_range.Start)
                else:
                    comment_anchor_range = doc.Range(Start=target_range.Start, End=target_range.Start +1 if target_range.End > target_range.Start else target_range.Start)

                comment_obj = doc.Comments.Add(Range=comment_anchor_range, Text=full_comment_text)
                comment_obj.Author = "格式检查工具"
            except Exception as e_para:
                print(f"错误 (add_comments_to_document_static): 为段落 {win32_para_idx_1based} 添加批注时出错 - {e_para}")
                # traceback.print_exc()

        doc.SaveAs(os.path.abspath(commented_docx_path))
        return os.path.abspath(commented_docx_path)
    except Exception as e_main:
        print(f"错误 (add_comments_to_document_static): 主处理错误 - {e_main}")
        # traceback.print_exc()
        return None
    finally:
        if doc:
            try: doc.Close(SaveChanges=False)
            except Exception: pass
        if word_app:
            try: word_app.Quit()
            except Exception: pass
        # No temp_dir to remove with this simplified logic
        if coinitialized: # Only call CoUninitialize if CoInitializeEx was successful
            pythoncom.CoUninitialize()