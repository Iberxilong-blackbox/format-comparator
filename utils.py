def contains_cjk_characters(text: str) -> bool:
    """
    检查文本是否包含中日韩(CJK)字符
    
    以下Unicode范围包括：
    - CJK统一汉字 (U+4E00-U+9FFF)
    - CJK扩展A (U+3400-U+4DBF)
    - CJK扩展B (U+20000-U+2A6DF)
    - CJK兼容汉字 (U+F900-U+FAFF)
    - CJK部首扩展 (U+2E80-U+2EFF)
    - CJK笔画 (U+31C0-U+31EF)
    - 中日韩符号和标点 (U+3000-U+303F)
    - 中文标点 (U+FF00-U+FFEF部分)
    
    Args:
        text: 要检查的文本
        
    Returns:
        是否包含CJK字符
    """
    # 基本汉字和常用扩展
    if any('\u4e00' <= char <= '\u9fff' for char in text):
        return True
    
    # CJK扩展A
    if any('\u3400' <= char <= '\u4dbf' for char in text):
        return True
    
    # CJK兼容汉字
    if any('\uf900' <= char <= '\ufaff' for char in text):
        return True
        
    # CJK部首扩展
    if any('\u2e80' <= char <= '\u2eff' for char in text):
        return True
        
    # CJK笔画
    if any('\u31c0' <= char <= '\u31ef' for char in text):
        return True
        
    # 中日韩符号和标点
    if any('\u3000' <= char <= '\u303f' for char in text):
        return True
        
    # 全角ASCII、拉丁文字母和中文标点（FF00-FFEF）中的中文标点部分
    if any(('\uff01' <= char <= '\uff0f') or ('\uff1a' <= char <= '\uff20') or 
           ('\uff3b' <= char <= '\uff40') or ('\uff5b' <= char <= '\uff65') for char in text):
        return True
        
    # 其他常见中文标点
    chinese_puncts = '。，、；：？！""''（）【】《》〈〉『』「」﹃﹄〔〕…—～﹏￥'
    if any(char in chinese_puncts for char in text):
        return True
        
    return False

import json
import os
import re #确保 re 已导入
import unicodedata # 确保导入 unicodedata
from typing import List

# 数学字符规范化映射表
MATH_CHAR_NORM_MAP = {
    # Latin Mathematical Alphanumeric Symbols (subset from DocxProcessor)
    '𝜃': 'θ', '𝛿': 'δ', # '𝑎': 'a', '𝑥': 'x', # 'a' and 'x' might be too common, handle carefully or rely on NFKD
    # Italic (Lowercase)
    '𝑎': 'a', '𝑏': 'b', '𝑐': 'c', '𝑑': 'd', '𝑒': 'e', '𝑓': 'f', '𝑔': 'g', 'ℎ': 'h', '𝑖': 'i', '𝑗': 'j',
    '𝑘': 'k', '𝑙': 'l', '𝑚': 'm', '𝑛': 'n', '𝑜': 'o', '𝑝': 'p', '𝑞': 'q', '𝑟': 'r', '𝑠': 's', '𝑡': 't',
    '𝑢': 'u', '𝑣': 'v', '𝑤': 'w', '𝑥': 'x', '𝑦': 'y', '𝑧': 'z',
    # Italic (Uppercase)
    '𝐴': 'A', '𝐵': 'B', '𝐶': 'C', '𝐷': 'D', '𝐸': 'E', '𝐹': 'F', '𝐺': 'G', '𝐻': 'H', '𝐼': 'I', '𝐽': 'J',
    '𝐾': 'K', '𝐿': 'L', '𝑀': 'M', '𝑁': 'N', '𝑂': 'O', '𝑃': 'P', '𝑄': 'Q', '𝑅': 'R', '𝑆': 'S', '𝑇': 'T',
    '𝑈': 'U', '𝑉': 'V', '𝑊': 'W', '𝑋': 'X', '𝑌': 'Y', '𝑍': 'Z',
    # Bold (Lowercase)
    '𝐚': 'a', '𝐛': 'b', '𝐜': 'c', '𝐝': 'd', '𝐞': 'e', '𝐟': 'f', '𝐠': 'g', '𝐡': 'h', '𝐢': 'i', '𝐣': 'j',
    '𝐤': 'k', '𝐥': 'l', '𝐦': 'm', '𝐧': 'n', '𝐨': 'o', '𝐩': 'p', '𝐪': 'q', '𝐫': 'r', '𝐬': 's', '𝐭': 't',
    '𝐮': 'u', '𝐯': 'v', '𝐰': 'w', '𝐱': 'x', '𝐲': 'y', '𝐳': 'z',
    # Bold (Uppercase)
    '𝐀': 'A', '𝐁': 'B', '𝐂': 'C', '𝐃': 'D', '𝐄': 'E', '𝐅': 'F', '𝐆': 'G', '𝐇': 'H', '𝐈': 'I', '𝐉': 'J',
    '𝐊': 'K', '𝐋': 'L', '𝐌': 'M', '𝐍': 'N', '𝐎': 'O', '𝐏': 'P', '𝐐': 'Q', '𝐑': 'R', '𝐒': 'S', '𝐓': 'T',
    '𝐔': 'U', '𝐕': 'V', '𝐖': 'W', '𝐗': 'X', '𝐘': 'Y', '𝐙': 'Z',
    # Greek Mathematical Alphanumeric Symbols (subset from DocxProcessor)
    # Italic (Lowercase)
    '𝛼': 'α', '𝛽': 'β', '𝛾': 'γ', # '𝛿': 'δ', (already above)
    '𝜀': 'ε', '𝜁': 'ζ', '𝜂': 'η', # '𝜃': 'θ', (already above)
    '𝜄': 'ι', '𝜅': 'κ', '𝜆': 'λ', '𝜇': 'μ', '𝜈': 'ν', '𝜉': 'ξ', '𝜊': 'ο', '𝜋': 'π', '𝜌': 'ρ',
    '𝜎': 'σ', '𝜏': 'τ', '𝜐': 'υ', '𝜑': 'φ', '𝜒': 'χ', '𝜓': 'ψ', '𝜔': 'ω',
    # Italic (Uppercase)
    '𝚨': 'Α', '𝚩': 'Β', '𝚪': 'Γ', '𝚫': 'Δ', '𝚬': 'Ε', '𝚭': 'Ζ', '𝚮': 'Η', '𝚯': 'Θ', '𝚰': 'Ι', '𝚱': 'Κ',
    '𝚲': 'Λ', '𝚳': 'Μ', '𝚴': 'Ν', '𝚵': 'Ξ', '𝚶': 'Ο', '𝚷': 'Π', '𝚸': 'Ρ', '𝚺': 'Σ', '𝚻': 'Τ',
    '𝚼': 'Υ', '𝚽': 'Φ', '𝚾': 'Χ', '𝚿': 'Ψ', '𝛀': 'Ω',
    # Add more mappings as needed, e.g., for bold Greek, script, fraktur, double-struck, etc.
    # Common symbols that might differ
    '≠': '=', # Not equal to equal (for looser comparison if desired, or handle separately)
    '~': '~', # Tilde, often used in math, ensure it's consistent
    # Consider other symbols like plus, minus, dot, etc. if they have multiple Unicode representations
}


# 默认的标题样式列表，作为回退选项
DEFAULT_HEADING_STYLES = ["标题一", "标题二", "标题三", "标题四"]

def extract_heading_styles_from_template(template_path: str = "templates/default.json") -> List[str]:
    """
    从指定的模板JSON文件中提取包含“标题”的样式名称，并按级别排序。

    Args:
        template_path: 模板JSON文件的路径。

    Returns:
        排序后的标题样式名称列表。如果出错或未找到，则返回默认列表的副本。
    """
    heading_styles = []
    try:
        # 检查模板文件是否存在
        if not os.path.exists(template_path):
            print(f"警告: 模板配置文件 {template_path} 未找到。将使用默认列表。")
            return DEFAULT_HEADING_STYLES[:] # 返回副本

        with open(template_path, 'r', encoding='utf-8') as f:
            template_config = json.load(f)

            # 检查 JSON 结构是否符合预期
            if "样式" not in template_config or not isinstance(template_config["样式"], dict):
                print(f"警告: 在 {template_path} 中未找到有效的 '样式' 字典。将使用默认列表。")
                return DEFAULT_HEADING_STYLES[:] # 返回副本

            # 提取包含“标题”的样式名称
            heading_styles = [
                style_name for style_name in template_config["样式"]
                if isinstance(style_name, str) and "标题" in style_name
            ]

            # 如果没有提取到任何标题样式
            if not heading_styles:
                print(f"警告: 未能从 {template_path} 提取到任何包含'标题'的样式。将使用默认列表。")
                return DEFAULT_HEADING_STYLES[:] # 返回副本

            # 定义排序函数，按标题级别排序
            def get_heading_level(style_name):
                # 尝试匹配中文数字 "标题一", "标题二", ...
                match_cn = re.search(r'标题([一二三四五六七八九])', style_name)
                if match_cn:
                    num_map = {'一': 1, '二': 2, '三': 3, '四': 4, '五': 5, '六': 6, '七': 7, '八': 8, '九': 9}
                    return num_map.get(match_cn.group(1), 99) # 默认值设为较大数

                # 尝试匹配阿拉伯数字 "标题1", "标题2", ...
                match_num = re.search(r'标题(\d+)', style_name)
                if match_num:
                    try:
                        return int(match_num.group(1))
                    except ValueError:
                        return 99 # 如果数字转换失败

                # 如果两种模式都匹配不上，返回一个较大的默认值，使其排在后面
                return 99

            # 应用排序
            heading_styles.sort(key=get_heading_level)

    except json.JSONDecodeError:
        print(f"警告: 模板配置文件 {template_path} 格式错误。将使用默认列表。")
        return DEFAULT_HEADING_STYLES[:] # 返回副本
    except Exception as e:
        # 捕获其他可能的异常，例如读取文件时的权限问题等
        print(f"警告: 从模板提取标题样式时发生意外错误: {e}。将使用默认列表。")
        return DEFAULT_HEADING_STYLES[:] # 返回副本

    # 最终返回提取并排序后的列表
    return heading_styles

# --- Functions moved from format_comparator.py ---

def _number_to_chinese(num: int) -> str:
    """将数字 1-9 转换为中文数字。"""
    chinese_map = {1: "一", 2: "二", 3: "三", 4: "四", 5: "五", 6: "六", 7: "七", 8: "八", 9: "九"}
    return chinese_map.get(num, str(num))


def _is_primarily_east_asian(text: str) -> bool:
    """Checks if the text contains a significant portion of East Asian characters."""
    if not text:
        return False
    east_asian_count = 0
    total_count = 0
    # Basic check for CJK Unified Ideographs, Hangul Syllables, Hiragana, Katakana
    # More comprehensive ranges can be added if needed.
    for char in text:
        total_count += 1
        # CJK Unified Ideographs U+4E00 to U+9FFF
        if '\u4e00' <= char <= '\u9fff':
            east_asian_count += 1
        # Hangul Syllables U+AC00 to U+D7AF
        elif '\uac00' <= char <= '\ud7af':
            east_asian_count += 1
        # Hiragana U+3040 to U+309F
        elif '\u3040' <= char <= '\u309f':
            east_asian_count += 1
        # Katakana U+30A0 to U+30FF
        elif '\u30a0' <= char <= '\u30ff':
            east_asian_count += 1

    # Heuristic: If more than 30% are East Asian characters, consider it primarily East Asian
    # This threshold might need adjustment based on typical content.
    return total_count > 0 and (east_asian_count / total_count) > 0.3


def normalize_text(text: str, NORM_FORM='NFKC') -> str: # Added NORM_FORM parameter
    """
    规范化文本以便比较：
    1. 确保是字符串。
    2. (可选) Unicode 规范化 (例如 NFKC)。
    3. 对数学字母数字符号等进行自定义规范化。
    4. 移除特殊的 "[FORMULA:]" 标记及其两侧可能的空格。
    5. 将所有类型的换行符统一为 '\n'。
    6. 将连续的内部空白字符（包括空格、制表符、换行符）替换为单个空格。
    7. 去除首尾空白。
    """
    if not isinstance(text, str):
        text = str(text)

    # 1. (可选) Unicode 规范化
    if NORM_FORM:
        text = unicodedata.normalize(NORM_FORM, text)

    # 2. 字符规范化 (自定义数学符号等)
    normalized_char_list = []
    for char_in_text in text:
        normalized_char_list.append(MATH_CHAR_NORM_MAP.get(char_in_text, char_in_text))
    text = "".join(normalized_char_list)

    # 3. 移除 [FORMULA:] 标记
    text = text.replace(" [FORMULA:] ", " ")
    text = text.replace("[FORMULA:] ", " ")
    text = text.replace(" [FORMULA:]", " ")
    text = text.replace("[FORMULA:]", "")

    # 4. 统一换行符
    text = text.replace('\r\n', '\n').replace('\r', '\n')
    
    # 5. 合并连续空白（包括因上述替换产生的多余空格）
    text = re.sub(r'\s+', ' ', text)
    
    # 6. 最后去除首尾空白
    return text.strip()