import os
import re
import json
import ast  # 添加此导入用于解析Python字面量
from typing import List, Dict, Any, Optional, Union, Tuple
from openai import OpenAI, Timeout, APITimeoutError # 导入 APITimeoutError
import logging # 导入日志模块
import pandas as pd # Added for type hinting

# --- 配置日志记录器 ---
# Changed logger name slightly for clarity and to avoid conflicts if root logger is used elsewhere
logger = logging.getLogger('llm_parser_errors_win32com')
# 避免重复添加 handler
if not logger.hasHandlers():
    logger.setLevel(logging.WARNING) # 设置日志级别
    # 创建文件处理器，指定文件名和编码
    # Log file will be created in the win32com directory as llm_mapper.py is there
    file_handler = logging.FileHandler('llm_parsing_errors.log', encoding='utf-8')
    # 创建日志格式器
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    # 将格式器添加到处理器
    file_handler.setFormatter(formatter)
    # 将处理器添加到 logger
    logger.addHandler(file_handler)
# --- 日志配置结束 ---

def load_config():
    """
    从配置文件加载设置 (win32com version)
    
    Returns:
        Dict: 配置信息
    """
    current_dir = os.path.dirname(__file__)
    config_path = os.path.join(current_dir, "config.json")
    example_config_path = os.path.join(current_dir, "config_example.json")
    
    try:
        with open(config_path, "r", encoding="utf-8") as f:
            return json.load(f)
    except FileNotFoundError:
        try:
            with open(example_config_path, "r", encoding="utf-8") as f:
                config_data = json.load(f)
                print(f"警告: 使用示例配置文件 ({example_config_path})。请创建 {config_path} 并填入您的 API 密钥。")
                return config_data
        except FileNotFoundError:
            print(f"警告: 找不到配置文件 ({config_path} 或 {example_config_path})，使用默认配置。")
            return {
                "llm": {
                    "provider": "deepseek",
                    "api_key": "<YOUR_API_KEY>",
                    "base_url": "https://api.deepseek.com",
                    "model": "deepseek-chat",
                    "max_tokens": 2048,
                    "response_format": {"type": "json_object"} # Kept as per original, though plan says text
                }
            }
    except json.JSONDecodeError:
        print(f"警告: 配置文件 {config_path} 格式错误，使用默认配置。")
        return {
            "llm": {
                "provider": "deepseek",
                "api_key": "<YOUR_API_KEY>",
                "base_url": "https://api.deepseek.com",
                "model": "deepseek-chat",
                "max_tokens": 2048,
                "response_format": {"type": "json_object"} # Kept as per original
            }
        }

# 加载主配置
config = load_config()

# 创建 OpenAI 客户端实例
def create_llm_client(): # Name kept as per plan
    """
    根据配置创建 LLM 客户端 (win32com version)
    
    Returns:
        OpenAI: LLM 客户端实例
    """
    llm_config = config.get("llm", {})
    api_key = llm_config.get("api_key")
    base_url = llm_config.get("base_url")
    
    # More robust check for placeholder API keys
    if not api_key or api_key.startswith("<YOUR_") or api_key == "<DeepSeek API Key>":
        # Construct the expected path to config.json within the win32com directory
        expected_config_path = os.path.join(os.path.dirname(__file__), 'config.json')
        print(f"警告: 请在 {expected_config_path} 中设置您的 API 密钥 (llm.api_key)")
        return None
    
    return OpenAI(api_key=api_key, base_url=base_url)

class LLMStyleMapper: # Renamed from LLMStyleMapperGenerator
    """
    基于 LLM 的 Word 文档样式映射生成器 (win32com version)
    
    该类使用大型语言模型（LLM）分析 Word 文档内容，
    自动生成段落样式映射，提高样式应用的效率和准确性。
    """
    def __init__(self, template_manager, llm_client=None, split_titles: Optional[List[str]] = None, template_data: Optional[dict] = None):
        """初始化样式映射生成器 (win32com version)"""
        self.template_manager = template_manager # This will be TemplateManagerWin32 instance
        self.llm_client = llm_client
        self.config = config
        self.split_titles = split_titles if split_titles is not None else []
        self.template_data = template_data

    # Note: _segment_document, _find_title_indices_in_body might not be directly used
    # if generate_mapping directly receives a filtered doc_df.
    # Keeping them for now in case a more complex batching strategy is re-introduced.

    def _is_title_match(self, paragraph_text: str, config_title: str) -> bool:
        """
        检查段落文本是否匹配配置的标题，使用文本规范化和增强的正则表达式。
        (win32com version - uses normalize_text from .utils)
        """
        from .utils import normalize_text # Relative import for win32com
        normalized_para_text = normalize_text(paragraph_text)
        normalized_config_title = normalize_text(config_title)

        if not normalized_para_text or not normalized_config_title:
            return False
        try:
            core_pattern_part = re.escape(normalized_config_title).replace(r'\ ', r'\s+')
            numbering_prefix_pattern = r'(?:第?\s*[一二三四五六七八九十百千万亿\d]+\s*[、．.]?\s*)?'
            title_pattern_str = (
                r'^\s*' +
                numbering_prefix_pattern +
                core_pattern_part +
                r'\s*$'
            )
            title_pattern = re.compile(title_pattern_str, re.IGNORECASE)
            match = title_pattern.match(normalized_para_text)
            return bool(match)
        except Exception as e:
            print(f"警告: 在构建或匹配标题 '{config_title}' 的正则表达式时出错: {e}")
            return False

    def _extract_paragraphs_from_df(self, doc_df: pd.DataFrame) -> List[Dict[str, Any]]: # Renamed and signature confirmed
        """
        从 DataFrame 中提取段落索引和文本 (win32com version)。
        Args:
            doc_df: Pandas DataFrame containing paragraph data.
        Returns:
            List of dictionaries with 'idx' and 'text'.
        """
        paragraphs = []
        if 'paragraph_index' not in doc_df.columns or 'text' not in doc_df.columns:
            # Log this error or raise a more specific one if critical
            print("错误: _extract_paragraphs_from_df 期望的 DataFrame 缺少 'paragraph_index' 或 'text' 列。")
            return [] # Return empty list to prevent further errors downstream
            
        for _, row in doc_df.iterrows():
            paragraphs.append({
                "idx": row['paragraph_index'],
                "text": str(row['text']) # Ensure text is string
            })
        return paragraphs
    
    def _preprocess_paragraphs(self, paragraphs: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """
        对段落进行预处理，识别明显的格式特征并生成初步样式建议
        """
        processed = []
        for para in paragraphs:
            idx = para["idx"]
            text = str(para["text"]) # Ensure text is string
            result = {"idx": idx, "text": text}
            
            heading_match = re.match(r'^(\d+(\.\d+)*)\s+(.+)$', text)
            if heading_match:
                number = heading_match.group(1)
                level = number.count('.') + 1
                if level == 1: result["suggestion"] = "标题一"
                elif level == 2: result["suggestion"] = "标题二"
                elif level >= 3: result["suggestion"] = "标题三"
            elif re.match(r'^[•\-*]\s+(.+)$', text): result["suggestion"] = "列表项"
            elif re.match(r'^(\d+|[a-zA-Z]+|[ivxIVX]+)[\.、\)）]\s+(.+)$', text): result["suggestion"] = "列表项"
            elif re.match(r'^[>》]\s*(.+)$', text) or text.startswith('"') and text.endswith('"'): result["suggestion"] = "引用"
            processed.append(result)
        return processed
    
    def _build_llm_prompt(self, paragraphs: List[Dict[str, Any]], template_styles: Optional[List[str]] = None) -> Dict[str, Any]:
        """
        构建发送给 LLM 的 Prompt
        """
        system_prompt = """你是一个专业的论文文档样式分析助手。你会收到一个 JSON 列表，其中每个对象代表原始文档中的一个**非空文本段落**。每个对象包含：
- `idx`: 该段落在**原始完整文档**中的**绝对索引号**（从0开始计数）。请注意，由于只发送了非空段落，这些 `idx` **可能不是连续的**。
- `text`: 段落的文本内容。
- `suggestion` (可选): 基于规则预处理得出的初步样式建议。

可用的样式类型包括（请优先从提供的模板样式列表中选择，如果列表非空）：
- 标题一
- 标题二
- 标题三
- 标题四
- 正文
- 图题
- 表题
- 公式
- 参考文献
- 附录标题
- 页眉
- 页脚
- 目录标题
- 目录1
- 目录2
- 目录3

你的任务是为你收到的**每一个**段落（由其唯一的 `idx` 标识）判断最合适的样式 `style`。

- 对于包含 `suggestion` 的段落：请将该建议作为重要参考，结合上下文语义进行验证。如果建议合理，请采纳；如果认为建议不准确，请根据你的判断给出更合适的样式。
- 对于不包含 `suggestion` 的段落：请完全基于文本内容和上下文语义进行判断。

**输出格式要求：**
请严格按照以下格式输出每一对映射关系，**每对占一行**：
`段落索引号,样式名称`

例如：
`0,标题一`
`6,正文`
`7,正文`
`23,图题`

**重要提示：**
- **不要**包含任何额外的括号、引号、分号、Markdown 标记 (如 ```) 或其他任何无关字符。
- **确保**只输出索引号、一个英文逗号、样式名称和换行符。
- **必须**为你收到的**每一个**段落（及其对应的原始 `idx`）都输出一行对应的映射。"""

        if template_styles: # template_styles are unprefixed names
            style_list_str = ", ".join(sorted(list(set(template_styles)))) # Ensure unique and sorted for consistency
            system_prompt += f"\n\n请注意，你主要应该从以下模板样式列表中选择：{style_list_str}。如果这些样式都不适用，可以考虑上述通用样式类型。"
        
        user_prompt = json.dumps(paragraphs, ensure_ascii=False, indent=2)
        
        return {
            "system": system_prompt,
            "user": user_prompt
        }
    
    def _call_llm(self, prompt: Dict[str, Any], timeout: int = 120, mock_response_for_testing: Optional[str] = None) -> str:
        """调用 LLM API"""
        if mock_response_for_testing is not None:
            print("信息: 使用测试提供的模拟LLM响应。")
            return mock_response_for_testing

        if self.llm_client is None:
            print("LLM 客户端未初始化，将使用内置模拟响应。")
            return self._mock_llm_response(prompt)

        try:
            llm_params = self.config.get("llm", {})
            model = llm_params.get("model", "deepseek-chat")
            max_tokens = llm_params.get("max_tokens", 4096) # Increased default
            top_p = llm_params.get("top_p", 0.7) # Example value, can be configured
            temperature = llm_params.get("temperature", 0.1) # Example value for more deterministic output

            response = self.llm_client.chat.completions.create(
                model=model,
                messages=[
                    {"role": "system", "content": prompt["system"]},
                    {"role": "user", "content": prompt["user"]}
                ],
                max_tokens=max_tokens,
                top_p=top_p,
                temperature=temperature,
                timeout=Timeout(float(timeout))
            )
            if isinstance(response, str):
                # Log the first 500 characters of the unexpected string response
                unexpected_response_preview = response[:500]
                error_message = f"LLM API returned an unexpected string response. Preview: {unexpected_response_preview}"
                print(f"错误: {error_message}") # Log for immediate visibility
                raise RuntimeError(error_message) # Raise a more specific error
            return response.choices[0].message.content
        except APITimeoutError:
            print(f"错误: 调用 LLM API 超时 (超过 {timeout} 秒)")
            raise RuntimeError(f"调用 LLM API 超时 (超过 {timeout} 秒)")
        except Exception as e:
            print(f"错误: 调用 LLM API 时出错: {e}")
            raise RuntimeError(f"调用 LLM API 时出错: {e}")
    
    def _mock_llm_response(self, prompt: Dict[str, Any]) -> str:
        """生成模拟的 LLM 响应（用于测试）"""
        try:
            paragraphs = json.loads(prompt["user"])
            response_lines = []
            for para in paragraphs:
                idx = para["idx"]
                text = str(para["text"])
                style = para.get("suggestion", "正文") 
                if len(text) < 20 and idx == 0 and "suggestion" not in para : style = "标题一"
                elif len(text) < 20 and "suggestion" not in para : style = "标题二"
                response_lines.append(f"{idx},{style}")
            return "\n".join(response_lines)
        except Exception:
            return "0,标题一\n1,正文" 
    
    def _parse_llm_response(self, response: str) -> List[Dict[str, Any]]:
        """解析 LLM 返回的简单 CSV 格式 (idx,style) 响应"""
        logger_instance = logging.getLogger('llm_parser_errors_win32com')
        processed_response = re.sub(r'^\s*```[a-zA-Z]*\n', '', response.strip(), flags=re.MULTILINE)
        processed_response = re.sub(r'\n```\s*$', '', processed_response.strip(), flags=re.MULTILINE)
        processed_response = processed_response.strip()

        parsed_mappings = []
        lines = processed_response.splitlines()

        for i, line in enumerate(lines):
            original_line_for_debug = line # Keep original for debug
            line = line.strip()
            if not line: continue
            
            print(f"DEBUG _parse_llm_response: Processing line {i+1}/{len(lines)}: '{original_line_for_debug}' (stripped: '{line}')")
            
            parts = line.split(',', 1)
            print(f"DEBUG _parse_llm_response: Parts after split: {parts}")

            if len(parts) != 2:
                warning_msg = f"解析行 {i+1} 失败 - 格式无效: '{line}'"
                print(f"警告: {warning_msg}")
                logger_instance.warning(f"{warning_msg}\n--- Raw LLM Response ---\n{response}\n--- End Raw Response ---")
                continue
                
            idx_str, style_name_raw = parts[0].strip(), parts[1]
            print(f"DEBUG _parse_llm_response: idx_str='{idx_str}', style_name_raw='{style_name_raw}'")
            
            # Thoroughly clean style_name, removing leading/trailing whitespace and carriage returns
            style_name = style_name_raw.replace('\r', '').strip()
            print(f"DEBUG _parse_llm_response: Cleaned style_name='{style_name}'")

            # Normalize Chinese numeral titles to Arabic numerals for direct matching
            normalized_style_name = style_name
            if style_name == "标题一":
                normalized_style_name = "标题1"
            elif style_name == "标题二":
                normalized_style_name = "标题2"
            
            if normalized_style_name != style_name:
                print(f"DEBUG _parse_llm_response: Normalized style_name from '{style_name}' to '{normalized_style_name}'")
                style_name = normalized_style_name
            
            try:
                idx = int(idx_str)
            except ValueError:
                warning_msg = f"解析行 {i+1} 失败 - 索引不是有效整数: '{idx_str}' in line '{line}'"
                print(f"警告: {warning_msg}")
                logger_instance.warning(f"{warning_msg}\n--- Raw LLM Response ---\n{response}\n--- End Raw Response ---")
                continue
            if not style_name:
                 warning_msg = f"解析行 {i+1} 失败 - 样式名称为空: '{line}'"
                 print(f"警告: {warning_msg}")
                 logger_instance.warning(f"{warning_msg}\n--- Raw LLM Response ---\n{response}\n--- End Raw Response ---")
                 continue
            parsed_mappings.append({"idx": idx, "style": style_name})
        print(f"成功解析了 {len(parsed_mappings)} 个映射。")
        return parsed_mappings
    
    def _map_styles_to_template(self, llm_mappings: List[Dict[str, Any]], template_name: str) -> List[Dict[str, Any]]:
        """将 LLM 返回的样式名映射到指定模板的样式名"""
        template_content = None
        base_template_style_names = [] # Unprefixed names from template_content['样式']['样式']
        valid_prefixed_style_names_in_template = [] # Prefixed names that are valid in the template
        template_prefix = ''

        if self.template_data:
            template_content = self.template_data
        else:
            print(f"警告: _map_styles_to_template 未收到 template_data。尝试从 manager 加载 '{template_name}'。")
            try:
                template_content = self.template_manager.load_template_json(template_name=template_name)
            except Exception as e:
                print(f"警告: 通过 template_manager 加载模板 '{template_name}' 时出错: {e}")

        if not template_content:
            print(f"错误: 无法获取模板 '{template_name}' 的内容，无法进行样式映射。")
            return []

        # Extract prefix and base style names from the correct nested structure
        styles_level1 = template_content.get('样式', {})
        if isinstance(styles_level1, dict):
            template_prefix = styles_level1.get('prefix', '')
            actual_styles_dict = styles_level1.get('样式', {})
            if isinstance(actual_styles_dict, dict):
                raw_base_names = list(actual_styles_dict.keys())
                for raw_base_name in raw_base_names:
                    clean_base_name = raw_base_name.strip()
                    if clean_base_name: # Ensure not empty after stripping
                        base_template_style_names.append(clean_base_name)
                        valid_prefixed_style_names_in_template.append(f"{template_prefix}{clean_base_name}")
            else:
                print(f"警告: 模板 '{template_name}' 中的 '样式.样式' 不是预期的字典格式。")
        else:
            print(f"警告: 模板 '{template_name}' 中的 '样式' 不是预期的字典格式。")

        if not base_template_style_names:
            print(f"警告: 未能从模板 '{template_name}' 的 '样式.样式' 结构中获取基础样式列表。")
            # Fallback: try to get styles from the old location if new one is empty, for backward compatibility or misconfiguration
            if isinstance(styles_level1, dict) and not actual_styles_dict: # Check if '样式.样式' was empty but '样式' had keys
                 # This might indicate an older template format or a misconfiguration
                 # For now, we'll assume the new structure is authoritative.
                 pass


        if not valid_prefixed_style_names_in_template:
            print(f"警告: 未能从模板 '{template_name}' 构建有效的带前缀样式列表。LLM映射可能不准确。")
            # If absolutely no valid styles, we can't map.
            # However, the prompt to LLM would have also been empty, so LLM might return generic styles.
        
        print(f"DEBUG _map_styles_to_template: template_prefix = '{template_prefix}'")
        print(f"DEBUG _map_styles_to_template: base_template_style_names = {base_template_style_names}")
        print(f"DEBUG _map_styles_to_template: valid_prefixed_style_names_in_template = {valid_prefixed_style_names_in_template}")

        mapped_styles_output = []
        default_fallback_prefixed_style = f"{template_prefix}正文"
        if default_fallback_prefixed_style not in valid_prefixed_style_names_in_template:
            # If "自定义正文" is not valid, use the first valid style as fallback, or an empty string if none.
            default_fallback_prefixed_style = valid_prefixed_style_names_in_template[0] if valid_prefixed_style_names_in_template else ""
        print(f"DEBUG _map_styles_to_template: default_fallback_prefixed_style = '{default_fallback_prefixed_style}'")

        for item in llm_mappings:
            llm_idx = item["idx"]
            llm_style_name_unprefixed = item["style"] # LLM is prompted to return unprefixed names
            
            final_style_to_apply = ""

            # 1. Attempt direct match with prefix
            llm_style_name_unprefixed = item["style"] # Ensure this is defined before use in print
            print(f"\nDEBUG _map_styles_to_template: Processing item for idx={llm_idx}, llm_style_unprefixed='{llm_style_name_unprefixed}'")
            potential_direct_match_prefixed = f"{template_prefix}{llm_style_name_unprefixed}"
            print(f"DEBUG _map_styles_to_template: Attempting direct match with '{potential_direct_match_prefixed}'")
            
            is_direct_match = potential_direct_match_prefixed in valid_prefixed_style_names_in_template
            print(f"DEBUG _map_styles_to_template: Direct match result for '{potential_direct_match_prefixed}': {is_direct_match}")

            if is_direct_match:
                final_style_to_apply = potential_direct_match_prefixed
            else:
                print(f"DEBUG _map_styles_to_template: Direct match failed. Trying fuzzy match for '{llm_style_name_unprefixed}'.")
                # 2. If direct match fails, try fuzzy matching against base (unprefixed) template style names
                if base_template_style_names:
                    best_match_unprefixed = self._find_best_match_unprefixed(llm_style_name_unprefixed, base_template_style_names)
                    print(f"DEBUG _map_styles_to_template: Fuzzy best_match_unprefixed='{best_match_unprefixed}'")
                    potential_fuzzy_match_prefixed = f"{template_prefix}{best_match_unprefixed}"
                    print(f"DEBUG _map_styles_to_template: Attempting fuzzy match with '{potential_fuzzy_match_prefixed}'")
                    
                    is_fuzzy_match_valid = potential_fuzzy_match_prefixed in valid_prefixed_style_names_in_template
                    print(f"DEBUG _map_styles_to_template: Fuzzy match result for '{potential_fuzzy_match_prefixed}': {is_fuzzy_match_valid}")

                    if is_fuzzy_match_valid:
                        final_style_to_apply = potential_fuzzy_match_prefixed
                    else:
                        print(f"警告: LLM样式 '{llm_style_name_unprefixed}' (段落 {llm_idx}) 直接和模糊匹配均失败。模糊匹配结果 '{best_match_unprefixed}' (带前缀 '{potential_fuzzy_match_prefixed}') 无效。回退到默认。")
                        final_style_to_apply = default_fallback_prefixed_style
                else: # No base styles for fuzzy matching
                     print(f"警告: LLM样式 '{llm_style_name_unprefixed}' (段落 {llm_idx}) 直接匹配失败，且无基础样式进行模糊匹配。回退到默认。")
                     final_style_to_apply = default_fallback_prefixed_style
            
            print(f"DEBUG _map_styles_to_template: Style to apply before final fallback check: '{final_style_to_apply}' for idx={llm_idx}")
            if not final_style_to_apply and valid_prefixed_style_names_in_template:
                print(f"警告: LLM样式 '{llm_style_name_unprefixed}' (段落 {llm_idx}) 最终未能映射到有效样式，且默认回退也为空。将使用模板中第一个有效样式: '{valid_prefixed_style_names_in_template[0]}'")
                final_style_to_apply = valid_prefixed_style_names_in_template[0]
            elif not final_style_to_apply: # This case means valid_prefixed_style_names_in_template is also empty
                 print(f"错误: LLM样式 '{llm_style_name_unprefixed}' (段落 {llm_idx}) 无法映射，且模板中无有效样式可回退。跳过此映射。")
                 continue # Skip this mapping

            print(f"DEBUG _map_styles_to_template: Final style for idx={llm_idx} is '{final_style_to_apply}'")
            mapped_styles_output.append({"paragraph_index": llm_idx, "style": final_style_to_apply})
        return mapped_styles_output
    
    def _find_best_match_unprefixed(self, llm_style_name: str, unprefixed_template_styles: List[str]) -> str:
        """使用模糊匹配找到最匹配的无前缀样式名"""
        if not unprefixed_template_styles: return "正文" # Default if no styles to match against

        try:
            from fuzzywuzzy import fuzz 
            best_match = "正文" # Default
            highest_similarity = 0
            for style_option in unprefixed_template_styles:
                similarity = fuzz.ratio(style_option.lower(), llm_style_name.lower())
                if similarity > highest_similarity:
                    highest_similarity = similarity
                    best_match = style_option
            if highest_similarity > 75: # Adjusted threshold
                return best_match
        except ImportError:
            # Fallback if fuzzywuzzy is not available
            if llm_style_name in unprefixed_template_styles:
                return llm_style_name
        return "正文" # Ultimate fallback
    
    def generate_mapping(self,
                         doc_df: pd.DataFrame,
                         middle_start_index: Optional[int], # Added
                         back_start_index: Optional[int],
                         template_name: str,
                         max_retries: int = 1,
                         context_size: int = 10,
                         mock_llm_response_str_for_testing: Optional[str] = None) -> List[Dict[str, Any]]: # Reduced retries
        """
       生成样式映射 (win32com version)。
       Args:
            doc_df: Pandas DataFrame containing paragraph data.
            middle_start_index: Index where the body matter of the document starts.
            back_start_index: Index where the back matter of the document starts.
            template_name: Name of the template to map against.
            max_retries: Max retries for LLM calls.
            context_size: Context size for retries.
            mock_llm_response_str_for_testing: Optional string to use as a direct LLM response for testing.
        Returns:
            List of style mappings.
        """
        print(f"开始生成样式映射 (win32com)，模板: {template_name}")
        if mock_llm_response_str_for_testing:
            print("信息: 本次 generate_mapping 调用将使用外部提供的模拟LLM响应。")

        try:
            all_paragraphs_data = self._extract_paragraphs_from_df(doc_df)
        except ValueError as e:
            print(f"错误: 从DataFrame提取段落时出错: {e}")
            return []
        except Exception as e_general:
            print(f"错误: 提取段落时发生未知错误: {e_general}")
            return []

        paragraphs_to_process = []
        # Determine the actual start and end for processing
        actual_start_index = middle_start_index if middle_start_index is not None else 0
        actual_end_index = back_start_index # This can be None

        for para in all_paragraphs_data:
            para_idx = para['idx']
            # Condition: paragraph index must be >= actual_start_index
            # AND (actual_end_index is None OR paragraph index < actual_end_index)
            if para_idx >= actual_start_index and \
               (actual_end_index is None or para_idx < actual_end_index):
                paragraphs_to_process.append(para)

        if middle_start_index is not None:
            print(f"LLM 映射将从段落索引 {actual_start_index} 开始。")
        if back_start_index is not None:
            print(f"LLM 映射将处理到段落索引 {actual_end_index - 1} 为止。")
        else:
            print(f"LLM 映射将处理到文档末尾（从索引 {actual_start_index} 开始）。")
        
        print(f"筛选后，LLM 将处理 {len(paragraphs_to_process)} 个段落。")

        if not paragraphs_to_process:
            print("筛选后没有需要 LLM 处理的段落。")
            return []

        preprocessed_paragraphs = self._preprocess_paragraphs(paragraphs_to_process)
        
        current_batch = preprocessed_paragraphs # Process all in one batch for now
        all_llm_raw_mappings = [] 

        template_styles_for_prompt = []
        unprefixed_template_styles = [] # For _map_styles_to_template
        
        # Ensure template_data is loaded if not already available
        if not self.template_data:
            print(f"警告: LLMStyleMapper 初始化时未提供 template_data。尝试从 manager 加载 '{template_name}'。")
            try:
                # self.template_manager is TemplateManagerWin32 instance
                self.template_data = self.template_manager.load_template_json(template_name=template_name)
                if not self.template_data:
                    print(f"错误: 无法加载模板 '{template_name}' 数据。无法继续生成 LLM 映射。")
                    return []
            except Exception as e_load:
                print(f"错误: 加载模板 '{template_name}' 数据时出错: {e_load}。无法继续。")
                return []


        # Correctly extract base style names for the LLM prompt
        if self.template_data:
            styles_level1_for_prompt = self.template_data.get('样式', {})
            if isinstance(styles_level1_for_prompt, dict):
                actual_styles_dict_for_prompt = styles_level1_for_prompt.get('样式', {})
                if isinstance(actual_styles_dict_for_prompt, dict):
                    # These are the unprefixed names LLM should work with
                    # Ensure these are also cleaned
                    cleaned_style_names_for_prompt = []
                    for style_key in actual_styles_dict_for_prompt.keys():
                        cleaned_key = style_key.strip()
                        if cleaned_key:
                            cleaned_style_names_for_prompt.append(cleaned_key)
                    
                    template_styles_for_prompt = cleaned_style_names_for_prompt
                    # unprefixed_template_styles is used by _map_styles_to_template's fuzzy matching,
                    # and should be the same list of cleaned names.
                    unprefixed_template_styles = cleaned_style_names_for_prompt
            
        if not template_styles_for_prompt:
            print("警告: 未能从模板提取样式列表以供LLM提示。LLM可能无法准确映射。")
        
        prompt = self._build_llm_prompt(current_batch, template_styles_for_prompt)
        
        llm_response_str = ""
        for attempt in range(max_retries + 1):
            try:
                llm_response_str = self._call_llm(prompt, mock_response_for_testing=mock_llm_response_str_for_testing)
                break
            except RuntimeError as e:
                print(f"LLM 调用失败 (尝试 {attempt + 1}/{max_retries + 1}): {e}")
                if attempt == max_retries:
                    print("达到最大重试次数，LLM 映射生成失败。")
                    return [] 
            except Exception as e_call:
                print(f"LLM 调用时发生未知错误 (尝试 {attempt + 1}): {e_call}")
                if attempt == max_retries: return []

        if not llm_response_str:
            print("LLM 未返回有效响应。")
            return []

        parsed_batch_mappings = self._parse_llm_response(llm_response_str)
        all_llm_raw_mappings.extend(parsed_batch_mappings)

        # Pass unprefixed_template_styles to _map_styles_to_template if that's what it expects
        # Or adjust _map_styles_to_template to handle prefixed styles from self.template_data
        final_mapped_styles = self._map_styles_to_template(all_llm_raw_mappings, template_name) # template_name is used if self.template_data is None
        
        # Optional: Save mapping
        # self.save_mapping_to_file(final_mapped_styles, f"user_files/debug/debug_llm_output_{template_name}.json")

        print(f"LLM 样式映射生成完成，共处理 {len(final_mapped_styles)} 个段落的映射。")
        return final_mapped_styles

    def save_mapping_to_file(self, mappings: List[Dict[str, Any]], output_path: str) -> None:
        """
        将生成的样式映射保存到 JSON 文件。
        """
        try:
            output_dir = os.path.dirname(output_path)
            if output_dir and not os.path.exists(output_dir):
                os.makedirs(output_dir)
            with open(output_path, 'w', encoding='utf-8') as f:
                json.dump(mappings, f, ensure_ascii=False, indent=2)
            print(f"样式映射已保存到: {output_path}")
        except Exception as e:
            print(f"错误: 保存样式映射到文件时出错: {e}")

# --- Main block for testing (win32com version) ---
if __name__ == '__main__':
    print("LLMStyleMapper (win32com version) - 测试模块运行中...")
    
    from template_manager_win32 import TemplateManagerWin32 # Placed import here
    from pathlib import Path # Placed import here
    # import pandas as pd # Already imported at the top

    # 1. 定义模拟LLM回复
    # 假设模板 "2_20250520202508202667" 的前缀是 "自定义"
    # 模板中定义的无前缀样式（根据您提供的模板JSON）:
    # "正文", "标题1", "标题2", "标题3", "图题", "表题"
    mock_llm_output_str = """0,标题1
1,正文
2,图题
3,表格标题
4,一个不存在的样式
5,参考文献
6,标提2
7,页眉
8,目录1
10,标题3
11,表题
"""
    # 解释预期结果 (基于模板 "2_20250520202508202667" 前缀 "自定义"):
    # 0,标题1         -> 期望: 自定义标题1 (直接匹配)
    # 1,正文          -> 期望: 自定义正文 (直接匹配)
    # 2,图题          -> 期望: 自定义图题 (直接匹配)
    # 3,表格标题       -> 模糊匹配 "表题", 期望: 自定义表题
    # 4,一个不存在的样式 -> 回退, 期望: 自定义正文 (因为 "自定义正文" 在模板中有效)
    # 5,参考文献       -> 模板中无此样式, 回退, 期望: 自定义正文
    # 6,标提2         -> 模糊匹配 "标题2", 期望: 自定义标题2
    # 7,页眉          -> 模板中无此样式, 回退, 期望: 自定义正文
    # 8,目录1         -> 模板中无此样式, 回退, 期望: 自定义正文
    # 10,标题3        -> 期望: 自定义标题3 (直接匹配)
    # 11,表题         -> 期望: 自定义表题 (直接匹配)

    # 2. 创建模拟 DataFrame
    mock_para_indices = []
    try:
        mock_para_indices = [int(line.split(',')[0]) for line in mock_llm_output_str.strip().split('\n') if line.strip()]
    except ValueError as e:
        print(f"解析模拟LLM输出中的索引时出错: {e}")
        print("请检查 mock_llm_output_str 格式是否为 '索引,样式名'")
        exit()
    
    mock_data = {
        'paragraph_index': mock_para_indices,
        'text': [f"这是段落 {i} 的模拟文本。" for i in mock_para_indices],
    }
    mock_df = pd.DataFrame(mock_data)

    # 3. 设置模板名称和加载模板数据
    # 根据数据库输出，Name 字段存储的是 '2'，而不是完整的文件名（不含.json）
    mock_template_name = "2"
    script_dir = Path(__file__).parent
    user_files_dir = script_dir / "user_files"
    
    print(f"测试模块将使用模板: {mock_template_name}")
    print(f"预期的模板文件路径: {user_files_dir / 'templates' / (mock_template_name + '.json')}")

    tm = TemplateManagerWin32(base_user_dir=user_files_dir)
    
    # Print all available templates from DB for debugging
    print("\n--- 可用模板列表 (来自数据库) ---")
    available_templates_from_db = tm.list_selectable_templates()
    if available_templates_from_db:
        for tpl_info in available_templates_from_db:
            print(f"  ID: {tpl_info.get('id')}, Name: {tpl_info.get('name')}, Name_repr: {tpl_info.get('name_repr')}, Filename: {tpl_info.get('json_filename')}")
    else:
        print("  数据库中没有找到模板记录。")
    print("--- 列表结束 ---\n")

    print(f"repr(mock_template_name) before load: {repr(mock_template_name)}")
    template_data_for_mapper = tm.load_template_json(template_name=mock_template_name)

    if not template_data_for_mapper:
        print(f"错误: 无法加载测试模板 '{mock_template_name}'。")
        print(f"请确保模板文件存在于: {user_files_dir / 'templates' / (mock_template_name + '.json')}")
        exit()
    else:
        print(f"成功加载测试模板 '{mock_template_name}'。")

    # 4. 实例化 LLMStyleMapper
    mapper = LLMStyleMapper(
        template_manager=tm,
        llm_client=None,
        template_data=template_data_for_mapper
    )
    print("LLMStyleMapper 实例化成功。")

    # 5. 调用 generate_mapping 并传入模拟回复
    print(f"\n调用 generate_mapping，使用模拟LLM回复...")
    final_mappings = mapper.generate_mapping(
        doc_df=mock_df,
        middle_start_index=0,
        back_start_index=None,
        template_name=mock_template_name,
        mock_llm_response_str_for_testing=mock_llm_output_str
    )

    # 6. 打印结果
    print("\n--- 模拟LLM调用后的最终映射结果 ---")
    if final_mappings:
        for m in final_mappings:
            print(f"段落索引: {m.get('paragraph_index')}, 映射样式: {m.get('style')}")
    else:
        print("未能生成任何映射。")

    # 可选: 保存到文件进行详细检查
    output_test_file = script_dir / f"test_llm_output_simulated_{mock_template_name}.json"
    try:
        mapper.save_mapping_to_file(final_mappings, str(output_test_file))
        print(f"\n测试映射结果已保存到: {output_test_file}")
    except Exception as e_save:
        print(f"\n保存测试映射结果失败: {e_save}")