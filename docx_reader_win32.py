import win32com.client
import pandas as pd
from typing import Dict, Any, List, Tuple, Optional
import pythoncom # Important for COM threading if used in threads
import re
from utils import normalize_text # Added for segmentation logic

# Assuming unit_converter.py will be in the same 'win32com' directory
# from .unit_converter import UnitConverter

class DocxReaderWin32:
    """
    Reads and extracts formatting information from Word documents using win32com.
    This class is intended to replace the functionality of the original DocxProcessor
    by using COM automation instead of python-docx and direct XML parsing.
    """

    def __init__(self, word_visible: bool = False, word_display_alerts: bool = False):
        """
        Initializes the DocxReaderWin32.

        Args:
            word_visible (bool): Whether the Word application window should be visible during processing.
            word_display_alerts (bool): Whether Word should display alerts during processing.
                                       Corresponds to Application.DisplayAlerts (WdAlertLevel).
                                       False typically means wdAlertsNone (0). True might mean wdAlertsAll (-1)
                                       or just not setting it to None. For simplicity, we'll map False to 0.
        """
        pythoncom.CoInitialize() # Initialize COM for this thread
        self.word_app: Optional[win32com.client.Dispatch] = None
        self.doc: Optional[win32com.client.Dispatch] = None
        self.constants: Optional[Any] = None # To store win32com.client.constants
        self._word_visible_on_init = word_visible
        self._word_display_alerts_on_init = 0 if not word_display_alerts else -1 # wdAlertsNone = 0, wdAlertsAll = -1
        # self.unit_converter = UnitConverter() # Initialize if needed for internal conversions

    def _initialize_constants(self):
        """Initializes win32com constants if not already done."""
        if self.constants is None and self.word_app is not None:
            try:
                self.constants = win32com.client.constants
                print("Successfully loaded win32com.client.constants", flush=True)
            except AttributeError:
                # This can happen if Word application is not fully initialized
                # or if constants are not available for some reason.
                # Fallback: Manually define critical constants or raise an error.
                print("Warning: win32com.client.constants could not be accessed. Using fallback. Enum mapping might be incomplete if real constants differ.", flush=True)
                # Define some common ones manually as a fallback if needed, or rely on direct integer comparison.
                # For robust solution, ensure Word app is properly up.
                # Example manual definitions (values might vary slightly with Word versions, check VBA Object Browser):
                self.constants = type('Constants', (), {
                    'wdAlignParagraphLeft': 0,
                    'wdAlignParagraphCenter': 1,
                    'wdAlignParagraphRight': 2,
                    'wdAlignParagraphJustify': 3,
                    'wdAlignParagraphDistribute': 4, # Less common
                    'wdAlignParagraphThaiJustify': 7, # Less common

                    'wdOutlineLevelBodyText': 10,
                    'wdOutlineLevel1': 1,
                    'wdOutlineLevel2': 2,
                    'wdOutlineLevel3': 3,
                    'wdOutlineLevel4': 4,
                    'wdOutlineLevel5': 5,
                    'wdOutlineLevel6': 6,
                    'wdOutlineLevel7': 7,
                    'wdOutlineLevel8': 8,
                    'wdOutlineLevel9': 9,

                    'wdLineSpaceSingle': 0,
                    'wdLineSpace1pt5': 1,
                    'wdLineSpaceDouble': 2,
                    'wdLineSpaceAtLeast': 3,
                    'wdLineSpaceExactly': 4,
                    'wdLineSpaceMultiple': 5,
                    
                    'wdListNoNumbering': 0,
                    'wdListSimpleNumbering': 2, # Example, check actual values
                    'wdListBullet': 1,          # Example

                    'wdUnderlineNone': 0,
                    'wdUnderlineSingle': 1,
                    # ... add other constants as needed
                })


    def open_document(self, file_path: str) -> bool:
        """
        Opens a Word document using win32com.
        Manages the Word application instance. Uses visibility and alert settings from __init__.

        Args:
            file_path (str): The absolute path to the .docx file.

        Returns:
            bool: True if the document was opened successfully, False otherwise.
        """
        # Ensure COM is initialized for the current thread if using multi-threading later
        # pythoncom.CoInitialize() 
        try:
            if self.word_app is None:
                # Try to get an existing instance first, if not, create a new one
                try:
                    self.word_app = win32com.client.GetActiveObject("Word.Application")
                except pythoncom.com_error:
                    self.word_app = win32com.client.Dispatch("Word.Application")
            
            self.word_app.Visible = self._word_visible_on_init
            if hasattr(self.word_app, 'DisplayAlerts'): # DisplayAlerts might not exist on all Word versions/objects
                self.word_app.DisplayAlerts = self._word_display_alerts_on_init # e.g., wdAlertsNone (0)
            
            self.doc = self.word_app.Documents.Open(FileName=file_path, ReadOnly=True, AddToRecentFiles=False)
            self._initialize_constants() # Initialize constants after document is open and app is confirmed
            return True
        except Exception as e:
            print(f"Error opening document '{file_path}' with win32com: {e}")
            if self.doc:
                try:
                    self.doc.Close(SaveChanges=0) # 0 for wdDoNotSaveChanges
                except: # noqa E722
                    pass
                self.doc = None
            # Only quit if we created the instance and it has no other docs,
            # or if GetActiveObject failed and we created a new one that then failed.
            # This logic can be complex; for now, if open fails, we don't aggressively quit a potentially shared instance.
            # if self.word_app and self.word_app.Documents.Count == 0:
            #     self.word_app.Quit()
            #     self.word_app = None
            return False

    def close_document(self) -> None:
        """
        Closes the currently open Word document without saving changes.
        """
        if self.doc:
            try:
                self.doc.Close(SaveChanges=0) # wdDoNotSaveChanges
            except Exception as e:
                print(f"Error closing document: {e}")
            finally:
                self.doc = None
        
        # Decide if word_app should be quit here.
        # Generally, the creator of the Word app instance should be responsible for quitting it.
        # If open_document always creates a new instance, then close_document (or a dedicated quit method) should quit it.
        # If open_document can attach to an existing instance, quitting is more complex.
        # For now, quit_word will handle quitting.

    def quit_word(self) -> None:
        """
        Quits the Word application instance if it's running and was potentially started by this class.
        This is a simplified quit, assuming this class instance "owns" the Word app instance.
        """
        if self.word_app:
            try:
                # Ensure the current document is closed if it's still referenced
                if self.doc:
                    self.doc.Close(SaveChanges=0)
                    self.doc = None
                
                # This will quit the Word application. 
                # If Word was already running and we attached via GetActiveObject, this will close that instance.
                # If we dispatched a new instance, it will close that new instance.
                self.word_app.Quit()
            except Exception as e:
                print(f"Error quitting Word application: {e}")
            finally:
                self.word_app = None
        pythoncom.CoUninitialize() # If CoInitialize was called in __init__

    def get_paragraph_data_df(self,
                              first_chapter_title: str = "绪论",
                              back_markers: List[str] = ["参考文献", "致谢", "附录"],
                              add_segment_type_column: bool = True
                              ) -> Tuple[pd.DataFrame, Optional[int], Optional[int]]:
        if not self.doc:
            raise RuntimeError("No document is open. Call open_document() first.")

        all_para_data: List[Dict[str, Any]] = []
        middle_start_index: Optional[int] = None
        back_start_index: Optional[int] = None

        normalized_first_title = normalize_text(first_chapter_title)
        normalized_back_markers = [normalize_text(marker) for marker in back_markers]

        core_pattern_part = re.escape(normalized_first_title).replace(r'\ ', r'\s*')
        first_chapter_pattern_str = (
            r"^\s*(?:(?:第\s*[一二三四五六七八九十百千万亿零〇０-９\d]+\s*[章节篇部])|(?:[一二三四五六七八九十百千万亿零〇０-９\d]+\s*[\.\s、．]))?\s*" +
            core_pattern_part +
            r"\s*$"
        )
        first_chapter_pattern = re.compile(first_chapter_pattern_str, re.IGNORECASE)

        print("Starting paragraph data extraction...")
        for i, para_obj in enumerate(self.doc.Paragraphs):
            paragraph_index_0_based = i
            try:
                para_info = self._get_paragraph_info(para_obj, paragraph_index_0_based)
                all_para_data.append(para_info)
                
                current_para_text_normalized = normalize_text(para_info.get('text', ''))

                # Find middle_start_index
                if middle_start_index is None:
                    if first_chapter_pattern.match(current_para_text_normalized):
                        middle_start_index = paragraph_index_0_based
                        print(f"Found first chapter title '{first_chapter_title}' (regex match) at paragraph index {middle_start_index}", flush=True)
                
                # Find back_start_index
                if middle_start_index is not None and back_start_index is None and paragraph_index_0_based >= middle_start_index:
                    for norm_marker in normalized_back_markers:
                        marker_pattern_str = r"^\s*" + re.escape(norm_marker) + r"\s*$"
                        marker_pattern = re.compile(marker_pattern_str, re.IGNORECASE)
                        if marker_pattern.match(current_para_text_normalized):
                            back_start_index = paragraph_index_0_based
                            print(f"Found back marker '{norm_marker}' at paragraph index {back_start_index}", flush=True)
                            break
            
            except Exception as e:
                print(f"Error processing paragraph at COM index {i+1} (0-based {paragraph_index_0_based}): {e}", flush=True)
                all_para_data.append({
                    'paragraph_index': paragraph_index_0_based,
                    'text': f"[Error: {e}]",
                    'style_name': "[Error]",
                    'outline_level': None,
                    'alignment': None,
                    'left_indent_pt': None,
                    'right_indent_pt': None,
                    'first_line_indent_pt': None,
                    'space_before_pt': None,
                    'space_after_pt': None,
                    'line_spacing_rule': None,
                    'line_spacing_value': None,
                    'formatted_line_spacing': "[Error]",
                    'font_info': [],
                    'list_info': {}
                })
        
        print(f"Processed {len(all_para_data)} paragraphs.", flush=True)

        if middle_start_index is None:
            print(f"Warning: First chapter title '{first_chapter_title}' not found. Assuming body starts at paragraph 0.", flush=True)
            middle_start_index = 0
        
        # back_start_index remains None if no markers found, indicating body extends to the end.

        df = pd.DataFrame(all_para_data)
        
        if add_segment_type_column and not df.empty:
            segment_types = []
            for idx_val in df['paragraph_index']:
                if idx_val < middle_start_index:
                    segment_types.append('front_matter')
                elif back_start_index is None or idx_val < back_start_index:
                    segment_types.append('body_matter')
                else: # idx_val >= back_start_index
                    segment_types.append('back_matter')
            df['segment_type'] = segment_types
        
        # Ensure all expected columns exist in the DataFrame
        expected_cols = [
            'paragraph_index', 'text', 'style_name', 'outline_level',
            'alignment', 'left_indent_pt', 'right_indent_pt', 'first_line_indent_pt',
            'space_before_pt', 'space_after_pt', 'line_spacing_rule', 'line_spacing_value',
            'formatted_line_spacing', 'font_info', 'list_info',
            'paragraph_actual_font_size_pt' # Added for char unit conversion context
            # Add more as they are implemented in _get_paragraph_info
        ]
        if add_segment_type_column:
            expected_cols.append('segment_type')

        for col in expected_cols:
            if col not in df.columns:
                df[col] = None # Add missing columns with None
                
        return df, middle_start_index, back_start_index

    def get_document_default_fonts(self) -> Dict[str, Any]:
        """
        Extracts default font information for the document (e.g., for Normal style).

        Returns:
            Dict[str, Any]: {'ascii': str, 'eastAsia': str, 'size_pt': float}
        """
        if not self.doc:
            raise RuntimeError("No document is open. Call open_document() first.")
        
        print("Warning: get_document_default_fonts is a skeleton and not fully implemented.")
        # Example:
        # try:
        #     normal_style = self.doc.Styles("Normal")
        #     default_font = normal_style.Font
        #     return {
        #         'ascii': default_font.NameAscii,
        #         'eastAsia': default_font.NameFarEast,
        #         'size_pt': default_font.Size
        #     }
        # except Exception as e:
        #     print(f"Could not get Normal style font: {e}")
        #     return {}
        return {}
        # raise NotImplementedError("get_document_default_fonts is not yet implemented.")

    def get_page_setup_info(self) -> Dict[str, Any]:
        """
        Extracts page setup information (margins, paper size, orientation).

        Returns:
            Dict[str, Any]: Structured page setup information.
        """
        if not self.doc:
            raise RuntimeError("No document is open. Call open_document() first.")
        
        print("Warning: get_page_setup_info is a skeleton and not fully implemented.")
        # Example:
        # page_setup = self.doc.PageSetup
        # return {
        #     'page_size': {
        #         'width_pt': page_setup.PageWidth,
        #         'height_pt': page_setup.PageHeight,
        #         'orientation': page_setup.Orientation 
        #     },
        #     'page_margins': {
        #         'top_pt': page_setup.TopMargin, 
        #         'bottom_pt': page_setup.BottomMargin,
        #         'left_pt': page_setup.LeftMargin,
        #         'right_pt': page_setup.RightMargin,
        #         'header_pt': page_setup.HeaderDistance,
        #         'footer_pt': page_setup.FooterDistance,
        #         'gutter_pt': page_setup.Gutter
        #     }
        # }
        return {}
        # raise NotImplementedError("get_page_setup_info is not yet implemented.")

    def get_document_metadata(self) -> Dict[str, Any]:
        """
        Extracts document metadata (author, title, creation_date, etc.).

        Returns:
            Dict[str, Any]: A dictionary of document properties.
        """
        if not self.doc:
            raise RuntimeError("No document is open. Call open_document() first.")
        
        metadata: Dict[str, Any] = {}
        try:
            # Accessing BuiltInDocumentProperties collection
            # Each item in this collection is a DocumentProperty object
            for prop in self.doc.BuiltInDocumentProperties:
                prop_name = ""
                prop_value: Any = None
                try:
                    prop_name = str(prop.Name) # Ensure name is a string
                    prop_value = prop.Value
                    # Handle cases where Value might be an object with specific properties (e.g., dates)
                    # For now, we take it as is. Further type checking/conversion can be added.
                    # Example: if isinstance(prop_value, pywintypes.TimeType):
                    # prop_value = str(prop_value) # Convert to string or datetime object
                except pythoncom.com_error as ce:
                    # Some properties might exist but have no value or throw an error on access
                    print(f"  COM Error accessing property '{prop_name or 'Unknown'}': {ce}")
                    prop_value = None # Or some other placeholder like "[Access Error]"
                except Exception as e:
                    print(f"  Unexpected error accessing property '{prop_name or 'Unknown'}': {e}")
                    prop_value = None
                
                if prop_name: # Only add if name was successfully retrieved
                    metadata[prop_name] = prop_value
            
            # You can also access custom document properties if needed:
            # for prop in self.doc.CustomDocumentProperties:
            #     try:
            #         metadata[f"Custom_{prop.Name}"] = prop.Value
            #     except:
            #         metadata[f"Custom_{prop.Name}"] = None

        except pythoncom.com_error as ce:
            print(f"COM Error accessing BuiltInDocumentProperties collection: {ce}")
        except Exception as e:
            print(f"Unexpected error reading document metadata: {e}")
            # Depending on desired robustness, you might want to return partial metadata or raise
        
        return metadata

    def get_formula_info(self) -> List[Dict[str, Any]]:
        """
        Extracts information about OMML formulas in the document.

        Returns:
            List[Dict[str, Any]]: A list of dictionaries, each representing a formula.
        """
        if not self.doc:
            raise RuntimeError("No document is open. Call open_document() first.")
        
        print("Warning: get_formula_info is a skeleton and not fully implemented.")
        # formulas = []
        # try:
        #     for i in range(1, self.doc.OMaths.Count + 1):
        #         omath = self.doc.OMaths(i)
        #         # Determine paragraph index - this is complex and needs a robust way
        #         # para_index = ... 
        #         formulas.append({
        #             'type': 'paragraph' if omath.Type == 0 else 'inline', # wdOMathDisplay = 0, wdOMathInline = 1
        #             'paragraph_index': None, # Placeholder for actual index
        #             'omml_content': omath.Range.OMML
        #         })
        # except Exception as e:
        #     print(f"Error extracting OMaths: {e}")
        # return formulas
        return []
        # raise NotImplementedError("get_formula_info is not yet implemented.")

    def extract_full_text(self) -> str:
        """
        Extracts the full plain text content of the document.
        To be used by a new text_extractor_win32.py if text extraction from docx is needed.

        Returns:
            str: The plain text content of the document.
        """
        if not self.doc:
            raise RuntimeError("No document is open. Call open_document() first.")
        try:
            return self.doc.Content.Text
        except Exception as e:
            print(f"Error extracting full text: {e}")
            return ""

    # --- Private Helper Methods (to be implemented) ---
    def _get_paragraph_info(self, para_obj: Any, paragraph_index: int) -> Dict[str, Any]:
        """
        Helper to extract all relevant information for a single paragraph.
        Starts with text and style name.
        """
        print(f"  Processing para_info for index: {paragraph_index}", flush=True)
        para_data: Dict[str, Any] = {'paragraph_index': paragraph_index}
        
        try:
            raw_text = para_obj.Range.Text
            if raw_text is not None:
                cleaned_text = raw_text.rstrip('\r\n\x07\x0b\x0c')
                cleaned_text = cleaned_text.replace('\x07', '').replace('\x0b', '').replace('\x0c', '')
                para_data['text'] = cleaned_text
            else:
                para_data['text'] = "" 
        except Exception as e:
            print(f"  Error getting text for paragraph {paragraph_index}: {e}", flush=True)
            para_data['text'] = "[Error]"

        print(f"    _get_paragraph_info: Text extracted for {paragraph_index}", flush=True)
        try:
            style_obj = para_obj.Style
            para_data['style_name'] = style_obj.NameLocal 
        except Exception as e:
            print(f"  Error getting style name for paragraph {paragraph_index}: {e}", flush=True)
            para_data['style_name'] = "[Error]"
        
        print(f"    _get_paragraph_info: Style extracted for {paragraph_index}", flush=True)
        try:
            para_format = para_obj.Format 
            
            # Helper function to safely get constant values by string name from self.constants
            def get_const(name: str, default_value: int) -> int:
                if self.constants:
                    try:
                        # This should work if self.constants is the real COM object 
                        # or our fallback type object where keys are attributes.
                        return getattr(self.constants, name)
                    except AttributeError:
                        # This handles if the specific constant name is missing or if self.constants is the fallback
                        # and the name was intended as a dict key (which it isn't for the type() fallback).
                        # For the type() fallback, getattr should have worked if the key was in the dict.
                        # This also covers if self.constants is the real COM object but a very specific/rare constant isn't exposed.
                        # print(f"Debug: Constant '{name}' not found via direct getattr, trying dict access on fallback or using default.", flush=True)
                        if isinstance(self.constants, type) and name in self.constants.__dict__: # Check if it's our fallback type and has the key
                             return self.constants.__dict__[name] # Access as dict item for fallback
                        # print(f"Warning: Constant '{name}' not found, using default {default_value}", flush=True)
                        return default_value
                return default_value

            # Outline Level
            raw_outline_level = para_obj.OutlineLevel
            const_wdOutlineLevelBodyText = get_const('wdOutlineLevelBodyText', 10)
            const_wdOutlineLevel1 = get_const('wdOutlineLevel1', 1)
            const_wdOutlineLevel9 = get_const('wdOutlineLevel9', 9)

            if raw_outline_level == const_wdOutlineLevelBodyText:
                para_data['outline_level'] = None 
            elif const_wdOutlineLevel1 <= raw_outline_level <= const_wdOutlineLevel9:
                para_data['outline_level'] = raw_outline_level
            else: 
                para_data['outline_level'] = raw_outline_level 
            
            # Alignment
            raw_alignment = para_format.Alignment
            alignment_str = str(raw_alignment) 
            if self.constants: 
                alignment_map = {
                    get_const('wdAlignParagraphLeft', 0): "left",
                    get_const('wdAlignParagraphCenter', 1): "center",
                    get_const('wdAlignParagraphRight', 2): "right",
                    get_const('wdAlignParagraphJustify', 3): "justify",
                    get_const('wdAlignParagraphDistribute', 4): "distribute", 
                    get_const('wdAlignParagraphThaiJustify', 7): "thai_justify"
                }
                alignment_str = alignment_map.get(raw_alignment, str(raw_alignment)) 
            para_data['alignment'] = alignment_str
            
            para_data['left_indent_pt'] = para_format.LeftIndent
            para_data['right_indent_pt'] = para_format.RightIndent
            para_data['first_line_indent_pt'] = para_format.FirstLineIndent 
            para_data['space_before_pt'] = para_format.SpaceBefore
            para_data['space_after_pt'] = para_format.SpaceAfter
            
            raw_line_spacing_rule = para_format.LineSpacingRule
            raw_line_spacing_value = para_format.LineSpacing
            line_spacing_rule_str = str(raw_line_spacing_rule) 
            line_spacing_value_for_df = raw_line_spacing_value

            if self.constants:
                ls_map = {
                    get_const('wdLineSpaceSingle', 0): "single",
                    get_const('wdLineSpace1pt5', 1): "one_and_half",
                    get_const('wdLineSpaceDouble', 2): "double",
                    get_const('wdLineSpaceAtLeast', 3): "at_least",
                    get_const('wdLineSpaceExactly', 4): "exactly",
                    get_const('wdLineSpaceMultiple', 5): "multiple"
                }
                line_spacing_rule_str = ls_map.get(raw_line_spacing_rule, str(raw_line_spacing_rule))
                
                print(f"[DEBUG] Para {paragraph_index}: LineSpacingRule={raw_line_spacing_rule}, RawValue={raw_line_spacing_value}", flush=True)
                
                const_wdLineSpaceSingle = get_const('wdLineSpaceSingle', 0)
                const_wdLineSpace1pt5 = get_const('wdLineSpace1pt5', 1)
                const_wdLineSpaceDouble = get_const('wdLineSpaceDouble', 2)
                const_wdLineSpaceMultiple = get_const('wdLineSpaceMultiple', 5)
                
                if raw_line_spacing_rule == const_wdLineSpaceSingle:
                    line_spacing_value_for_df = 1.0
                elif raw_line_spacing_rule == const_wdLineSpace1pt5:
                    line_spacing_value_for_df = 1.5
                elif raw_line_spacing_rule == const_wdLineSpaceDouble:
                    line_spacing_value_for_df = 2.0
                elif raw_line_spacing_rule == const_wdLineSpaceMultiple:
                    if raw_line_spacing_value is not None:
                        try:
                            line_spacing_value_for_df = round(float(raw_line_spacing_value), 2)
                        except (ValueError, TypeError):
                            print(f"  Warning: Could not convert line_spacing_value '{raw_line_spacing_value}' to float.", flush=True)
                else:
                    line_spacing_value_for_df = raw_line_spacing_value
            
            para_data['line_spacing_rule'] = line_spacing_rule_str
            para_data['line_spacing_value'] = line_spacing_value_for_df

            formatted_ls = f"Rule: {raw_line_spacing_rule}, Value: {raw_line_spacing_value}" 
            if line_spacing_rule_str == "single":
                formatted_ls = "单倍"
            elif line_spacing_rule_str == "one_and_half":
                formatted_ls = "1.5倍"
            elif line_spacing_rule_str == "double":
                formatted_ls = "2倍"
            elif line_spacing_rule_str == "at_least" and line_spacing_value_for_df is not None:
                try:
                    formatted_ls = f"最小值: {float(line_spacing_value_for_df):.2f} pt"
                except (ValueError, TypeError):
                    formatted_ls = f"最小值: {line_spacing_value_for_df} (raw)"
            elif line_spacing_rule_str == "exactly" and line_spacing_value_for_df is not None:
                try:
                    formatted_ls = f"固定值: {float(line_spacing_value_for_df):.2f} pt"
                except (ValueError, TypeError):
                    formatted_ls = f"固定值: {line_spacing_value_for_df} (raw)"
            elif line_spacing_rule_str == "multiple" and line_spacing_value_for_df is not None:
                formatted_ls = f"多倍: {line_spacing_value_for_df}"
            
            para_data['formatted_line_spacing'] = formatted_ls

        except Exception as e:
            print(f"  Error getting paragraph format properties for paragraph {paragraph_index}: {e}", flush=True)
            para_data['outline_level'] = None
            para_data['alignment'] = None
            para_data['left_indent_pt'] = None
            para_data['right_indent_pt'] = None
            para_data['first_line_indent_pt'] = None
            para_data['space_before_pt'] = None
            para_data['space_after_pt'] = None
            para_data['line_spacing_rule'] = None
            para_data['line_spacing_value'] = None
            para_data['formatted_line_spacing'] = "[Error]"

        print(f"    _get_paragraph_info: Para format extracted for {paragraph_index}", flush=True)
        try:
            list_format = para_obj.Range.ListFormat
            para_data['list_info'] = {
                'is_list': list_format.ListType != 0, 
                'list_type': list_format.ListType, 
                'list_level': list_format.ListLevelNumber,
                'list_string': list_format.ListString, 
                'list_value': list_format.ListValue 
            }
        except Exception as e:
            print(f"  Error getting list format info for paragraph {paragraph_index}: {e}", flush=True)
            para_data['list_info'] = {
                'is_list': None, 'list_type': None, 'list_level': None,
                'list_string': "[Error]", 'list_value': None
            }
        print(f"    _get_paragraph_info: List info extracted for {paragraph_index}", flush=True)

        try:
            para_data['font_info'] = self._get_runs_info(para_obj.Range)
        except Exception as e:
            print(f"  Error getting font_info for paragraph {paragraph_index}: {e}", flush=True)
            para_data['font_info'] = [{'text': '[Error]', 'error_message': str(e)}]
        
        # Attempt to get paragraph's main font size from the first valid run
        para_data['paragraph_actual_font_size_pt'] = None
        if isinstance(para_data.get('font_info'), list) and para_data.get('font_info'):
            for run_detail in para_data['font_info']:
                if isinstance(run_detail.get('size'), (int, float)):
                    para_data['paragraph_actual_font_size_pt'] = run_detail['size']
                    # print(f"    [DEBUG _get_paragraph_info] Para {paragraph_index} actual_font_size_pt set to: {run_detail['size']}", flush=True)
                    break
            if para_data['paragraph_actual_font_size_pt'] is None:
                pass
                # print(f"    [DEBUG _get_paragraph_info] Para {paragraph_index} could not determine actual_font_size_pt from runs.", flush=True)
        else:
            pass
            # print(f"    [DEBUG _get_paragraph_info] Para {paragraph_index} has no font_info or it's not a list.", flush=True)


        print(f"  Finished para_info for index: {paragraph_index}", flush=True)
        return para_data

    def _get_runs_info(self, para_range_obj: Any) -> List[Dict[str, Any]]:
        """
        Helper to iterate through effective "runs" in a paragraph's Range
        and extract font information for each distinct formatting segment.
        This is a simplified first pass, iterating by Words.
        """
        runs_data: List[Dict[str, Any]] = []
        if not para_range_obj:
            return runs_data
        try:
            for i in range(1, para_range_obj.Words.Count + 1):
                word_range = para_range_obj.Words(i)
                raw_word_text = word_range.Text
                
                # Clean the word text
                cleaned_word_text = raw_word_text
                if cleaned_word_text:
                    # Similar to paragraph text cleaning, but \r\n might be too aggressive for single words.
                    # Focus on problematic characters observed.
                    cleaned_word_text = cleaned_word_text.replace('\r', '') # Carriage return
                    cleaned_word_text = cleaned_word_text.replace('\x07', '') # Bell char
                    cleaned_word_text = cleaned_word_text.replace('\x0b', '') # Vertical tab
                    cleaned_word_text = cleaned_word_text.replace('\x0c', '') # Form feed
                    cleaned_word_text = cleaned_word_text.replace('\t', ' ') # Replace tab with space, or remove if preferred
                    # Trailing space from Word's .Words collection is common, strip it.
                    # Also strip leading/trailing whitespace that might have been introduced or was already there.
                    cleaned_word_text = cleaned_word_text.strip()

                if cleaned_word_text: # Proceed only if there's content after cleaning
                    font_obj = word_range.Font
                    font_details = self._get_font_info(font_obj)
                    font_details['text'] = cleaned_word_text
                    runs_data.append(font_details)
        except Exception as e:
            print(f"    Error iterating words/extracting runs_info: {e}")
            runs_data.append({'text': '[Error in _get_runs_info]', 'error_message': str(e)})
            
        # Fallback for paragraphs that might not be properly processed by Words collection (e.g., empty paragraphs with formatting)
        # or if Words collection is empty but paragraph has text.
        if not runs_data and para_range_obj.Text:
            raw_para_text = para_range_obj.Text
            # Clean the fallback paragraph text
            cleaned_para_text = raw_para_text.replace('\r', '').replace('\x07', '').replace('\x0b', '').replace('\x0c', '').replace('\t', ' ').strip()
            if cleaned_para_text:
                try:
                    font_obj = para_range_obj.Font
                    font_details = self._get_font_info(font_obj)
                    font_details['text'] = cleaned_para_text
                    runs_data.append(font_details)
                except Exception as e_fallback:
                    print(f"    Error in fallback font extraction for paragraph: {e_fallback}")
                    runs_data.append({'text': cleaned_para_text, 'error_message': 'Fallback font extraction failed'})
        return runs_data

    def _convert_color_to_hex(self, color_val: Optional[int]) -> Optional[str]:
        """
        Converts a Word color value (RGB integer) to a hex string #RRGGBB.
        Handles special Word color constants like wdColorAutomatic or wdUndefined.
        """
        if color_val is None or color_val == -16777216:  # wdColorAutomatic (often black or default)
            return "#000000" # Defaulting automatic to black, or could be None
        if color_val == 9999999: # wdUndefined, wdToggle, often means 'auto' or 'inherit'
            return None # Represents automatic or inherited color

        # Ensure color_val is a valid integer for bitwise operations
        if not isinstance(color_val, int):
            return None

        # Word's RGB is an integer: (Blue * 256^2) + (Green * 256) + Red
        # So we need to extract R, G, B
        blue = (color_val >> 16) & 0xFF
        green = (color_val >> 8) & 0xFF
        red = color_val & 0xFF
        return f"#{red:02x}{green:02x}{blue:02x}"

    WD_UNDERLINE_MAP = {
        0: "none",             # wdUnderlineNone
        1: "single",           # wdUnderlineSingle
        2: "words",            # wdUnderlineWords
        3: "double",           # wdUnderlineDouble
        4: "dotted",           # wdUnderlineDotted
        5: "thick",            # wdUnderlineThick
        6: "dash",             # wdUnderlineDash
        7: "dot_dash",         # wdUnderlineDotDash
        8: "dot_dot_dash",     # wdUnderlineDotDotDash
        9: "wave",             # wdUnderlineWave
        10: "dotted_heavy",    # wdUnderlineDottedHeavy (WdUnderline enumeration in VBA)
        11: "dashed_heavy",    # wdUnderlineDashHeavy
        12: "dash_dot_heavy",  # wdUnderlineDashDotHeavy
        13: "dash_dot_dot_heavy",# wdUnderlineDashDotDotHeavy
        14: "wave_heavy",      # wdUnderlineWaveHeavy
        20: "dash_long",       # wdUnderlineDashLong
        23: "wave_double",     # wdUnderlineWaveDouble
        27: "dash_long_heavy", # wdUnderlineDashLongHeavy
        9999999: "automatic"   # wdUndefined / wdAutomatic
    }

    def _get_font_info(self, font_obj: Any) -> Dict[str, Any]:
        """ Helper to extract detailed info from a win32com Font object """
        details: Dict[str, Any] = {}
        WD_UNDEFINED = 9999999  # Standard value for wdUndefined in Word COM

        try:
            # --- Font Names and Size Extraction ---
            raw_name = font_obj.Name
            raw_name_ascii = font_obj.NameAscii
            raw_name_fareast = font_obj.NameFarEast
            raw_size_pt = font_obj.Size

            # Convert wdUndefined to None
            processed_name = None if raw_name == WD_UNDEFINED else raw_name
            processed_name_ascii = None if raw_name_ascii == WD_UNDEFINED else raw_name_ascii
            processed_name_fareast = None if raw_name_fareast == WD_UNDEFINED else raw_name_fareast
            processed_size_pt = None if raw_size_pt == WD_UNDEFINED else raw_size_pt

            # Heuristic: If specific font names (ASCII/FarEast) are None or theme font strings (e.g., "+Body"),
            # and the general 'Name' property is a resolved font name, use 'Name' as a fallback.
            final_name_fareast = processed_name_fareast
            if (processed_name_fareast is None or (isinstance(processed_name_fareast, str) and processed_name_fareast.startswith('+'))) and \
               (processed_name is not None and not (isinstance(processed_name, str) and processed_name.startswith('+'))):
                final_name_fareast = processed_name
            
            final_name_ascii = processed_name_ascii
            if (processed_name_ascii is None or (isinstance(processed_name_ascii, str) and processed_name_ascii.startswith('+'))) and \
               (processed_name is not None and not (isinstance(processed_name, str) and processed_name.startswith('+'))):
                final_name_ascii = processed_name

            # --- Color Extraction ---
            color_hex_str: Optional[str] = None
            original_color_rgb_val: Optional[int] = None
            original_font_color_bgr_val: Optional[int] = font_obj.Color

            if hasattr(font_obj, 'TextColor') and font_obj.TextColor.Type == 1: # wdColorRGB
                original_color_rgb_val = font_obj.TextColor.RGB
                color_hex_str = self._convert_color_to_hex(original_color_rgb_val)
            else:
                font_color_from_prop = font_obj.Color
                if font_color_from_prop == -16777216:  # wdColorAutomatic
                    color_hex_str = "#000000"
                elif font_color_from_prop == WD_UNDEFINED:
                    color_hex_str = None
                else:
                    color_hex_str = None

            # --- Underline Extraction ---
            underline_type_val = font_obj.Underline
            underline_str: Optional[str] = None
            if underline_type_val == WD_UNDEFINED:
                underline_str = None  # So comparator's default ('none') can be used
            elif underline_type_val == 0: # wdUnderlineNone
                underline_str = "none"
            else:
                underline_str = self.WD_UNDERLINE_MAP.get(underline_type_val, str(underline_type_val))

            # --- Bold/Italic Extraction ---
            raw_bold = font_obj.Bold
            processed_bold = bool(raw_bold) if raw_bold != WD_UNDEFINED else None
            
            raw_italic = font_obj.Italic
            processed_italic = bool(raw_italic) if raw_italic != WD_UNDEFINED else None

            # Keys here MUST match what FormatComparatorWin32 expects via TEMPLATE_TO_RUNINFO_KEY_MAP
            details = {
                'name': processed_name,
                'font_ascii': final_name_ascii,    # Corrected key
                'font_eastasia': final_name_fareast, # Corrected key
                'size': processed_size_pt,         # Corrected key (was size_pt)
                'bold': processed_bold,
                'italic': processed_italic,
                'underline_type': underline_str, # Corrected key (was underline_type_str, and 'underline' was separate)
                'color_hex': color_hex_str,
                # Storing original values for debugging if needed by other tools
                '_original_underline_type_val': underline_type_val,
                '_original_color_rgb_val': original_color_rgb_val,
                '_original_font_color_bgr_val': original_font_color_bgr_val,
                '_raw_name': raw_name, # For deeper debugging if necessary
                '_raw_name_ascii': raw_name_ascii,
                '_raw_name_fareast': raw_name_fareast,
                '_raw_size_pt': raw_size_pt,
                '_raw_bold': raw_bold,
                '_raw_italic': raw_italic
            }
            
            # DEBUG LOGS - Placed after 'details' is populated to show what's being returned.
            # The user moved these logs here, which is fine.
            print(f"    [DEBUG _get_font_info] Populated 'details' dict content:")
            print(f"      name: '{details.get('name')}' (Type: {type(details.get('name'))})")
            print(f"      font_ascii: '{details.get('font_ascii')}' (Type: {type(details.get('font_ascii'))})")
            print(f"      font_eastasia: '{details.get('font_eastasia')}' (Type: {type(details.get('font_eastasia'))})")
            print(f"      size: '{details.get('size')}' (Type: {type(details.get('size'))})")
            print(f"      bold: '{details.get('bold')}' (Type: {type(details.get('bold'))})")
            print(f"      italic: '{details.get('italic')}' (Type: {type(details.get('italic'))})")
            print(f"      underline_type: '{details.get('underline_type')}' (Type: {type(details.get('underline_type'))})")
            print(f"      color_hex: '{details.get('color_hex')}' (Type: {type(details.get('color_hex'))})")
            
        except Exception as e:
            print(f"      Error extracting font details: {e}")
            # Ensure basic keys exist even on error, to prevent DataFrame issues
            # Using keys expected by FormatComparatorWin32
            details.update({
                'name': "[Error]", 'font_ascii': "[Error]", 'font_eastasia': "[Error]",
                'size': None, 'bold': None, 'italic': None,
                'underline_type': "[Error]", 'color_hex': None,
                'error_message': str(e)
            })
        return details

    def _get_paragraph_format_info(self, para_format_obj: Any) -> Dict[str, Any]:
        """ Helper to extract detailed info from a win32com ParagraphFormat object """
        raise NotImplementedError("_get_paragraph_format_info is not yet implemented.")

    def _get_list_format_info(self, list_format_obj: Any) -> Dict[str, Any]:
        """ Helper to extract list/numbering info from a ListFormat object """
        raise NotImplementedError("_get_list_format_info is not yet implemented.")

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close_document()