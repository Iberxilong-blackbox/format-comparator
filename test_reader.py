print("--- test_reader.py script started ---", flush=True)
import os
import pythoncom # Important for explicit COM initialization/uninitialization in some contexts
import pandas as pd # Import pandas for DataFrame display
# If test_reader.py is in the win32com/ directory:
from docx_reader_win32 import DocxReaderWin32
# If test_reader.py is in the project root (one level above win32com/):
# from win32com.docx_reader_win32 import DocxReaderWin32 

def run_paragraph_extraction_test(file_path: str):
    """
    Tests the get_paragraph_data_df method of DocxReaderWin32.
    """
    print("\n--- Starting Paragraph Extraction Test ---", flush=True)
    pythoncom.CoInitialize()
    reader = DocxReaderWin32()
    print(f"Attempting to open document: {file_path}", flush=True)

    # Configure pandas display options
    pd.set_option('display.max_columns', None)
    pd.set_option('display.width', 200)
    pd.set_option('display.max_colwidth', 50) # Show more text content per cell

    # Define how many paragraphs to test for detailed print, but process all
    max_paragraphs_to_print = 5 

    if reader.open_document(file_path, visible=True): # Make Word visible for this test
        print("Document opened successfully (Word should be visible).", flush=True)
        try:
            print(f"\n--- Extracting All Paragraph Data ---", flush=True)
            # For target_basic.docx, "G1标题一" is a likely first chapter title
            df_full, middle_idx, back_idx = reader.get_paragraph_data_df(first_chapter_title="G1标题一") 
            
            print(f"\nFull DataFrame Info:", flush=True)
            df_full.info(verbose=True, show_counts=True) # More verbose info
            
            # Display a limited number of rows for brevity in console
            df_display = df_full.head(max_paragraphs_to_print)

            print(f"\nDataFrame Head (up to first {max_paragraphs_to_print} rows):", flush=True)
            print(df_display)

            if not df_display.empty:
                print(f"\nSample of extracted data (up to first {max_paragraphs_to_print} paragraphs, selected columns):", flush=True)
                cols_to_show = ['paragraph_index', 'text', 'style_name', 'alignment', 
                                'left_indent_pt', 'line_spacing_rule', 'line_spacing_value', 
                                'formatted_line_spacing']
                
                # Add list_info columns if they exist after potential normalization
                # This part of test_reader.py was getting complex; simplify for now
                # and focus on whether the main df_full has the data.
                # We can inspect df_full directly or print specific dicts.

                existing_cols_to_show = [col for col in cols_to_show if col in df_display.columns]
                print(df_display[existing_cols_to_show])

                if 'list_info' in df_display.columns:
                    print(f"\nList Info for first {min(max_paragraphs_to_print, len(df_display))} paragraphs:", flush=True)
                    for i in range(min(max_paragraphs_to_print, len(df_display))):
                        para_idx_val = df_display.loc[df_display.index[i], 'paragraph_index']
                        list_info_val = df_display.loc[df_display.index[i], 'list_info']
                        print(f"  Para {para_idx_val}: {list_info_val}", flush=True)

                if 'font_info' in df_display.columns:
                    print(f"\nFont Info for first {min(max_paragraphs_to_print, len(df_display))} paragraphs:", flush=True)
                    for i in range(min(max_paragraphs_to_print, len(df_display))):
                         para_idx_val = df_display.loc[df_display.index[i], 'paragraph_index']
                         font_info_val = df_display.loc[df_display.index[i], 'font_info']
                         print(f"  Para {para_idx_val}: {font_info_val}", flush=True)
            
            print(f"\nMiddle Start Index (for 'G1标题一'): {middle_idx}", flush=True)

        except Exception as e:
            print(f"An error occurred during paragraph data extraction: {e}", flush=True)
            import traceback
            traceback.print_exc()
        finally:
            print("\nClosing document...", flush=True)
            reader.close_document()
            print("Document closed.", flush=True)
    else:
        print(f"Failed to open document: {file_path}", flush=True)

    print("Quitting Word application instance...", flush=True)
    reader.quit_word()
    print("Word application instance quit.", flush=True)
    
    print(f"Attempting CoUninitialize...", flush=True)
    try:
        pythoncom.CoUninitialize()
        print(f"CoUninitialize successful.", flush=True)
    except Exception as e_uninit:
        print(f"Error during CoUninitialize: {e_uninit}", flush=True)
    print("--- Paragraph Extraction Test Finished ---", flush=True)


if __name__ == "__main__":
    project_root = os.path.abspath(os.path.dirname(__file__))
    # default_test_file = os.path.join(project_root, "target_basic.docx")
    
    # print(f"Test file path: {default_test_file}", flush=True)

    # test_docx_file_path = default_test_file

    # if not os.path.isabs(test_docx_file_path):
    #     print(f"错误：路径不是绝对路径: {test_docx_file_path}", flush=True)
    # elif not os.path.exists(test_docx_file_path):
    #     print(f"错误：文件未找到: {test_docx_file_path}", flush=True)
    # elif not test_docx_file_path.lower().endswith(".docx"):
    #     print(f"错误：文件不是 .docx 格式: {test_docx_file_path}", flush=True)
    # else:
    #     run_paragraph_extraction_test(test_docx_file_path)