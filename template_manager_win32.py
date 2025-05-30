# win32com/template_manager_win32.py
import sqlite3
import json
import os
from pathlib import Path
import datetime
import re
from typing import Optional, Tuple, List, Dict, Any

class TemplateManagerWin32:
    """
    Manages style templates for the win32com version of the application.
    Templates are stored as JSON files in a user directory, and metadata is kept in an SQLite database.
    """
    def __init__(self, base_user_dir: Path = Path("user_files")):
        """
        Initializes the TemplateManagerWin32.

        Args:
            base_user_dir (Path): The base directory for user-specific files (e.g., 'win32com/user_files').
                                  This path should be relative to where the app is run or an absolute path.
        """
        self.base_user_dir = base_user_dir
        self.templates_dir = self.base_user_dir / "templates"
        self.db_path = self.base_user_dir / "templates_map.db"
        
        self._init_db()

    def _sanitize_filename(self, name: str) -> str:
        """Sanitizes a string to be used as a filename."""
        name = re.sub(r'[^\w\s-]', '', name).strip()
        name = re.sub(r'[-\s]+', '-', name)
        return name if name else "unnamed_template"

    def _init_db(self):
        """Initializes the database and templates directory if they don't exist."""
        try:
            self.templates_dir.mkdir(parents=True, exist_ok=True)
            
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS templates (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    name TEXT NOT NULL UNIQUE,
                    json_filename TEXT NOT NULL UNIQUE,
                    created_at TEXT DEFAULT CURRENT_TIMESTAMP
                )
            """)
            conn.commit()
        except sqlite3.Error as e:
            print(f"Database initialization error: {e}")
            raise # Re-raise after logging, as this is critical
        finally:
            if conn:
                conn.close()

    def save_template(self, name: str, style_rules_dict: dict) -> Tuple[bool, str]:
        """
        Saves a new template.

        Args:
            name (str): The name of the template.
            style_rules_dict (dict): A dictionary containing the style rules ("样式" part).

        Returns:
            Tuple[bool, str]: (success, message or template_id if successful)
        """
        if not name:
            return False, "Template name cannot be empty."
        if not isinstance(style_rules_dict, dict):
            return False, "Style rules must be a dictionary."

        sanitized_name = self._sanitize_filename(name)
        json_filename = f"{sanitized_name}_{datetime.datetime.now().strftime('%Y%m%d%H%M%S%f')}.json"
        json_file_path = self.templates_dir / json_filename

        full_template_content = {
            "name": name,
            "样式": style_rules_dict,
            "_source_filename": json_filename # For self-reference if needed
        }

        conn = None
        try:
            # Save JSON file
            with open(json_file_path, 'w', encoding='utf-8') as f:
                json.dump(full_template_content, f, ensure_ascii=False, indent=2)

            # Save metadata to DB
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute("""
                INSERT INTO templates (name, json_filename)
                VALUES (?, ?)
            """, (name, json_filename))
            conn.commit()
            new_id = cursor.lastrowid
            return True, f"Template '{name}' saved successfully with ID {new_id}."

        except sqlite3.IntegrityError: # Handles UNIQUE constraint violation for name or json_filename
            if json_file_path.exists(): # Clean up orphaned file if DB insert failed
                try: os.remove(json_file_path)
                except OSError: pass
            return False, f"Failed to save template. A template with name '{name}' or filename '{json_filename}' might already exist."
        except IOError as e:
            return False, f"Error saving template JSON file '{json_filename}': {e}"
        except Exception as e: # Catch-all for other unexpected errors
            if json_file_path.exists(): # Clean up
                try: os.remove(json_file_path)
                except OSError: pass
            return False, f"An unexpected error occurred while saving template: {e}"
        finally:
            if conn:
                conn.close()

    def list_selectable_templates(self) -> List[Dict[str, Any]]:
        """
        Lists all available templates for selection.

        Returns:
            List[Dict[str, Any]]: A list of dictionaries, each containing
                                  'id', 'name'.
        """
        conn = None
        try:
            conn = sqlite3.connect(self.db_path)
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()
            # Also select json_filename for debugging purposes
            cursor.execute("SELECT id, name, json_filename FROM templates ORDER BY name COLLATE NOCASE ASC")
            templates = []
            for row in cursor.fetchall():
                tpl_dict = dict(row)
                # Add repr(name) for exact string representation
                tpl_dict['name_repr'] = repr(tpl_dict['name'])
                templates.append(tpl_dict)
            return templates
        except sqlite3.Error as e:
            print(f"Error listing templates: {e}")
            return []
        finally:
            if conn:
                conn.close()

    def load_template_json(self, template_id: Optional[int] = None, template_name: Optional[str] = None) -> Optional[Dict[str, Any]]:
        """
        Loads the JSON content of a specific template by its ID or name.

        Args:
            template_id (Optional[int]): The ID of the template to load.
            template_name (Optional[str]): The name of the template to load.
                                         If both id and name are provided, id takes precedence.
        Returns:
            Optional[Dict[str, Any]]: The parsed JSON content of the template, or None if not found or error.
        """
        if template_id is None and template_name is None:
            print("Error: template_id or template_name must be provided to load_template_json.")
            return None

        conn = None
        query = "SELECT json_filename FROM templates WHERE "
        params = []

        if template_id is not None:
            query += "id = ?"
            params.append(template_id)
        elif template_name is not None:
            query += "name = ?"
            params.append(template_name)
        
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute(query, tuple(params))
            row = cursor.fetchone()

            if row:
                json_filename = row[0]
                json_file_path = self.templates_dir / json_filename
                if json_file_path.exists():
                    with open(json_file_path, 'r', encoding='utf-8') as f:
                        return json.load(f)
                else:
                    print(f"Error: Template file '{json_filename}' not found for template id/name: {template_id}/{template_name}.")
                    return None
            else:
                print(f"Error: No template found with id/name: {template_id}/{template_name}.")
                return None
        except sqlite3.Error as e:
            print(f"Database error loading template (id/name: {template_id}/{template_name}): {e}")
            return None
        except IOError as e:
            print(f"File error loading template (id/name: {template_id}/{template_name}): {e}")
            return None
        except json.JSONDecodeError as e:
            print(f"JSON decode error for template (id/name: {template_id}/{template_name}): {e}")
            return None
        finally:
            if conn:
                conn.close()

    def delete_template(self, template_id: Optional[int] = None, template_name: Optional[str] = None) -> Tuple[bool, str]:
        """
        Deletes a template by its ID or name. This involves removing the DB record and the JSON file.
        """
        if template_id is None and template_name is None:
            return False, "Template ID or name must be provided for deletion."

        conn = None
        json_filename_to_delete = None
        
        select_query = "SELECT json_filename FROM templates WHERE "
        delete_query = "DELETE FROM templates WHERE "
        params = []

        if template_id is not None:
            select_query += "id = ?"
            delete_query += "id = ?"
            params.append(template_id)
        elif template_name is not None:
            select_query += "name = ?"
            delete_query += "name = ?"
            params.append(template_name)
        
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()

            # First, get the filename to delete the file
            cursor.execute(select_query, tuple(params))
            row = cursor.fetchone()
            if not row:
                return False, f"Template not found for deletion (id/name: {template_id}/{template_name})."
            json_filename_to_delete = row[0]

            # Delete DB record
            cursor.execute(delete_query, tuple(params))
            if cursor.rowcount == 0:
                # Should not happen if select found it, but as a safeguard
                conn.rollback()
                return False, f"Failed to delete template DB record (id/name: {template_id}/{template_name})."
            
            conn.commit()

            # Delete JSON file
            if json_filename_to_delete:
                json_file_path_to_delete = self.templates_dir / json_filename_to_delete
                if json_file_path_to_delete.exists():
                    try:
                        os.remove(json_file_path_to_delete)
                    except OSError as e:
                        # DB record deleted, but file deletion failed. This is a partial success/warning.
                        return True, f"Template DB record deleted, but failed to delete JSON file '{json_filename_to_delete}': {e}"
                else:
                     # DB record deleted, file was already missing.
                     return True, f"Template DB record deleted. JSON file '{json_filename_to_delete}' was not found."
            
            return True, f"Template (id/name: {template_id}/{template_name}) deleted successfully."

        except sqlite3.Error as e:
            if conn: conn.rollback()
            return False, f"Database error deleting template (id/name: {template_id}/{template_name}): {e}"
        except Exception as e:
            if conn: conn.rollback()
            return False, f"Unexpected error deleting template (id/name: {template_id}/{template_name}): {e}"
        finally:
            if conn:
                conn.close()

if __name__ == '__main__':
    # Example Usage (assuming this script is in win32com/ and user_files is a sibling)
    # For testing, you might run this from the project root if paths are adjusted or use absolute paths.
    # This example assumes the script is run in a context where 'user_files' is at the same level as the script.
    
    # Determine base path for user_files relative to this script's location
    # This makes the __main__ example more robust if run directly.
    script_dir = Path(__file__).parent
    example_base_user_dir = script_dir / "user_files"
    print(f"Using base user directory for example: {example_base_user_dir.resolve()}")

    manager = TemplateManagerWin32(base_user_dir=example_base_user_dir)

    print("\n--- Testing TemplateManagerWin32 ---")

    # Clean up previous test templates if they exist by name
    print("\nAttempting to delete pre-existing test templates...")
    manager.delete_template(template_name="Test Template Alpha")
    manager.delete_template(template_name="Test Template Beta")

    # 1. Save a new template
    print("\n1. Saving 'Test Template Alpha'...")
    styles_alpha = {"正文": {"字体": {"大小": "12pt", "名称": "宋体"}}}
    success_alpha, msg_alpha = manager.save_template(
        name="Test Template Alpha",
        style_rules_dict=styles_alpha
    )
    print(f"Save Alpha: {success_alpha}, Message: {msg_alpha}")

    # 2. Save another template
    print("\n2. Saving 'Test Template Beta'...")
    styles_beta = {"标题1": {"字体": {"大小": "16pt", "名称": "黑体", "粗体": True}}}
    success_beta, msg_beta = manager.save_template(
        name="Test Template Beta",
        style_rules_dict=styles_beta
    )
    print(f"Save Beta: {success_beta}, Message: {msg_beta}")

    # 3. List selectable templates
    print("\n3. Listing selectable templates...")
    selectable = manager.list_selectable_templates()
    if selectable:
        print("Selectable templates:")
        for tpl in selectable:
            print(f"  ID: {tpl['id']}, Name: {tpl['name']}")
    else:
        print("No selectable templates found.")

    # 4. Load a template by name
    print("\n4. Loading 'Test Template Alpha' by name...")
    loaded_alpha_by_name = manager.load_template_json(template_name="Test Template Alpha")
    if loaded_alpha_by_name:
        print(f"Loaded 'Test Template Alpha' (by name) content: {json.dumps(loaded_alpha_by_name, indent=2, ensure_ascii=False)}")
    else:
        print("'Test Template Alpha' not found by name.")

    # 5. Load a template by ID (assuming 'Test Template Beta' was saved and got an ID)
    # Find Beta's ID first from the list
    beta_id = None
    for tpl_item in selectable:
        if tpl_item['name'] == "Test Template Beta":
            beta_id = tpl_item['id']
            break
    
    if beta_id is not None:
        print(f"\n5. Loading template with ID {beta_id} ('Test Template Beta')...")
        loaded_beta_by_id = manager.load_template_json(template_id=beta_id)
        if loaded_beta_by_id:
            print(f"Loaded template ID {beta_id} content: {json.dumps(loaded_beta_by_id, indent=2, ensure_ascii=False)}")
        else:
            print(f"Template with ID {beta_id} not found.")
    else:
        print("\n5. Could not find ID for 'Test Template Beta' to test loading by ID.")

    # 6. Test saving a template with a duplicate name
    print("\n6. Attempting to save 'Test Template Alpha' again (should fail due to name conflict)...")
    success_alpha_dup, msg_alpha_dup = manager.save_template(
        name="Test Template Alpha", style_rules_dict={"正文": {"字体": {"大小": "10pt"}}}
    )
    print(f"Save Alpha Duplicate: {success_alpha_dup}, Message: {msg_alpha_dup}")


    # 7. Delete a template
    if beta_id:
        print(f"\n7. Deleting template with ID {beta_id} ('Test Template Beta')...")
        del_success, del_msg = manager.delete_template(template_id=beta_id)
        print(f"Delete Beta (ID: {beta_id}): {del_success}, Message: {del_msg}")
    
    print("\n8. Deleting 'Test Template Alpha' by name...")
    del_alpha_success, del_alpha_msg = manager.delete_template(template_name="Test Template Alpha")
    print(f"Delete Alpha (Name): {del_alpha_success}, Message: {del_alpha_msg}")


    # Verify deletion by listing again
    print("\n9. Listing templates after deletion...")
    selectable_after_delete = manager.list_selectable_templates()
    if selectable_after_delete:
        print("Selectable templates after deletion:")
        for tpl in selectable_after_delete:
            print(f"  ID: {tpl['id']}, Name: {tpl['name']}")
    else:
        print("No selectable templates found after deletion.")
    
    print("\n--- Test complete ---")