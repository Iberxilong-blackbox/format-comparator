# unit_converter.py

import re
from typing import Union, Optional, Tuple, Any

# Conversion constants
CM_TO_PT = 28.3464567
INCH_TO_PT = 72.0

class UnitConversionError(ValueError):
    """Custom exception for unit conversion errors."""
    pass

class UnitConverter:
    """
    Handles parsing and conversion of various units used in document formatting,
    primarily targeting points (pt) and relative line heights (multiple).
    """

    # Regex to parse values like "12pt", "1.5 倍", "10 cm", "2 字符"
    # Allows optional space between value and unit
    _VALUE_UNIT_REGEX = re.compile(r"^\s*(-?\d+(?:\.\d+)?)\s*([a-zA-Z\u4e00-\u9fa5]+)?\s*$")

    def parse_value(self, value: Any) -> Tuple[Optional[float], Optional[str]]:
        """
        Parses a value to extract its numerical part and unit string.

        Args:
            value: The input value. Can be numeric (int, float) or string
                   (e.g., "12pt", "1.5 倍", "10", "2 字符").

        Returns:
            A tuple containing:
            - The numerical value as a float, or None if parsing fails.
            - The unit string (lowercase), or None if no unit is found or parsing fails.
              Units like '倍', '行' are returned as 'multiple', 'line'.
              Chinese units like '磅', '厘米', '英寸', '字符' are normalized.
        """
        numeric_value: Optional[float] = None
        unit_str: Optional[str] = None

        if isinstance(value, (int, float)):
            numeric_value = float(value)
            # Unit is implicitly None unless context suggests otherwise (handled in comparison logic)
        elif isinstance(value, dict) and '值' in value and '单位' in value:
            # Handle dictionary format {'值': V, '单位': U}
            raw_val = value['值']
            raw_unit = value['单位']
            if isinstance(raw_val, (int, float)):
                numeric_value = float(raw_val)
                if isinstance(raw_unit, str):
                    unit_str = raw_unit.lower()
                    # Normalize common units (same logic as below)
                    if unit_str in ['倍', 'multiple']: unit_str = 'multiple'
                    elif unit_str in ['行', 'line']: unit_str = 'line'
                    elif unit_str in ['磅', 'pt']: unit_str = 'pt'
                    elif unit_str in ['厘米', 'cm']: unit_str = 'cm'
                    elif unit_str in ['英寸', 'inch']: unit_str = 'inch'
                    elif unit_str in ['字符', 'char']: unit_str = 'char'
                    # else: keep other units as is
                # else: unit is not a string, keep unit_str as None
            # else: value is not numeric, keep numeric_value as None
        elif isinstance(value, str):
            match = self._VALUE_UNIT_REGEX.match(value.strip())
            if match:
                try:
                    numeric_value = float(match.group(1))
                    unit_str = match.group(2)
                    if unit_str:
                        unit_str = unit_str.lower()
                        # Normalize common units
                        if unit_str in ['倍', 'multiple']:
                            unit_str = 'multiple'
                        elif unit_str in ['行', 'line']:
                            unit_str = 'line'
                        elif unit_str in ['磅', 'pt']:
                            unit_str = 'pt'
                        elif unit_str in ['厘米', 'cm']:
                            unit_str = 'cm'
                        elif unit_str in ['英寸', 'inch']:
                            unit_str = 'inch'
                        elif unit_str in ['字符', 'char']:
                            unit_str = 'char'
                        # Keep other units as is for now (e.g., 'twip' if encountered)
                except (ValueError, TypeError):
                    numeric_value = None
                    unit_str = None
            else:
                 # Handle case where string might just be a number "12"
                 try:
                     numeric_value = float(value.strip())
                 except ValueError:
                     numeric_value = None # Cannot parse as number
        # else: other types are not supported

        return numeric_value, unit_str

    def convert_value(self,
                      value: Union[int, float],
                      from_unit: Optional[str],
                      to_unit: str,
                      font_size_pt: Optional[Union[int, float]] = None
                     ) -> Optional[float]:
        """
        Converts a numerical value from one unit to another.

        Args:
            value: The numerical value to convert.
            from_unit: The unit of the input value (e.g., 'pt', 'cm', 'char', 'multiple', 'line', None).
                    Case-insensitive. None implies a base unit (often pt or multiple depending on context).
            to_unit: The target unit (e.g., 'pt', 'multiple', 'twips'). Case-insensitive.
            font_size_pt: The font size in points, required for 'char' conversion to/from 'pt' or 'twips'.

        Returns:
            The converted value as a float (for pt, multiple, line) or int (for twips),
            or None if conversion is not possible
            or not applicable (e.g., converting 'multiple' to 'pt').

        Raises:
            UnitConversionError: If 'char' conversion is attempted without font_size_pt,
                                 or if font_size_pt is invalid.
            TypeError: If value or font_size_pt (when provided) is not numeric.
        """
        if not isinstance(value, (int, float)):
            raise TypeError(f"Value must be numeric, got {type(value)}")
        if font_size_pt is not None and not isinstance(font_size_pt, (int, float)):
             raise TypeError(f"font_size_pt must be numeric when provided, got {type(font_size_pt)}")

        from_unit_lower = from_unit.lower() if isinstance(from_unit, str) else None
        to_unit_lower = to_unit.lower() if isinstance(to_unit, str) else None

        # --- Handle conversions TO 'pt' ---
        if to_unit_lower == 'pt':
            # Treat '磅' (bang) the same as 'pt' since parse_value normalizes '磅' to 'pt'
            # but here we might get '磅' directly from the template dict unit
            if from_unit_lower in ['pt', '磅']:
                return float(value)
            elif from_unit_lower == 'cm':
                return float(value) * CM_TO_PT
            elif from_unit_lower == 'inch':
                return float(value) * INCH_TO_PT
            elif from_unit_lower in ['char', '字符']: # Accept both English and Chinese keys
                if font_size_pt is None:
                    raise UnitConversionError("Font size (font_size_pt) is required for 'char'/'字符' to 'pt' conversion.")
                if font_size_pt <= 0:
                    raise UnitConversionError("Font size must be positive for 'char' to 'pt' conversion.")
                # Assuming 1 char width is approximately equal to the font size in points
                return float(value) * float(font_size_pt)
            elif from_unit_lower is None:
                 # If from_unit is None, assume it's already pt (common case for docx properties)
                 return float(value)
            elif from_unit_lower in ['multiple', 'line']:
                # Convert relative units (multiple/line) to absolute pt using font size
                if font_size_pt is None:
                     # Cannot convert without font size
                     # print(f"  [DEBUG UnitConverter] Warning: Cannot convert '{from_unit_lower}' to 'pt' without font_size_pt.")
                     return None
                if font_size_pt <= 0:
                     # print(f"  [DEBUG UnitConverter] Warning: Cannot convert '{from_unit_lower}' to 'pt' with non-positive font_size_pt: {font_size_pt}")
                     return None
                # Perform conversion: value * font_size
                converted = float(value) * float(font_size_pt)
                # print(f"  [DEBUG UnitConverter] Converted {value} {from_unit_lower} to {converted} pt using font size {font_size_pt}") # DEBUG
                return converted
            else:
                raise UnitConversionError(f"Unsupported unit conversion from '{from_unit}' to 'pt'")

        # --- Handle conversions TO 'multiple' ---
        elif to_unit_lower == 'multiple':
            # Treat '倍' (bei) the same as 'multiple'
            if from_unit_lower in ['multiple', '倍']:
                return float(value)
            elif from_unit_lower == 'line':
                 # Treat 'line' as equivalent to 'multiple' for line spacing comparison
                 return float(value)
            elif from_unit_lower is None:
                 # If from_unit is None, assume it's already multiple (common for line spacing floats)
                 return float(value)
            elif from_unit_lower in ['pt', 'cm', 'inch', 'char']:
                # Cannot directly convert absolute units to relative 'multiple'
                return None
            else:
                raise UnitConversionError(f"Unsupported unit conversion from '{from_unit}' to 'multiple'")

        # --- Handle conversions TO 'line' (mainly for consistency if needed) ---
        elif to_unit_lower == 'line':
             if from_unit_lower == 'line':
                 return float(value)
             elif from_unit_lower == 'multiple':
                 # Treat 'multiple' as equivalent to 'line'
                 return float(value)
             # Other conversions to 'line' are generally not meaningful
             return None
 
        # --- Handle conversions TO 'twips' ---
        elif to_unit_lower == 'twips':
            # First, convert the value to points (pt)
            try:
                pt_value = self.convert_value(value, from_unit, 'pt', font_size_pt=font_size_pt)
                if pt_value is None:
                    # If conversion to pt failed (e.g., trying to convert 'multiple' without font size)
                    return None
                # Convert points to twips (1 pt = 20 twips)
                return int(round(pt_value * 20)) # Return as integer
            except (UnitConversionError, TypeError) as e:
                # Re-raise or handle? For now, let the caller handle errors from the pt conversion step.
                # Or return None? Let's return None to indicate failure.
                # print(f"  [DEBUG UnitConverter] Error during intermediate pt conversion for twips: {e}")
                return None
 
        # --- Target unit not supported ---
        else:
            raise UnitConversionError(f"Unsupported target unit for conversion: '{to_unit}'")
 
         # return None # Should not be reached if logic is complete - removed as twips path returns


# Example Usage (for testing purposes)
if __name__ == "__main__":
    converter = UnitConverter()

    print("--- Testing parse_value ---")
    test_values = [12, 10.5, "15pt", "1.5 倍", "2 cm ", " 3 inch", "2字符", "12", "-5", "invalid", "1 line", "1.15multiple"]
    for tv in test_values:
        num_val, unit = converter.parse_value(tv)
        print(f"Input: {tv!r} -> Parsed Value: {num_val}, Parsed Unit: {unit!r}")

    print("\n--- Testing convert_value (to pt) ---")
    print(f"10 cm to pt: {converter.convert_value(10, 'cm', 'pt')}")
    print(f"2 inches to pt: {converter.convert_value(2, 'inch', 'pt')}")
    print(f"12 pt to pt: {converter.convert_value(12, 'pt', 'pt')}")
    print(f"15 (None unit) to pt: {converter.convert_value(15, None, 'pt')}") # Assumes pt
    try:
        print(f"2 char (no font size) to pt: {converter.convert_value(2, 'char', 'pt')}")
    except UnitConversionError as e:
        print(f"Error converting char: {e}")
    print(f"2 char (font 12pt) to pt: {converter.convert_value(2, 'char', 'pt', font_size_pt=12)}")
    print(f"2 char (font 10.5pt) to pt: {converter.convert_value(2, 'char', 'pt', font_size_pt=10.5)}")
    print(f"1.5 multiple to pt: {converter.convert_value(1.5, 'multiple', 'pt')}") # Should be None
    print(f"1 line to pt: {converter.convert_value(1, 'line', 'pt')}") # Should be None

    print("\n--- Testing convert_value (to multiple) ---")
    print(f"1.5 multiple to multiple: {converter.convert_value(1.5, 'multiple', 'multiple')}")
    print(f"1 line to multiple: {converter.convert_value(1, 'line', 'multiple')}")
    print(f"1.15 (None unit) to multiple: {converter.convert_value(1.15, None, 'multiple')}") # Assumes multiple
    print(f"12 pt to multiple: {converter.convert_value(12, 'pt', 'multiple')}") # Should be None

    print("\n--- Testing convert_value (to twips) ---")
    print(f"12 pt to twips: {converter.convert_value(12, 'pt', 'twips')}") # Expected: 240
    print(f"1 cm to twips: {converter.convert_value(1, 'cm', 'twips')}") # Expected: ~567
    print(f"0.5 inch to twips: {converter.convert_value(0.5, 'inch', 'twips')}") # Expected: 720
    print(f"2 char (font 12pt) to twips: {converter.convert_value(2, 'char', 'twips', font_size_pt=12)}") # Expected: 480
    try:
        print(f"1.5 multiple to twips: {converter.convert_value(1.5, 'multiple', 'twips')}") # Expected: None (or error)
    except UnitConversionError as e:
        print(f"Error converting multiple to twips: {e}") # Expected if error raised
    print(f"1.5 multiple (font 10pt) to twips: {converter.convert_value(1.5, 'multiple', 'twips', font_size_pt=10)}") # Expected: 300