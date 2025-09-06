import pandas as pd
import re
from typing import List, Dict, Any, Optional

class ExcelToTextConverter:
    def __init__(self):
        pass
    
    def convert_excel_to_text(self, excel_file_path: str) -> str:
        """Convert Excel file to text format"""
        try:
            # Read all sheets from Excel file
            excel_sheets = pd.read_excel(excel_file_path, sheet_name=None)
            
            # Find the MAPPING sheet (case-insensitive)
            mapping_sheet = None
            for sheet_name, sheet_data in excel_sheets.items():
                if sheet_name.upper() == 'MAPPING':
                    mapping_sheet = sheet_data
                    break
            
            if mapping_sheet is None:
                raise Exception("MAPPING sheet not found in Excel file")
            
            # Find required columns dynamically
            begin_col = self._find_column(mapping_sheet, ['Begin'])
            beta_field_col = self._find_column(mapping_sheet, ['BETA Field Name', 'Beta Field Name', 'beta field', 'field name'])
            mapping_instructions_col = self._find_column(mapping_sheet, ['Mapping Instructions for Programmer', 'mapping instructions'])
            
            if begin_col is None or beta_field_col is None:
                raise Exception("Required columns (Begin, BETA Field Name) not found")
            
            # Generate header
            header = self._generate_header()
            
            # Process rows
            field_mappings = []
            for index, row in mapping_sheet.iterrows():
                try:
                    begin_value = str(row[begin_col]).strip()
                    beta_field = str(row[beta_field_col]).strip()
                    
                    # Skip if begin value is not a valid number
                    try:
                        begin_num = float(begin_value)
                        # Convert to integer and format as 4-digit string
                        begin_value = f"{int(begin_num):04d}"
                    except ValueError:
                        continue
                    
                    # Generate DEMO field name
                    demo_field = self._generate_demo_field_name(beta_field, begin_value)
                    
                    # Get mapping option from instructions
                    mapping_option = "01"  # default
                    if mapping_instructions_col and pd.notna(row[mapping_instructions_col]):
                        instructions = str(row[mapping_instructions_col])
                        mapping_option = self._extract_mapping_option(instructions)
                    
                    field_mappings.append({
                        'begin': begin_value,
                        'mapping_option': mapping_option,
                        'demo_field': demo_field
                    })
                    
                except Exception:
                    continue  # Skip problematic rows
            
            # Generate output
            output_lines = header + self._generate_field_records(field_mappings)
            
            return '\n'.join(output_lines)
            
        except Exception as e:
            raise Exception(f"Error converting Excel file: {str(e)}")
    
    def _find_column(self, df: pd.DataFrame, possible_names: List[str]) -> Optional[str]:
        """Find column by name (case-insensitive)"""
        # First try exact matches
        for col in df.columns:
            col_str = str(col).strip()
            for name in possible_names:
                if col_str.lower() == name.lower():
                    return col
        
        # Then try partial matches
        for col in df.columns:
            col_lower = str(col).lower().strip()
            for name in possible_names:
                if name.lower() in col_lower:
                    return col
        return None
    
    def _generate_header(self) -> List[str]:
        """Generate the header section"""
        return [
            "***********************************************************************",
            "****",
            "**** DATA FILE FOR LPL/CUS PERSHING DEMO TRANSLATION",
            "****",
            "***********************************************************************",
            "****  FILE CONVERSION SUMMARY OF RECORD TYPES USED:",
            "****",
            "****  0000 - EDIT CONTROL FIELDS",
            "****",
            "***********************************************************************",
            "**** RECORD TYPE 0000 - EDIT CONTROL FIELDS",
            "****",
            "****   COL 06-09     FIELD STARTING POSITION",
            "****   COL 12-13     MAPPING OPTION (01 THRU 99)",
            "****   COL 16-30     DEFAULT VALUE #1 (IF APPLICABLE)",
            "****   COL 33-47     DEFAULT VALUE #2 (IF APPLICABLE)",
            "****   COL 50-72     DEMO FIELD NAME",
            "****",
            "**** ****  **  ***************  ***************  *** LEVEL 1 FIELDS  **"
        ]
    
    def _generate_demo_field_name(self, beta_field: str, position: str) -> str:
        """Generate DEMO field name from BETA field name"""
        if not beta_field or beta_field == 'nan' or beta_field.strip() == '':
            return f"DEMO-POSITION-{position.zfill(4)}"
        
        # If it already starts with DEMO-, return as is
        if beta_field.startswith('DEMO-'):
            return beta_field
        
        # Convert common prefixes to DEMO-
        if beta_field.startswith('PER-'):
            return beta_field.replace('PER-', 'DEMO-', 1)
        elif beta_field.startswith('IRAS-'):
            return beta_field.replace('IRAS-', 'DEMO-', 1)
        elif beta_field.startswith('DIST-'):
            return beta_field.replace('DIST-', 'DEMO-', 1)
        elif beta_field.startswith('DOCA-'):
            return beta_field.replace('DOCA-', 'DEMO-', 1)
        else:
            return f"DEMO-{beta_field}"
    
    def _extract_mapping_option(self, instructions: str) -> str:
        """Extract mapping option from instructions text"""
        if not instructions or instructions == 'nan':
            return "01"
        
        # Look for patterns like "1.", "2.", etc.
        match = re.search(r'^(\d+)\.', instructions.strip())
        if match:
            number = int(match.group(1))
            return f"{number:02d}"
        
        return "01"  # default
    
    def _generate_field_records(self, field_mappings: List[Dict[str, Any]]) -> List[str]:
        """Generate the field record lines"""
        records = []
        for mapping in field_mappings:
            begin = mapping['begin'].zfill(4)
            mapping_option = mapping['mapping_option']
            demo_field = mapping['demo_field']
            
            # Format: 0000 0001  01                                    DEMO-FIELD-NAME
            # Position: 0000 (4) + space + begin (4) + 2 spaces + mapping (2) + 35 spaces + field name
            record = f"0000 {begin}  {mapping_option}                                    {demo_field}"
            records.append(record)
        
        return records

# Test function
if __name__ == "__main__":
    converter = ExcelToTextConverter()
    try:
        result = converter.convert_excel_to_text("Input1.xlsx")
        print("SUCCESS: Conversion completed")
        print("First 25 lines:")
        lines = result.split('\n')
        for i, line in enumerate(lines[:25]):
            print(f"{i+1:2d}: {line}")
        print(f"\nTotal lines: {len(lines)}")
    except Exception as e:
        print(f"ERROR: {e}")
