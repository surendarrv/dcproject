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
            
            # Find the MAPPING sheet (flexible matching - contains "mapping" anywhere in name)
            mapping_sheet = None
            for sheet_name, sheet_data in excel_sheets.items():
                if 'mapping' in sheet_name.lower():
                    mapping_sheet = sheet_data
                    break
            
            if mapping_sheet is None:
                raise Exception("No sheet containing 'mapping' found in Excel file")
            
            # Find required columns dynamically (case-insensitive and flexible)
            begin_col = self._find_column(mapping_sheet, ['begin', 'position', 'start'])
            beta_field_col = self._find_column(mapping_sheet, ['BETA Field Name'])
            mapping_instructions_col = self._find_column(mapping_sheet, ['mapping instructions for programmer', 'mapping instructions', 'instructions', 'mapping'])
            
            # Debug: Print which columns were found
            print(f"DEBUG: Found columns - Begin: '{begin_col}', BETA Field: '{beta_field_col}', Instructions: '{mapping_instructions_col}'")
            print(f"DEBUG: Available columns: {list(mapping_sheet.columns)}")
            
            # Debug: Show first few rows of BETA Field Name column
            if beta_field_col:
                print(f"DEBUG: First 10 rows of '{beta_field_col}' column:")
                for i in range(min(10, len(mapping_sheet))):
                    value = mapping_sheet.iloc[i][beta_field_col]
                    print(f"  Row {i}: '{value}' (type: {type(value)})")
            
            if begin_col is None or beta_field_col is None:
                raise Exception("Required columns (Begin/Position, BETA Field Name/Field Name) not found")
            
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
                    
                    # Generate DEMO field name - use BETA field name as-is if available
                    demo_field = self._generate_demo_field_name(beta_field, begin_value)
                    
                    # Skip rows that don't have a BETA field name
                    if not demo_field:
                        continue
                    
                    # Get mapping option from instructions
                    mapping_option = "01"  # default
                    table_reference = None
                    if mapping_instructions_col and pd.notna(row[mapping_instructions_col]):
                        instructions = str(row[mapping_instructions_col])
                        mapping_option = self._extract_mapping_option(instructions)
                        table_reference = self._extract_table_reference(instructions)
                        
                        # Debug: Print table references found
                        if table_reference:
                            print(f"DEBUG: Found table reference '{table_reference}' in row {index} for field '{demo_field}'")
                    
                    field_mappings.append({
                        'begin': begin_value,
                        'mapping_option': mapping_option,
                        'demo_field': demo_field,
                        'table_reference': table_reference
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
        """Use BETA field name as-is from the mapping sheet with special handling"""
        # Clean the beta_field string
        if pd.isna(beta_field) or beta_field is None:
            return ""  # Return empty string if no BETA field name
        
        beta_field = str(beta_field).strip()
        
        # Handle empty or invalid values
        if not beta_field or beta_field == 'nan' or beta_field == '' or beta_field == 'None':
            return ""  # Return empty string if no BETA field name
        
        # Special handling for first row
        if beta_field == "DE-DEMO=KEY-AREA":
            return "DEMO-KEY-ACCT"
        
        # Return the BETA field name as-is
        return beta_field
    
    def _extract_mapping_option(self, instructions: str) -> str:
        """Extract mapping option from instructions text"""
        if not instructions or instructions == 'nan':
            return "01"
        
        # Look for patterns like "1.", "2.", etc.
        match = re.search(r'^(\d+)\.', instructions.strip())
        if match:
            number = int(match.group(1))
            return f"{number:02d}"
        
        return "01"
    
    def _extract_table_reference(self, instructions: str) -> str:
        """Extract table reference from instructions text"""
        if not instructions or instructions == 'nan':
            return None
        
        # Look for table references like "Table A", "TABLE B", "table C", etc.
        match = re.search(r'table\s+([A-Za-z0-9]+)', instructions, re.IGNORECASE)
        if match:
            return match.group(1).lower()
        
        return None  # default
    
    def _generate_field_records(self, field_mappings: List[Dict[str, Any]]) -> List[str]:
        """Generate the field record lines"""
        records = []
        for mapping in field_mappings:
            begin = mapping['begin'].zfill(4)
            mapping_option = mapping['mapping_option']
            demo_field = mapping['demo_field']
            table_reference = mapping.get('table_reference')
            
            # Format: 0000 0001  01                                    DEMO-FIELD-NAME
            # Position: 0000 (4) + space + begin (4) + 2 spaces + mapping (2) + 35 spaces + field name
            record = f"0000 {begin}  {mapping_option}                                    {demo_field}"
            
            # Add table reference comment if present
            if table_reference:
                record += f"  # table {table_reference}"
            
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
