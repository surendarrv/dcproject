import pandas as pd
import re
from typing import List, Dict, Any, Optional

class ExcelToTextConverter:
    def __init__(self):
        pass
    
    def convert_excel_to_text(self, excel_file_path: str, apply_prefix: bool = True, mode: str = 'DEMO/IRAS/DIST-normal') -> str:
        """Convert Excel file to text format"""
        try:
            # Read all sheets from Excel file
            excel_sheets = pd.read_excel(excel_file_path, sheet_name=None)
            
            # Get list of available sheet names for table reference checking
            available_sheets = [sheet_name.lower().strip() for sheet_name in excel_sheets.keys()]
            
            # Extract title from Cover page sheet (first row, first cell)
            title = "DATA FILE FOR CONVERTION"  # default fallback
            cover_page_sheet = None
            for sheet_name, sheet_data in excel_sheets.items():
                sheet_name_lower = sheet_name.lower().replace(' ', '')
                if 'cover' in sheet_name_lower and 'page' in sheet_name_lower:
                    cover_page_sheet = sheet_data
                    break
            
            if cover_page_sheet is not None and len(cover_page_sheet) > 0:
                # Get the first column header (which might contain the title)
                first_column_header = cover_page_sheet.columns[0]
                if pd.notna(first_column_header) and str(first_column_header).strip():
                    title = str(first_column_header).strip()
                    print(f"DEBUG: Using title from Cover page column header: '{title}'")
                else:
                    # Fallback to first data cell if column header is empty
                    first_cell_value = cover_page_sheet.iloc[0, 0]
                    if pd.notna(first_cell_value) and str(first_cell_value).strip():
                        title = str(first_cell_value).strip()
                        print(f"DEBUG: Using title from Cover page first cell: '{title}'")
                    else:
                        print("DEBUG: Cover page first cell is empty, using default title")
            else:
                print("DEBUG: No Cover page sheet found, using default title")
            
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
            
            # Detect prefix from column header
            detected_prefix = self._detect_column_prefix(beta_field_col)
            column_prefix = None
            
            if detected_prefix:
                print(f"DEBUG: Detected prefix '{detected_prefix}' in column header: '{beta_field_col}'")
                if apply_prefix:
                    column_prefix = detected_prefix
                    print(f"INFO: Applying prefix '{detected_prefix}' to all field names in this column")
                else:
                    print(f"INFO: Prefix '{detected_prefix}' detected but not applied (user choice)")
            else:
                print(f"DEBUG: No prefix detected in column header: '{beta_field_col}'")
            
            # Process rows first to analyze field patterns
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
                    demo_field = self._generate_demo_field_name(beta_field, begin_value, column_prefix)
                    
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
            
            # Analyze field patterns and generate dynamic header
            dynamic_field_name = self._analyze_field_name_pattern(field_mappings)
            header = self._generate_header(title, dynamic_field_name)
            
            # Generate output
            output_lines = header + self._generate_field_records(field_mappings, available_sheets)
            
            # Process table references and append at the end
            table_referenced_fields = [mapping for mapping in field_mappings if mapping.get('table_reference')]
            if table_referenced_fields:
                if mode == 'DOCA':
                    # For DOCA mode, perform intelligent DOCA extraction
                    table_output = self._generate_doca_table_reference_output(table_referenced_fields, excel_sheets, available_sheets)
                elif mode == 'IRAS':
                    # For IRAS mode, process DSTRBTN SORC table references
                    table_output = self._generate_iras_table_reference_output(table_referenced_fields, excel_sheets, available_sheets)
                else:
                    # For other modes (DEMO), use normal table reference processing
                    table_output = self._generate_table_reference_output(table_referenced_fields, excel_sheets, available_sheets)
                output_lines.extend(table_output)
            
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
    
    def _generate_header(self, title: str = "DATA FILE FOR LPL/CUS PERSHING DEMO TRANSLATION", field_name: str = "DEMO FIELD NAME") -> List[str]:
        """Generate the header section with dynamic field name"""
        # Replace newline characters with spaces in the title
        title = title.replace('\n', ' ').replace('\r', ' ')
        return [
            "***********************************************************************",
            "****",
            f"**** {title}",
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
            f"****   COL 50-72     {field_name}",
            "****",
            "**** ****  **  ***************  ***************  *** LEVEL 1 FIELDS  **"
        ]
    
    def _analyze_field_name_pattern(self, mapping_fields: List[Dict[str, Any]]) -> str:
        """Analyze LEVEL 1 FIELD patterns to determine the appropriate field name"""
        if not mapping_fields:
            return "DEMO FIELD NAME"
        
        # Extract all LEVEL 1 FIELD values (stored as 'demo_field' in the mapping)
        level1_fields = []
        for field in mapping_fields:
            level1_field = field.get('demo_field', '').strip()
            if level1_field:
                level1_fields.append(level1_field)
        
        if not level1_fields:
            return "DEMO FIELD NAME"
        
        print(f"DEBUG: Analyzing {len(level1_fields)} LEVEL 1 FIELD values: {level1_fields[:10]}...")
        
        # Count prefixes
        prefix_counts = {}
        for field in level1_fields:
            # Extract prefix (everything before the first non-alphanumeric character or end of string)
            prefix = ""
            for char in field:
                if char.isalnum():
                    prefix += char.upper()
                else:
                    break
            
            if prefix:
                prefix_counts[prefix] = prefix_counts.get(prefix, 0) + 1
        
        print(f"DEBUG: Prefix counts: {prefix_counts}")
        
        # Find the most common prefix
        if not prefix_counts:
            return "DEMO FIELD NAME"
        
        most_common_prefix = max(prefix_counts, key=prefix_counts.get)
        most_common_count = prefix_counts[most_common_prefix]
        total_fields = len(level1_fields)
        
        print(f"DEBUG: Most common prefix: '{most_common_prefix}' ({most_common_count}/{total_fields} fields)")
        
        # Determine field name based on prefix patterns
        if most_common_prefix.startswith('DEMO'):
            return "DEMO FIELD NAME"
        elif most_common_prefix.startswith('IRAS'):
            return "IRAS FIELD NAME"
        elif most_common_prefix.startswith('DOCA'):
            return "DOCA FIELD NAME"
        elif most_common_prefix.startswith('DIST'):
            return "DIST FIELD NAME"
        else:
            # For other prefixes, use the prefix itself
            return f"{most_common_prefix} FIELD NAME"
    
    def _detect_column_prefix(self, column_name: str) -> str:
        """Detect if a column header indicates a prefix that should be applied to all field names"""
        if not column_name:
            return None
            
        column_name = str(column_name).strip()
        
        # Look for patterns like "BETA Field Name DIST-", "BETA Field Name IRAS-", etc.
        # Check if the column name ends with a known prefix followed by a dash
        known_prefixes = ['DIST', 'IRAS', 'DEMO', 'PER', 'CDC', 'LPL', 'CUSO', 'RDW']
        
        for prefix in known_prefixes:
            # Look for patterns like "BETA Field Name\nDIST-" or "BETA Field Name DIST-"
            if f"{prefix}-" in column_name.upper():
                return prefix
        
        return None
    
    def detect_prefix_in_file(self, excel_file_path: str) -> dict:
        """Detect if the Excel file has a prefix in the BETA Field Name column header"""
        try:
            # Read all sheets from Excel file
            excel_sheets = pd.read_excel(excel_file_path, sheet_name=None)
            available_sheets = list(excel_sheets.keys())
            
            # Find mapping sheet
            mapping_sheet = None
            for sheet_name in available_sheets:
                if 'mapping' in sheet_name.lower():
                    mapping_sheet = excel_sheets[sheet_name]
                    break
            
            if mapping_sheet is None:
                return {'has_prefix': False, 'prefix': None, 'column_name': None}
            
            # Find BETA Field Name column
            beta_field_col = self._find_column(mapping_sheet, ['BETA Field Name'])
            
            if beta_field_col is None:
                return {'has_prefix': False, 'prefix': None, 'column_name': None}
            
            # Detect prefix
            detected_prefix = self._detect_column_prefix(beta_field_col)
            
            return {
                'has_prefix': detected_prefix is not None,
                'prefix': detected_prefix,
                'column_name': beta_field_col
            }
            
        except Exception as e:
            print(f"ERROR: Failed to detect prefix: {e}")
            return {'has_prefix': False, 'prefix': None, 'column_name': None}
    
    def _generate_demo_field_name(self, beta_field: str, position: str, column_prefix: str = None) -> str:
        """Use BETA field name as-is from the mapping sheet with special handling for prefixes"""
        # Clean the beta_field string
        if pd.isna(beta_field) or beta_field is None:
            return ""  # Return empty string if no BETA field name
        
        beta_field = str(beta_field).strip()
        
        # Handle empty or invalid values
        if not beta_field or beta_field == 'nan' or beta_field == '' or beta_field == 'None':
            return ""  # Return empty string if no BETA field name
        
        # Special handling for KEY-AREA pattern
        if beta_field.endswith("KEY-AREA"):
            # Check for specific patterns that should remain as-is
            # Look for IRAS, DIST, DOCO, or DOCA anywhere in the field name before KEY-AREA
            beta_field_upper = beta_field.upper()
            if any(prefix in beta_field_upper for prefix in ["IRAS", "DIST", "DOCO", "DOCA"]):
                return beta_field  # Return as-is for fields containing IRAS, DIST, DOCO, or DOCA
            else:
                return "DEMO-KEY-ACCT"  # Default conversion for other KEY-AREA patterns
        
        # If column prefix is detected and provided, apply it
        if column_prefix:
            # Don't add prefix if the field already has it
            if not beta_field.upper().startswith(f"{column_prefix}-"):
                return f"{column_prefix}-{beta_field}"
        
        # Handle prefixes like "DIST-", "IRAS-", etc.
        # If the field name contains a prefix (ends with a dash), preserve it
        if '-' in beta_field and not beta_field.startswith('DEMO-'):
            # Check if it's a prefix pattern (like "DIST-", "IRAS-", etc.)
            parts = beta_field.split('-', 1)  # Split only on first dash
            if len(parts) == 2 and len(parts[0]) <= 10:  # Reasonable prefix length
                prefix = parts[0].upper()
                field_name = parts[1]
                
                # Common prefixes that should be preserved
                valid_prefixes = ['DIST', 'IRAS', 'DEMO', 'PER', 'CDC', 'LPL', 'CUSO', 'RDW']
                
                if prefix in valid_prefixes:
                    return f"{prefix}-{field_name}"
        
        # Return the BETA field name as-is (preserving any existing prefixes)
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
        
        # Look for DSTRBTN SORC references first
        if re.search(r'dstrbtn\s+sorc', instructions, re.IGNORECASE):
            return 'dstrbtn sorc'
        
        # Look for table references like "Table A", "TABLE B", "table C", etc.
        match = re.search(r'table\s+([A-Za-z0-9]+)', instructions, re.IGNORECASE)
        if match:
            return match.group(1).lower()
        
        return None  # default
    
    def _check_table_sheet_exists(self, table_reference: str, available_sheets: List[str]) -> bool:
        """Check if a table reference sheet exists in the Excel file"""
        print(f"DEBUG: _check_table_sheet_exists called with table_reference='{table_reference}'")
        print(f"DEBUG: Available sheets: {available_sheets}")
        
        if not table_reference:
            return False
        
        # Check for various possible sheet name formats
        table_name_variations = [
            f"table {table_reference}",
            f"table{table_reference}",
            f"table_{table_reference}",
            f"table-{table_reference}",
            f"table {table_reference.upper()}",
            f"table{table_reference.upper()}",
            f"table_{table_reference.upper()}",
            f"table-{table_reference.upper()}",
            table_reference,
            table_reference.upper(),
            f"t{table_reference}",
            f"t_{table_reference}",
            f"t-{table_reference}",
            f"t{table_reference.upper()}",
            f"t_{table_reference.upper()}",
            f"t-{table_reference.upper()}"
        ]
        
        # Also check for sheets that contain the pattern "table" + table_reference (case insensitive)
        for sheet_name in available_sheets:
            sheet_lower = sheet_name.lower()
            table_ref_lower = table_reference.lower()
            # Look for pattern: table + table_reference (case insensitive)
            if 'table' in sheet_lower and table_ref_lower in sheet_lower:
                print(f"DEBUG: Found matching sheet: '{sheet_name}' for table reference '{table_reference}'")
                return True
        
        print(f"DEBUG: No matching sheet found for table reference '{table_reference}'")
        return False
    
    def _generate_field_records(self, field_mappings: List[Dict[str, Any]], available_sheets: List[str]) -> List[str]:
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
            
            # Add hidden marker for table reference rows (for frontend styling)
            if table_reference:
                record += " <!-- TABLE_REF -->"
            
            records.append(record)
        
        return records
    
    def _generate_table_reference_output(self, table_referenced_fields: List[Dict[str, Any]], excel_sheets: Dict[str, pd.DataFrame], available_sheets: List[str]) -> List[str]:
        """Generate table reference output for all table-referenced fields"""
        table_output = []
        
        # Process each table-referenced field
        for field_mapping in table_referenced_fields:
            begin = field_mapping['begin']
            demo_field = field_mapping['demo_field']
            table_reference = field_mapping['table_reference']
            
            # Check if table sheet exists
            if not self._check_table_sheet_exists(table_reference, available_sheets):
                continue
            
            # Find the table sheet
            table_sheet = self._find_table_sheet(table_reference, excel_sheets)
            if table_sheet is None:
                continue
            
            # Generate table data for this field
            field_table_output = self._process_single_table_reference(begin, demo_field, table_reference, table_sheet)
            if field_table_output:
                table_output.extend(field_table_output)
        
        return table_output
    
    def _generate_doca_table_reference_output(self, table_referenced_fields: List[Dict[str, Any]], excel_sheets: Dict[str, pd.DataFrame], available_sheets: List[str]) -> List[str]:
        """Generate DOCA mode table reference output with intelligent extraction"""
        table_output = []
        
        # Add a header for table references
        table_output.append("")
        table_output.append("**** TABLE REFERENCE DATA (DOCA MODE) ****")
        table_output.append("")
        
        # Process each table-referenced field
        for field_mapping in table_referenced_fields:
            begin = field_mapping['begin']
            demo_field = field_mapping['demo_field']
            table_reference = field_mapping['table_reference']
            
            # Add field header
            table_output.append(f"**** {demo_field} {begin}{table_reference.upper()} (TABLE {table_reference.upper()} #1-TOP ON MAPPING)")
            
            # Check if table sheet exists
            if not self._check_table_sheet_exists(table_reference, available_sheets):
                table_output.append("**** ALERT: Table sheet not found ****")
                table_output.append("")
                continue
            
            # Find the table sheet
            table_sheet = self._find_table_sheet(table_reference, excel_sheets)
            if table_sheet is None:
                table_output.append("**** ALERT: Table sheet not found ****")
                table_output.append("")
                continue
            
            # Perform DOCA-specific extraction
            doca_output = self._extract_doca_table_data(table_sheet)
            if doca_output:
                table_output.extend(doca_output)
            else:
                table_output.append("*** No DOCA data found in table ***")
            
            table_output.append("")
        
        return table_output
    
    def _extract_doca_table_data(self, table_sheet: pd.DataFrame) -> List[str]:
        """Extract DOCA table data with intelligent record type detection"""
        doca_output = []
        
        try:
            # Detect columns dynamically
            column_mapping = self._detect_doca_columns(table_sheet)
            if not column_mapping:
                print("DEBUG: Could not detect DOCA columns")
                return []
            
            print(f"DEBUG: Detected DOCA columns: {column_mapping}")
            
            # Find data rows (skip header rows)
            data_rows = self._find_doca_data_rows(table_sheet)
            if not data_rows:
                print("DEBUG: No data rows found in DOCA table")
                return []
            
            print(f"DEBUG: Found {len(data_rows)} data rows")
            
            # Group rows by record type and generate output
            record_groups = self._classify_doca_records(data_rows, column_mapping)
            
            for record_type, records in record_groups.items():
                if records:  # Only process if we have records
                    record_output = self._format_doca_record_type(record_type, records, column_mapping)
                    doca_output.extend(record_output)
                    doca_output.append("")  # Add spacing between record types
            
            return doca_output
            
        except Exception as e:
            print(f"DEBUG: Error in DOCA extraction: {e}")
            return [f"**** ERROR: DOCA extraction failed - {str(e)} ****"]
    
    def _detect_doca_columns(self, table_sheet: pd.DataFrame) -> Dict[str, int]:
        """Detect DOCA columns by analyzing content patterns"""
        column_mapping = {}
        
        # Analyze each column to identify its purpose
        for col_idx, col_name in enumerate(table_sheet.columns):
            col_data = table_sheet.iloc[:, col_idx].dropna()
            
            if len(col_data) == 0:
                continue
                
            # Convert to strings for analysis
            col_strings = [str(val).strip() for val in col_data if str(val).strip()]
            
            if not col_strings:
                continue
            
            # Detect NAINST column (single letters: A, C, E, G, I, J, K, Q, S, T, U)
            if self._is_nainst_column(col_strings):
                column_mapping['NAINST'] = col_idx
                print(f"DEBUG: Detected NAINST column at index {col_idx}")
            
            # Detect NA-ACCT-CAT column (3-digit codes)
            elif self._is_acct_cat_column(col_strings):
                column_mapping['NA-ACCT-CAT'] = col_idx
                print(f"DEBUG: Detected NA-ACCT-CAT column at index {col_idx}")
            
            # Detect NA-PRODUCT-CLASS column (codes like BBK, QBK, W9, etc.)
            elif self._is_product_class_column(col_strings):
                column_mapping['NA-PRODUCT-CLASS'] = col_idx
                print(f"DEBUG: Detected NA-PRODUCT-CLASS column at index {col_idx}")
            
            # Detect FRD-DOC-CODE column (document codes)
            elif self._is_doc_code_column(col_strings):
                column_mapping['FRD-DOC-CODE'] = col_idx
                print(f"DEBUG: Detected FRD-DOC-CODE column at index {col_idx}")
            
            # Detect FRD-STATUS-CODE column (R, X)
            elif self._is_status_code_column(col_strings):
                column_mapping['FRD-STATUS-CODE'] = col_idx
                print(f"DEBUG: Detected FRD-STATUS-CODE column at index {col_idx}")
            
            # Detect FRD-REEL-NO column (00000, numbers)
            elif self._is_reel_no_column(col_strings):
                column_mapping['FRD-REEL-NO'] = col_idx
                print(f"DEBUG: Detected FRD-REEL-NO column at index {col_idx}")
            
            # Detect FNP-SIGNED-CYMD column (Y, N)
            elif self._is_signed_cymd_column(col_strings):
                column_mapping['FNP-SIGNED-CYMD'] = col_idx
                print(f"DEBUG: Detected FNP-SIGNED-CYMD column at index {col_idx}")
            
            # Detect Warning Report column (Y, N)
            elif self._is_warning_report_column(col_strings):
                column_mapping['Warning-Report'] = col_idx
                print(f"DEBUG: Detected Warning-Report column at index {col_idx}")
        
        return column_mapping
    
    def _is_nainst_column(self, col_strings: List[str]) -> bool:
        """Check if column contains NAINST values (single letters)"""
        nainst_values = {'A', 'C', 'E', 'G', 'I', 'J', 'K', 'Q', 'S', 'T', 'U'}
        matches = sum(1 for val in col_strings if val in nainst_values)
        return matches >= len(col_strings) * 0.7  # 70% match threshold
    
    def _is_acct_cat_column(self, col_strings: List[str]) -> bool:
        """Check if column contains NA-ACCT-CAT values (3-digit codes)"""
        acct_cat_values = {'002', '004', '011', '012', '015', '016', '017', '018', '019', '020', '026', '027', '028', '029', '030', '032', '033'}
        matches = sum(1 for val in col_strings if val in acct_cat_values)
        return matches >= len(col_strings) * 0.7
    
    def _is_product_class_column(self, col_strings: List[str]) -> bool:
        """Check if column contains NA-PRODUCT-CLASS values"""
        product_class_values = {'BBK', 'QBK', 'W9', 'CA', 'DVP', 'INAA', 'INAS'}
        matches = sum(1 for val in col_strings if val in product_class_values)
        return matches >= len(col_strings) * 0.7
    
    def _is_doc_code_column(self, col_strings: List[str]) -> bool:
        """Check if column contains FRD-DOC-CODE values"""
        doc_code_values = {'A1', 'A2', 'CR', 'CRT', 'ICA', 'LLC', 'LT', 'NCR', 'PA', 'SPC', 'TC', 'TOD', 'ESA', 'BOI', 'CO02', 'TDTK'}
        matches = sum(1 for val in col_strings if val in doc_code_values)
        return matches >= len(col_strings) * 0.7
    
    def _is_status_code_column(self, col_strings: List[str]) -> bool:
        """Check if column contains FRD-STATUS-CODE values (R, X)"""
        status_values = {'R', 'X'}
        matches = sum(1 for val in col_strings if val in status_values)
        return matches >= len(col_strings) * 0.7
    
    def _is_reel_no_column(self, col_strings: List[str]) -> bool:
        """Check if column contains FRD-REEL-NO values (numbers, 00000)"""
        matches = sum(1 for val in col_strings if val.isdigit() or val == '00000')
        return matches >= len(col_strings) * 0.7
    
    def _is_signed_cymd_column(self, col_strings: List[str]) -> bool:
        """Check if column contains FNP-SIGNED-CYMD values (Y, N)"""
        signed_values = {'Y', 'N'}
        matches = sum(1 for val in col_strings if val in signed_values)
        return matches >= len(col_strings) * 0.7
    
    def _is_warning_report_column(self, col_strings: List[str]) -> bool:
        """Check if column contains Warning Report values (Y, N)"""
        warning_values = {'Y', 'N'}
        matches = sum(1 for val in col_strings if val in warning_values)
        return matches >= len(col_strings) * 0.7
    
    def _find_doca_data_rows(self, table_sheet: pd.DataFrame) -> List[int]:
        """Find rows that contain actual data (skip header rows)"""
        data_rows = []
        
        for row_idx in range(len(table_sheet)):
            row = table_sheet.iloc[row_idx]
            
            # Skip rows that are mostly empty or contain header-like text
            non_empty_cells = sum(1 for val in row if pd.notna(val) and str(val).strip())
            
            if non_empty_cells >= 3:  # At least 3 non-empty cells
                # Check if this looks like a data row (not a header)
                row_text = ' '.join(str(val) for val in row if pd.notna(val)).lower()
                
                # Skip rows with header-like keywords
                header_keywords = ['record type', 'column', 'input', 'output', 'mapping', 'document', 'field']
                if not any(keyword in row_text for keyword in header_keywords):
                    data_rows.append(row_idx)
        
        return data_rows
    
    def _classify_doca_records(self, data_rows: List[int], column_mapping: Dict[str, int]) -> Dict[str, List[int]]:
        """Classify data rows into different record types"""
        record_groups = {
            '01': [],  # NA-PRODUCT-CLASS related
            '02': [],  # NAINST/NA-ACCT-CAT related
            '03': [],  # MSI related
            '04': [],  # NAINST related
            '05': [],  # NAINST/PRODUCT-CLASS related
            '08': [],  # NA-ACCT-CAT related
            '10': [],  # NA-TRAN-LVL related
            'N1': [],  # NAINST/MSI negative logic
            'N2': [],  # NAOPTN-I-O negative logic
            '14': [],  # NALINES search text
            '17': [],  # TABLE A mapping
            '19': []   # NAINST/PRODUCT-CLASS related
        }
        
        # For now, classify all rows as type 01 (will be refined)
        record_groups['01'] = data_rows
        
        return record_groups
    
    def _format_doca_record_type(self, record_type: str, record_rows: List[int], column_mapping: Dict[str, int]) -> List[str]:
        """Format a specific record type with its data"""
        output = []
        
        # Add record type header
        if record_type == '01':
            output.extend([
                "**************************************************************************",
                "** RECORD TYPE 01 - NA-PRODUCT-CLASS RELATED DOCUMENTS",
                "**        (IF NA-PRODUCT-CLASS = XXX)",
                "**",
                "**     COL 04-07 = NA-PRODUCT-CLASS (605)",
                "**     COL 10-13 = FRD-DOC-CODE",
                "**     COL 16-16 = FRD-STATUS-CODE",
                "**     COL 19-23 = FRD-REEL-NO",
                "**     COL 26-26 = FRD-SIGNED-CYMD ('Y' TO USE CURRENT DATE OR 'N' TO INITIALIZE",
                "**     COL 29-88 = FNP-NOTEPAD",
                "**     COL 91-91 = GENERATE WARNING FOR DOCA BEING CREATED ('Y' OR 'N')",
                "**",
                "** ****  ****  *  *****  *  ************************************************************  *"
            ])
        
        # Add data rows (simplified for now)
        for row_idx in record_rows[:5]:  # Limit to first 5 rows for now
            output.append("** BBK   A1    R  00000  N")
        
        return output
    
    def _find_table_sheet(self, table_reference: str, excel_sheets: Dict[str, pd.DataFrame]) -> Optional[pd.DataFrame]:
        """Find the table sheet for a given table reference"""
        print(f"DEBUG: Looking for table reference '{table_reference}'")
        print(f"DEBUG: Available sheets: {list(excel_sheets.keys())}")
        
        table_name_variations = [
            f"table {table_reference}",
            f"table{table_reference}",
            f"table_{table_reference}",
            f"table-{table_reference}",
            f"table {table_reference.upper()}",
            f"table{table_reference.upper()}",
            f"table_{table_reference.upper()}",
            f"table-{table_reference.upper()}",
            table_reference,
            table_reference.upper()
        ]
        
        # Check exact matches first
        for variation in table_name_variations:
            if variation in excel_sheets:
                print(f"DEBUG: Found exact match: '{variation}'")
                return excel_sheets[variation]
        
        # Check for sheets that match the pattern %Table%X% where X is uppercase
        # Prioritize exact matches first, then more specific patterns
        table_ref_upper = table_reference.upper()
        table_ref_lower = table_reference.lower()
        
        for sheet_name, sheet_data in excel_sheets.items():
            sheet_lower = sheet_name.lower()
            
            # 1. Check for exact "Table X" pattern first (highest priority)
            if sheet_lower == f'table {table_ref_lower}':
                print(f"DEBUG: Found exact table sheet: '{sheet_name}' (exact match)")
                return sheet_data
            
            # 2. Check for "Table X" at the beginning of sheet name
            if sheet_lower.startswith(f'table {table_ref_lower}'):
                print(f"DEBUG: Found table sheet: '{sheet_name}' (starts with 'table {table_ref_lower}')")
                return sheet_data
            
            # 3. Check for "Table X" followed by space or end of word (to avoid false matches)
            import re
            pattern = rf'\btable\s+{table_ref_upper}\b'
            if re.search(pattern, sheet_name, re.IGNORECASE):
                print(f"DEBUG: Found table sheet: '{sheet_name}' (word boundary match)")
                return sheet_data
        
        print(f"DEBUG: No table sheet found for reference '{table_reference}'")
        return None
    
    def _generate_iras_table_reference_output(self, table_referenced_fields: List[Dict[str, Any]], excel_sheets: Dict[str, pd.DataFrame], available_sheets: List[str]) -> List[str]:
        """Generate IRAS mode table reference output with DSTRBTN SORC processing"""
        table_output = []
        
        # Add a header for table references
        table_output.append("")
        table_output.append("**** TABLE REFERENCE DATA (IRAS MODE) ****")
        table_output.append("")
        
        # Group DSTRBTN SORC references by begin field to avoid duplicates
        dstrbtn_references = {}
        other_references = []
        
        for field_mapping in table_referenced_fields:
            begin = field_mapping['begin']
            demo_field = field_mapping['demo_field']
            table_reference = field_mapping['table_reference']
            
            # Check if this is a DSTRBTN SORC table reference
            if table_reference and 'dstrbtn' in table_reference.lower() and 'sorc' in table_reference.lower():
                # Group by begin field to handle multiple references to same table
                if begin not in dstrbtn_references:
                    dstrbtn_references[begin] = {
                        'demo_field': demo_field,
                        'table_reference': table_reference
                    }
            else:
                # For other table references, just add the marker
                other_references.append(table_reference)
        
        # First, extract TABLE A data for IRAS mode
        table_a_output = self._extract_table_a_data_for_iras(excel_sheets)
        if table_a_output:
            table_output.extend(table_a_output)
            table_output.append("")  # Add spacing
        
        # Then process each unique DSTRBTN SORC reference
        dstrbtn_sheet = self._find_dstrbtn_sorc_sheet(excel_sheets)
        if dstrbtn_sheet is not None:
            for begin_field, ref_info in dstrbtn_references.items():
                print(f"DEBUG: Processing DSTRBTN SORC reference for begin field '{begin_field}'")
                dstrbtn_output = self._process_dstrbtn_sorc_data(begin_field, ref_info['demo_field'], dstrbtn_sheet)
                if dstrbtn_output:
                    table_output.extend(dstrbtn_output)
                    table_output.append("")  # Add spacing between different begin fields
        
        # Add other table reference markers (but skip "table a" since we already extracted the data)
        for table_ref in other_references:
            if table_ref and table_ref.lower() != 'a':  # Skip "table a" since we already processed it
                table_output.append(f"# table {table_ref}")
        
        return table_output
    
    def _extract_table_a_data_for_iras(self, excel_sheets: Dict[str, pd.DataFrame]) -> List[str]:
        """Extract TABLE A data for IRAS mode in the required format using dynamic column detection"""
        output = []
        
        try:
            # Find TABLE A sheet
            table_a_sheet = None
            for sheet_name, sheet_data in excel_sheets.items():
                sheet_name_lower = sheet_name.lower().replace(' ', '').replace('-', '')
                if 'table' in sheet_name_lower and 'a' in sheet_name_lower:
                    table_a_sheet = sheet_data
                    print(f"DEBUG: Found TABLE A sheet: '{sheet_name}'")
                    break
            
            if table_a_sheet is None:
                print("DEBUG: No TABLE A sheet found")
                return []
            
            # Add header
            output.append("************************************************************************")
            output.append("**** RECORD TYPE TYPE - IRA-TYPE AND FEE-SCHEDULE-NEW")
            output.append("****   ( EXPECTING MULTIPLE CARDS )")
            output.append("****")
            output.append("****   COL 006-006 = PER-RETI-PLAN-TYPE")
            output.append("****   COL 009-009 = PER-RETI-ACCT-TYPE")
            output.append("****   COL 012-013 = IRAS-IRA-TYPE")
            output.append("****   COL 016-018 = IRAS-FEE-SCHEDULE-NEW")
            output.append("****   COL 021-021 = WARNING REPORT")
            output.append("****")
            output.append("**** *  *  **  ***  *")
            
            print(f"DEBUG: TABLE A sheet has {len(table_a_sheet)} rows and {len(table_a_sheet.columns)} columns")
            
            # Use dynamic column detection
            column_mapping = self._detect_columns_dynamically(table_a_sheet, "table_a")
            print(f"DEBUG: TABLE A column mapping: {column_mapping}")
            
            # Find the data range: from Row 1 to last non-empty row
            data_start_row = 1  # Start from row 1 (0-indexed)
            last_data_row = 0
            
            for row_idx in range(data_start_row, len(table_a_sheet)):
                has_data = False
                for col_idx in range(len(table_a_sheet.columns)):
                    cell_value = str(table_a_sheet.iloc[row_idx, col_idx]).strip()
                    if cell_value and cell_value != 'nan':
                        has_data = True
                        break
                if has_data:
                    last_data_row = row_idx
                else:
                    break
            
            print(f"DEBUG: TABLE A data range: rows {data_start_row} to {last_data_row}")
            
            # Extract data using dynamic column mapping
            for row_idx in range(data_start_row, last_data_row + 1):
                # Extract data using detected column positions
                plan_type = self._get_cell_value(table_a_sheet, row_idx, column_mapping.get('plan_type', 0))
                acct_type = self._get_cell_value(table_a_sheet, row_idx, column_mapping.get('acct_type', 1))
                ira_type = self._get_cell_value(table_a_sheet, row_idx, column_mapping.get('ira_type', 2))
                fee_schedule = self._get_cell_value(table_a_sheet, row_idx, column_mapping.get('fee_schedule', 3))
                
                # Debug: Show what values are being extracted
                if row_idx <= 5:  # Only show first few rows to avoid spam
                    print(f"DEBUG: Row {row_idx} - Plan: '{plan_type}' (col {column_mapping.get('plan_type', 0)}), Acct: '{acct_type}' (col {column_mapping.get('acct_type', 1)}), IRA: '{ira_type}' (col {column_mapping.get('ira_type', 2)}), Fee: '{fee_schedule}' (col {column_mapping.get('fee_schedule', 3)})")
                
                # Format the output line: TBLA       IR  001
                # or: TBLA 1  1  IR  001
                if plan_type and acct_type:
                    formatted_line = f"TBLA {plan_type:<2} {acct_type:<2} {ira_type:<2} {fee_schedule}"
                elif acct_type:
                    formatted_line = f"TBLA     {acct_type:<2} {ira_type:<2} {fee_schedule}"
                else:
                    formatted_line = f"TBLA       {ira_type:<2} {fee_schedule}"
                
                output.append(formatted_line)
                print(f"DEBUG: TABLE A row {row_idx}: '{formatted_line}'")
            
            # Add footer
            output.append("****")
            
        except Exception as e:
            print(f"DEBUG: Error extracting TABLE A data: {e}")
            output.append(f"**** ERROR: TABLE A extraction failed - {str(e)} ****")
        
        return output
    
    def _detect_columns_dynamically(self, sheet: pd.DataFrame, table_type: str) -> Dict[str, int]:
        """Dynamic column detection using header-based approach with fallback strategy for IRAS mode"""
        column_mapping = {}
        
        print(f"DEBUG: Starting header-based column detection for {table_type}")
        
        # Step 1: Try header-based detection first
        column_mapping.update(self._find_columns_by_header(sheet, table_type))
        
        # Step 2: Check if header detection was successful
        required_columns = self._get_required_columns(table_type)
        header_detection_success = len(column_mapping) > 0
        
        # Count how many required columns were found via header detection
        found_via_header = sum(1 for col_type in required_columns if col_type in column_mapping)
        total_required = len(required_columns)
        
        print(f"DEBUG: Header detection found {found_via_header}/{total_required} required columns")
        
        # Step 3: If header detection failed completely or found very few columns, use fallback strategy
        if not header_detection_success or found_via_header < (total_required // 2):
            print(f"DEBUG: Header detection failed or insufficient matches. Using fallback strategy.")
            column_mapping = self._get_fallback_column_mapping(table_type)
        else:
            # Fill any missing columns with fallback positions
            fallback_positions = {
                'plan_type': 0,    # First column
                'acct_type': 1,    # Second column  
                'ira_type': 2,     # Third column
                'fee_schedule': 3  # Fourth column
            }
            
            for col_type in required_columns:
                if col_type not in column_mapping and col_type in fallback_positions:
                    column_mapping[col_type] = fallback_positions[col_type]
                    print(f"DEBUG: Filling missing {col_type} with fallback position: column {fallback_positions[col_type]}")
        
        print(f"DEBUG: Final column mapping for {table_type}: {column_mapping}")
        return column_mapping
    
    def _find_columns_by_header(self, sheet: pd.DataFrame, table_type: str) -> Dict[str, int]:
        """Find columns by matching header patterns"""
        column_mapping = {}
        
        print(f"DEBUG: Analyzing headers for {table_type}")
        
        if table_type == "table_a":
            header_patterns = {
                'plan_type': ["PER-RETI-PLAN-TYPE", "PLAN-TYPE", "RETI-PLAN", "PLAN", "PER-RETI", "RETI"],
                'acct_type': ["PER-RETI-ACCT-TYPE", "ACCT-TYPE", "RETI-ACCT", "ACCT", "ACCOUNT", "PER-RETI-ACCT"],
                'ira_type': ["IRAS-IRA-TYPE", "IRA-TYPE", "IRA", "IRAS"],
                'fee_schedule': ["IRAS-FEE-SCHEDULE-NEW", "FEE-SCHEDULE", "FEE-SCHEDULE-NEW", "FEE", "SCHEDULE"]
            }
        elif table_type == "dstrbtn_sorc":
            header_patterns = {
                'ira_type': ["IRAS-IRA-TYPE", "IRA-TYPE", "IRA"],
                'dist_source_1': ["DISTRIBUTION-SOURCE", "DIST-SOURCE", "DISTRIBUTION"],
                'dist_source_2': ["DISTRIBUTION-SOURCE", "DIST-SOURCE", "DISTRIBUTION"],
                'warning_report': ["WARNING-REPORT", "WARNING", "REPORT"]
            }
        else:
            return column_mapping
        
        # Print all headers for debugging
        print(f"DEBUG: Available headers in {table_type}:")
        for col_idx in range(len(sheet.columns)):
            header_value = str(sheet.iloc[0, col_idx]).strip()
            print(f"  Column {col_idx}: '{header_value}'")
        
        for col_idx in range(len(sheet.columns)):
            header_value = str(sheet.iloc[0, col_idx]).lower().strip()
            
            for col_type, patterns in header_patterns.items():
                if col_type not in column_mapping:  # Only take first match
                    for pattern in patterns:
                        if pattern.lower() in header_value:
                            column_mapping[col_type] = col_idx
                            print(f"DEBUG: ✓ Found {col_type} in column {col_idx} using header pattern '{pattern}' -> '{header_value}'")
                            break
        
        print(f"DEBUG: Header-based mapping result: {column_mapping}")
        return column_mapping
    
    def _get_required_columns(self, table_type: str) -> List[str]:
        """Get list of required columns for each table type"""
        if table_type == "table_a":
            return ['plan_type', 'acct_type', 'ira_type', 'fee_schedule']
        elif table_type == "dstrbtn_sorc":
            return ['ira_type', 'dist_source_1', 'dist_source_2', 'warning_report']
        return []
    
    def _get_fallback_column_mapping(self, table_type: str) -> Dict[str, int]:
        """Get fallback column mapping when header detection fails completely"""
        print(f"DEBUG: Using complete fallback strategy for {table_type}")
        
        if table_type == "table_a":
            return {
                'plan_type': 0,    # First column
                'acct_type': 1,    # Second column  
                'ira_type': 2,     # Third column
                'fee_schedule': 3  # Fourth column
            }
        elif table_type == "dstrbtn_sorc":
            return {
                'ira_type': 0,        # First column
                'dist_source_1': 1,   # Second column
                'dist_source_2': 2,   # Third column
                'warning_report': 3   # Fourth column
            }
        else:
            return {}
    
    def _detect_column_by_data(self, sheet: pd.DataFrame, col_type: str) -> int:
        """Detect column type by analyzing data patterns"""
        start_row = 1  # Skip header row
        
        for col_idx in range(len(sheet.columns)):
            sample_values = []
            for row_idx in range(start_row, min(start_row + 10, len(sheet))):
                value = str(sheet.iloc[row_idx, col_idx]).strip()
                if value and value != 'nan':
                    sample_values.append(value)
            
            if not sample_values:
                continue
            
            print(f"DEBUG: Analyzing column {col_idx} for {col_type} with values: {sample_values}")
            
            # Analyze patterns based on column type
            if col_type == 'ira_type' and self._is_ira_type_column(sample_values):
                print(f"DEBUG: Detected ira_type in column {col_idx} with values: {sample_values}")
                return col_idx
            elif col_type == 'dist_source_1' and self._is_distribution_source_column(sample_values):
                return col_idx
            elif col_type == 'dist_source_2' and self._is_distribution_source_column(sample_values):
                return col_idx
            elif col_type == 'warning_report' and self._is_warning_column(sample_values):
                return col_idx
            elif col_type == 'plan_type' and self._is_plan_type_column(sample_values):
                print(f"DEBUG: Detected plan_type in column {col_idx} with values: {sample_values}")
                return col_idx
            elif col_type == 'acct_type' and self._is_acct_type_column(sample_values):
                print(f"DEBUG: Detected acct_type in column {col_idx} with values: {sample_values}")
                return col_idx
            elif col_type == 'fee_schedule' and self._is_fee_schedule_column(sample_values):
                print(f"DEBUG: Detected fee_schedule in column {col_idx} with values: {sample_values}")
                return col_idx
        
        return -1
    
    def _is_ira_type_column(self, values: List[str]) -> bool:
        """Check if column contains IRA type codes (2-character codes)"""
        ira_patterns = ['BN', 'BR', 'BS', 'CK', 'CP', 'CS', 'ED', 'IR', 'IS', 'RB', 'RH', 'RO', 'SS', 'XI', 'ZH']
        return any(v in ira_patterns for v in values[:5]) or \
               all(len(v) == 2 and v.isalpha() for v in values[:3])
    
    def _is_distribution_source_column(self, values: List[str]) -> bool:
        """Check if column contains distribution source codes"""
        dist_patterns = ['DEC2', 'REG', 'PREM', 'EDU2', 'RHE2', 'RTH2', 'SMP2']
        return any(v in dist_patterns for v in values[:5])
    
    def _is_warning_column(self, values: List[str]) -> bool:
        """Check if column contains warning report values"""
        warning_patterns = ['Y', 'N', 'YES', 'NO']
        return any(v.upper() in warning_patterns for v in values[:5])
    
    def _is_plan_type_column(self, values: List[str]) -> bool:
        """Check if column contains plan type values"""
        plan_patterns = ['1', '2', 'A', 'H', 'J', 'L', 'R', 'S', 'T']
        
        # Check if values match plan type patterns
        matches = sum(1 for v in values[:5] if v in plan_patterns)
        
        # Plan type should have multiple matches to be confident
        return matches >= 2
    
    def _is_acct_type_column(self, values: List[str]) -> bool:
        """Check if column contains account type values"""
        acct_patterns = ['1', '2', '3', '4', '5', '6', '8', '9', 'A', 'B', 'C']
        
        # Check if values match account type patterns
        matches = sum(1 for v in values[:5] if v in acct_patterns)
        
        # Also check for single character values (typical for account types)
        single_char_count = sum(1 for v in values[:5] if len(str(v).strip()) == 1)
        
        # Account type should have either pattern matches or mostly single characters
        return matches >= 2 or single_char_count >= 3
    
    def _is_fee_schedule_column(self, values: List[str]) -> bool:
        """Check if column contains fee schedule values"""
        fee_patterns = ['001', '003']
        return any(v in fee_patterns for v in values[:5]) or \
               all(len(v) == 3 and v.isdigit() for v in values[:3])
    
    def _get_cell_value(self, sheet: pd.DataFrame, row_idx: int, col_idx: int) -> str:
        """Helper method to safely get cell value"""
        if col_idx < 0 or col_idx >= len(sheet.columns) or row_idx >= len(sheet):
            return ''
        
        value = str(sheet.iloc[row_idx, col_idx]).strip()
        if value == 'nan':
            return ''
        
        # Convert literal "space" to actual space character
        if value.lower() == "space":
            value = " "
        
        return value
    
    def _find_dstrbtn_sorc_sheet(self, excel_sheets: Dict[str, pd.DataFrame]) -> pd.DataFrame:
        """Find the DSTRBTN SORC sheet in the Excel file"""
        for sheet_name, sheet_data in excel_sheets.items():
            sheet_name_lower = sheet_name.lower().replace(' ', '').replace('-', '')
            if 'dstrbtn' in sheet_name_lower and 'sorc' in sheet_name_lower:
                print(f"DEBUG: Found DSTRBTN SORC sheet: '{sheet_name}'")
                return sheet_data
        
        print("DEBUG: No DSTRBTN SORC sheet found")
        return None
    
    def _process_dstrbtn_sorc_data(self, begin: str, demo_field: str, dstrbtn_sheet: pd.DataFrame) -> List[str]:
        """Process DSTRBTN SORC sheet data and generate formatted output"""
        output = []
        
        try:
            # Search for the begin pattern in the DSTRBTN SORC sheet
            begin_position = self._find_begin_row(begin, dstrbtn_sheet)
            if begin_position is None:
                print(f"DEBUG: Begin pattern '{begin}' not found in DSTRBTN SORC sheet")
                return []
            
            row_idx, col_idx = begin_position
            print(f"DEBUG: Found begin pattern '{begin}' in DSTRBTN SORC sheet at row {row_idx}, col {col_idx}")
            
            # Extract the table data starting from the begin position
            # Pass the begin field to use as prefix (e.g., 297 -> 297a)
            table_data = self._extract_dstrbtn_sorc_table_data(dstrbtn_sheet, row_idx, col_idx, begin)
            if table_data:
                output.extend(table_data)
            
        except Exception as e:
            print(f"DEBUG: Error processing DSTRBTN SORC data: {e}")
            output.append(f"**** ERROR: DSTRBTN SORC processing failed - {str(e)} ****")
        
        return output
    
    def _extract_dstrbtn_sorc_table_data(self, dstrbtn_sheet: pd.DataFrame, start_row: int, start_col: int, begin_field: str) -> List[str]:
        """Extract table data from DSTRBTN SORC sheet in the required format"""
        output = []
        
        try:
            # Add header
            output.append("****")
            
            # Create the prefix from the begin field (e.g., 297 -> 297A)
            prefix = f"{begin_field}A"
            print(f"DEBUG: Using prefix '{prefix}' from begin field '{begin_field}'")
            
            print(f"DEBUG: Extracting DSTRBTN SORC data from row {start_row}, col {start_col}")
            print(f"DEBUG: Sheet has {len(dstrbtn_sheet)} rows and {len(dstrbtn_sheet.columns)} columns")
            
            # First, let's examine the entire sheet structure to understand the layout
            print("DEBUG: Examining DSTRBTN SORC sheet structure:")
            for row_idx in range(min(10, len(dstrbtn_sheet))):  # Show first 10 rows
                row_data = []
                for col_idx in range(min(8, len(dstrbtn_sheet.columns))):  # Show first 8 columns
                    cell_value = str(dstrbtn_sheet.iloc[row_idx, col_idx]).strip()
                    if cell_value == 'nan':
                        cell_value = ''
                    row_data.append(f"'{cell_value}'")
                print(f"DEBUG: Row {row_idx}: [{', '.join(row_data)}]")
            
            # Based on the image you provided, the structure should be:
            # Row 0: Headers
            # Row 1+: Data rows with IRA type codes in column 0, distribution sources in columns 1 and 2
            
            # Let's look for the actual data rows - start from row 1 and look for non-empty cells
            data_found = False
            for row_idx in range(1, len(dstrbtn_sheet)):
                print(f"DEBUG: Checking row {row_idx}")
                
                # Check all columns in this row to see what's there
                row_contents = []
                for col_idx in range(min(5, len(dstrbtn_sheet.columns))):
                    cell_value = str(dstrbtn_sheet.iloc[row_idx, col_idx]).strip()
                    if cell_value == 'nan':
                        cell_value = ''
                    row_contents.append(f"Col{col_idx}:'{cell_value}'")
                print(f"DEBUG: Row {row_idx} contents: {', '.join(row_contents)}")
                
                # Check if this row has any meaningful data
                has_data = False
                for col_idx in range(len(dstrbtn_sheet.columns)):
                    cell_value = str(dstrbtn_sheet.iloc[row_idx, col_idx]).strip()
                    if cell_value and cell_value != 'nan' and len(cell_value) > 0:
                        has_data = True
                        break
                
                if not has_data:
                    print(f"DEBUG: Row {row_idx} is empty, stopping")
                    break
                
                # Extract the data as-is from the DSTRBTN SORC sheet
                # Based on the debug output: Column 1=IRA Type, Column 2=Dist Source 1, Column 3=Dist Source 2, Column 4=Warning Report
                
                ira_type = str(dstrbtn_sheet.iloc[row_idx, 1]).strip() if pd.notna(dstrbtn_sheet.iloc[row_idx, 1]) else ''
                dist_source_1 = str(dstrbtn_sheet.iloc[row_idx, 2]).strip() if pd.notna(dstrbtn_sheet.iloc[row_idx, 2]) else ''
                dist_source_2 = str(dstrbtn_sheet.iloc[row_idx, 3]).strip() if pd.notna(dstrbtn_sheet.iloc[row_idx, 3]) else ''
                warning_report = str(dstrbtn_sheet.iloc[row_idx, 4]).strip() if pd.notna(dstrbtn_sheet.iloc[row_idx, 4]) else ''
                
                # Convert literal "space" to actual space character
                if ira_type.strip().lower() == "space":
                    ira_type = " "
                if dist_source_1.strip().lower() == "space":
                    dist_source_1 = " "
                if dist_source_2.strip().lower() == "space":
                    dist_source_2 = " "
                if warning_report.strip().lower() == "space":
                    warning_report = " "
                
                print(f"DEBUG: Row {row_idx} - IRA_TYPE: '{ira_type}', DIST1: '{dist_source_1}', DIST2: '{dist_source_2}', WARNING: '{warning_report}'")
                
                # Clean up the values
                if ira_type == 'nan':
                    ira_type = ''
                if dist_source_1 == 'nan':
                    dist_source_1 = ''
                if dist_source_2 == 'nan':
                    dist_source_2 = ''
                if warning_report == 'nan':
                    warning_report = ''  # Keep as blank if nan
                
                # Skip rows with empty IRA type
                if not ira_type:
                    print(f"DEBUG: Skipping row {row_idx} - empty IRA type")
                    continue
                
                # Format the output line: 297a **   ****  ****  **
                # 297a BN   DEC2  DEC2  Y
                formatted_line = f"{prefix} {ira_type:<3} {dist_source_1:<4} {dist_source_2:<4} {warning_report}"
                output.append(formatted_line)
                print(f"DEBUG: Added line: '{formatted_line}'")
                data_found = True
            
            # Add footer
            output.append("****")
            
            print(f"DEBUG: Total lines extracted: {len(output)}")
            if not data_found:
                print("DEBUG: No data found in DSTRBTN SORC sheet")
            
        except Exception as e:
            print(f"DEBUG: Error extracting DSTRBTN SORC table data: {e}")
            output.append(f"**** ERROR: Table data extraction failed - {str(e)} ****")
        
        return output
    
    def _process_single_table_reference(self, begin: str, demo_field: str, table_reference: str, table_sheet: pd.DataFrame) -> List[str]:
        """Process a single table reference and generate formatted output"""
        try:
            print(f"DEBUG: Processing table reference for field '{demo_field}' with begin '{begin}' in table '{table_reference}'")
            
            # Search for the begin pattern in the table sheet
            begin_position = self._find_begin_row(begin, table_sheet)
            if begin_position is None:
                print(f"DEBUG: No begin position found for '{begin}' in table '{table_reference}'")
                return [f"**** ALERT: Field {begin} has table reference but 'beg {begin}' not found in table sheet ****"]
            
            begin_row, begin_col = begin_position
            print(f"DEBUG: Found begin position at row {begin_row}, col {begin_col}")
            
            # Extract range data from rows below the begin row, starting from column M-1
            range_data = self._extract_range_data(table_sheet, begin_row, begin_col)
            if not range_data:
                print(f"DEBUG: No range data extracted for field '{demo_field}'")
                return [f"**** ALERT: No range data found for field {begin} in table sheet ****"]
            
            print(f"DEBUG: Extracted {len(range_data)} range data items for field '{demo_field}'")
            
            # Generate formatted output
            return self._format_table_output(begin, demo_field, table_reference, range_data)
            
        except Exception as e:
            print(f"DEBUG: Error processing table reference for field {demo_field}: {str(e)}")
            import traceback
            traceback.print_exc()
            return [f"**** ALERT: Error processing table reference for field {demo_field} ****"]
    
    def _find_begin_row(self, begin: str, table_sheet: pd.DataFrame) -> Optional[tuple]:
        """Find the row and column containing the begin pattern"""
        # Remove leading zeros from begin number for searching
        begin_trimmed = str(int(begin))  # This removes leading zeros
        
        # Try multiple patterns to find the begin reference
        patterns_to_try = [
            f"(Beg {begin_trimmed})",  # Capital B - as seen in Excel
            f"(beg {begin_trimmed})",  # Lowercase b - fallback
            f"Beg {begin_trimmed}",    # Capital B without parentheses
            f"beg {begin_trimmed}",    # Lowercase b without parentheses
            f"begin {begin_trimmed}",
            f"position {begin_trimmed}",
            f"pos {begin_trimmed}",
            f"({begin_trimmed})"
        ]
        
        print(f"DEBUG: Searching for begin patterns in table sheet with {len(table_sheet)} rows and {len(table_sheet.columns)} columns")
        print(f"DEBUG: Looking for begin number: '{begin_trimmed}'")
        
        # Search through all cells in the table sheet
        for row_idx in range(len(table_sheet)):
            for col_idx in range(len(table_sheet.columns)):
                cell_value = str(table_sheet.iloc[row_idx, col_idx])
                cell_lower = cell_value.lower()
                
                # Debug: Print first few rows to see what's in the table sheet
                if row_idx < 10 and col_idx < 20:  # Show more rows and columns
                    print(f"DEBUG: Table sheet row {row_idx}, col {col_idx}: '{cell_value}'")
                
                # Debug: Look for any cell containing the begin number
                if begin_trimmed in cell_value:
                    print(f"DEBUG: Found cell containing '{begin_trimmed}': row {row_idx}, col {col_idx}, value: '{cell_value}'")
                
                # Debug: Look for any cell containing 'beg' or 'Beg'
                if 'beg' in cell_lower or 'Beg' in cell_value:
                    print(f"DEBUG: Found cell containing 'beg': row {row_idx}, col {col_idx}, value: '{cell_value}'")
                
                # Try each pattern, but only match if it contains 'beg' (case insensitive)
                for pattern in patterns_to_try:
                    if pattern.lower() in cell_lower and ('beg' in cell_lower or 'Beg' in cell_value):
                        print(f"DEBUG: Found begin pattern '{pattern}' in row {row_idx}, column {col_idx}")
                        print(f"DEBUG: Cell content: '{cell_value}'")
                        
                        # For header rows (row 0), we need to find the actual data section
                        if row_idx == 0:
                            # This is a header row, find the corresponding data section
                            # The data is typically in the column where the begin pattern was found
                            # Look for the next section with actual data in the same column
                            for data_row in range(1, len(table_sheet)):
                                # Check if this row has data in the same column structure
                                if col_idx < len(table_sheet.columns):
                                    data_cell = str(table_sheet.iloc[data_row, col_idx])
                                    if data_cell.strip() and data_cell.strip() not in ['nan', 'none', '']:
                                        # Found the start of data section
                                        print(f"DEBUG: Found data section for {begin_trimmed} starting at row {data_row}, column {col_idx}")
                                        return (data_row, col_idx)
                        
                        return (row_idx, col_idx)
        
        print(f"DEBUG: Begin patterns for '{begin_trimmed}' not found in table sheet")
        print(f"DEBUG: Tried patterns: {patterns_to_try}")
        return None
    
    def _extract_range_data(self, table_sheet: pd.DataFrame, begin_row: int, begin_col: int) -> List[Dict[str, str]]:
        """Extract range data from rows below the begin row, with fallback column logic"""
        range_data = []
        
        # Start from the row after the begin row, and look at column M-1
        target_col = begin_col - 1
        if target_col < 0:
            target_col = 0
        
        print(f"DEBUG: Starting to extract range data from row {begin_row + 1}, column {target_col}")
        
        # Check if the first target column is empty, if so try previous column
        start_row = begin_row + 1
        if start_row < len(table_sheet):
            first_cell_value = str(table_sheet.iloc[start_row, target_col]).strip()
            if not first_cell_value or first_cell_value.lower() in ['nan', 'none', '']:
                print(f"DEBUG: Column {target_col} is empty, trying previous column {target_col - 1}")
                target_col = target_col - 1
                if target_col < 0:
                    print(f"DEBUG: ERROR: No data found in expected columns")
                    return []
                
                # Check if the fallback column is also empty
                fallback_cell_value = str(table_sheet.iloc[start_row, target_col]).strip()
                if not fallback_cell_value or fallback_cell_value.lower() in ['nan', 'none', '']:
                    print(f"DEBUG: ERROR: Fallback column {target_col} is also empty")
                    return []
                
                print(f"DEBUG: Using fallback column {target_col} with value: '{fallback_cell_value}'")
        
        for row_idx in range(begin_row + 1, len(table_sheet)):
                # Get the specific cell at row N+1, column M-1
                if target_col < len(table_sheet.columns):
                    cell_value = table_sheet.iloc[row_idx, target_col]
                    range_text = str(cell_value) if pd.notna(cell_value) else ""
                    
                    # Convert literal "space" to actual space character
                    if range_text.strip().lower() == "space":
                        range_text = " "
                    
                    print(f"DEBUG: Row {row_idx}, Col {target_col} value: '{range_text}'")
                    
                    # Check if this row is completely empty (both range_text and code are empty)
                    code = ""
                    warning_report = ""
                    if target_col + 1 < len(table_sheet.columns):
                        code_cell = table_sheet.iloc[row_idx, target_col + 1]
                        code = str(code_cell) if pd.notna(code_cell) else ""
                        # Convert literal "space" to actual space character
                        if code.strip().lower() == "space":
                            code = " "
                    if target_col + 2 < len(table_sheet.columns):
                        warning_cell = table_sheet.iloc[row_idx, target_col + 2]
                        warning_report = str(warning_cell) if pd.notna(warning_cell) else ""
                        # Convert literal "space" to actual space character
                        if warning_report.strip().lower() == "space":
                            warning_report = " "
                
                # If both range_text and code are empty or contain only whitespace, stop reading
                if not range_text.strip() and not code.strip():
                    print(f"DEBUG: Found empty row at {row_idx}, stopping range data extraction.")
                    break
                
                # Stop reading if we encounter descriptive text that indicates end of data section
                if range_text.strip() and any(desc_text in range_text for desc_text in [
                    'Pershing Description', 'L = Low', 'M = Medium', 'H = High', 
                    'Description', 'Legend', 'Notes', 'Total', 'Count'
                ]):
                    print(f"DEBUG: Found descriptive text '{range_text}' at {row_idx}, stopping range data extraction.")
                    break
                
                if range_text.strip() and range_text.strip() not in ['nan', 'none']:  # If there's actual content
                    # Skip rows that don't look like valid range data
                    if not self._is_valid_range_data(range_text, code):
                        print(f"DEBUG: Skipping invalid range data: '{range_text}' with code '{code}'")
                        continue
                    
                    # Convert range to numeric format
                    min_val, max_val = self._convert_range_to_numeric(range_text)
                    
                    range_info = {
                        'range_text': range_text,
                        'min_val': min_val,
                        'max_val': max_val,
                        'code': code,
                        'warning_report': warning_report
                    }
                    
                    range_data.append(range_info)
                    print(f"DEBUG: Added range info: {range_info}")
        
        print(f"DEBUG: Total range data extracted: {len(range_data)}")
        return range_data
    
    def _is_valid_range_data(self, range_text: str, code: str) -> bool:
        """Check if the range data is valid"""
        range_text = range_text.strip()
        code = code.strip()
        
        # Skip empty or invalid entries
        if not range_text or range_text.lower() in ['nan', 'none', '']:
            return False
        
        # Skip entries that look like field names or other non-range data
        if any(keyword in range_text.upper() for keyword in ['PLANNED', 'INVEST', 'AMT', 'TOTAL', 'FIELD']):
            return False
        
        # Valid range patterns - both range-based and non-range-based data
        valid_patterns = [
            'zero' in range_text.lower(),
            'blank' in range_text.lower(),
            'default' in range_text.lower(),
            '$' in range_text,
            '99,99,99' in range_text,
            '10,00,00' in range_text,
            # Year/date patterns
            '> 0 & <=' in range_text,  # "> 0 & <= 2026"
            ('-' in range_text and range_text.replace('-', '').replace(' ', '').isdigit()),  # "2027-2029"
            '>=' in range_text,  # ">= 2035"
            code in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'N', ' '],
            # For non-range data, accept various formats
            len(range_text) <= 10 and range_text.isalpha(),  # Letters like HITB, LWTB, MDTB, TPTB
            range_text.isdigit() and len(range_text) <= 3,   # Numbers like 32, 21, 37
            range_text == '#',
            range_text == '**',  # Double asterisks
            range_text == ' ',
            # Accept any non-empty text that's not obviously invalid
            len(range_text) > 0 and not any(char in range_text for char in ['(', ')', ':', ';'])
        ]
        
        return any(valid_patterns)
    
    def _parse_range_row(self, row_values: List[str]) -> Optional[Dict[str, str]]:
        """Parse a single row to extract range information"""
        if len(row_values) < 2:
            return None
        
        # Look for range patterns in different columns
        # The range text could be in column 0, 1, or 2
        range_text = ""
        code = ""
        
        # Try to find range text (contains $, Zero, Blank, Default, etc.)
        for i, value in enumerate(row_values):
            value_str = str(value).strip()
            if ('$' in value_str or 'zero' in value_str.lower() or 'blank' in value_str.lower() or 
                'default' in value_str.lower() or '99,99,99' in value_str or '10,00,00' in value_str):
                range_text = value_str
                # Code is usually in the next column
                if i + 1 < len(row_values):
                    code = str(row_values[i + 1]).strip()
                break
        
        # If no range text found, try to find code patterns (A, B, C, etc.)
        if not range_text:
            for i, value in enumerate(row_values):
                value_str = str(value).strip()
                if value_str in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'N']:
                    code = value_str
                    # Range text might be in previous column
                    if i > 0:
                        range_text = str(row_values[i - 1]).strip()
                    break
        
        # Skip empty or invalid rows
        if not range_text or range_text.lower() in ['nan', 'none', '']:
            return None
        
        # Parse the range
        min_val, max_val = self._convert_range_to_numeric(range_text)
        
        return {
            'range_text': range_text,
            'min_val': min_val,
            'max_val': max_val,
            'code': code
        }
    
    def _convert_range_to_numeric(self, range_text: str) -> tuple:
        """Convert range text to 15-digit numeric format or return as-is for non-range data"""
        range_text = range_text.strip()
        
        # Handle special cases
        if 'zero' in range_text.lower() or 'blank' in range_text.lower() or 'default' in range_text.lower():
            return "000000000000000", "000000000000000"
        
        if '99,99,99,99,99,99,999' in range_text:
            return "000000000000000", "000000000000000"
        
        if '10,00,00,00,00,00,000' in range_text:
            return "000000000000000", "000000000000000"
        
        # Handle the special case where we have just numbers like "100000000000000"
        if range_text.isdigit() and len(range_text) >= 10:
            try:
                val = int(range_text)
                if val == 100000000000000:
                    return "100000000000000", "100000000000000"
            except ValueError:
                pass
        
        # Handle ranges like "$25,000-$49,999"
        if '-' in range_text and '$' in range_text:
            parts = range_text.split('-')
            if len(parts) == 2:
                min_part = parts[0].replace('$', '').replace(',', '').strip()
                max_part = parts[1].replace('$', '').replace(',', '').strip()
                
                try:
                    min_val = int(min_part)
                    max_val = int(max_part)
                    return f"{min_val:015d}", f"{max_val:015d}"
                except ValueError:
                    pass
        
        # Handle ranges like "1-24,999", "25,000-49,999", etc. (without $ symbol)
        if '-' in range_text and ',' in range_text and '$' not in range_text:
            parts = range_text.split('-')
            if len(parts) == 2:
                min_part = parts[0].replace(',', '').strip()
                max_part = parts[1].replace(',', '').strip()
                
                try:
                    min_val = int(min_part)
                    max_val = int(max_part)
                    return f"{min_val:015d}", f"{max_val:015d}"
                except ValueError:
                    pass
        
        # Handle single values like "$1,000,000+"
        if '$' in range_text and '+' in range_text:
            value_part = range_text.replace('$', '').replace(',', '').replace('+', '').strip()
            try:
                min_val = int(value_part)
                max_val = 99999999999999  # 14 digits for "+" ranges (will be padded to 15 with leading zero)
                return f"{min_val:015d}", f"{max_val:015d}"
            except ValueError:
                pass
        
        # Handle patterns like "=>1,000,000" (greater than or equal to)
        if '=>' in range_text and ',' in range_text:
            value_part = range_text.replace('=>', '').replace(',', '').strip()
            try:
                min_val = int(value_part)
                max_val = 99999999999999  # 14 digits for "=>" ranges (will be padded to 15 with leading zero)
                return f"{min_val:015d}", f"{max_val:015d}"
            except ValueError:
                pass
        
        # Handle descriptive range patterns like "L - (long over 10 years)", "I - (intermediate 6-10 yrs)", etc.
        # Extract the first character and convert to numeric based on time ranges
        if ' - (' in range_text and 'years' in range_text.lower():
            first_char = range_text.strip()[0].upper()
            
            # Map time horizon codes to numeric ranges
            if first_char == 'L':  # Long over 10 years
                return "000000000000011", "000000000000999"  # 11+ years
            elif first_char == 'I':  # Intermediate 6-10 yrs
                return "000000000000006", "000000000000010"  # 6-10 years
            elif first_char == 'S':  # Short 0-5 yrs
                return "000000000000000", "000000000000005"  # 0-5 years
            else:
                # For other single character codes, use the character as-is
                return first_char, first_char
        
        # Handle single character codes like "T", "V", "R", "N", "C", etc.
        if len(range_text.strip()) == 1 and range_text.strip().isalpha():
            char = range_text.strip().upper()
            return char, char
        
        # Handle codes with descriptions like "N - (not available)", "C - (combo)"
        if ' - (' in range_text:
            first_char = range_text.strip()[0].upper()
            return first_char, first_char
        
        # Handle single values like "1<$25,000"
        if '<' in range_text and '$' in range_text:
            value_part = range_text.split('$')[1].replace(',', '').strip()
            try:
                max_val = int(value_part) - 1
                min_val = 1
                return f"{min_val:015d}", f"{max_val:015d}"
            except ValueError:
                pass
        
        # Handle year/date patterns
        # Pattern: "> 0 & <= 2026" -> 00000001 to 20261231
        if '> 0 & <=' in range_text:
            try:
                year_part = range_text.split('<=')[1].strip()
                year = int(year_part)
                return "00000001", f"{year}1231"
            except (ValueError, IndexError):
                pass
        
        # Pattern: "2027-2029" -> 20270101 to 20291231
        if '-' in range_text and range_text.replace('-', '').replace(' ', '').isdigit():
            try:
                parts = range_text.split('-')
                if len(parts) == 2:
                    start_year = int(parts[0].strip())
                    end_year = int(parts[1].strip())
                    return f"{start_year}0101", f"{end_year}1231"
            except ValueError:
                pass
        
        # Pattern: ">= 2035" -> 20350101 to 99999999
        if '>=' in range_text:
            try:
                year_part = range_text.split('>=')[1].strip()
                year = int(year_part)
                return f"{year}0101", "99999999"
            except (ValueError, IndexError):
                pass
        
        # For non-range data (letters, codes, numbers), return as-is
        if (len(range_text) <= 10 and range_text.isalpha()) or \
           (range_text.isdigit() and len(range_text) <= 3) or \
           range_text in ['#', '**', ' '] or \
           (len(range_text) > 0 and not any(char in range_text for char in ['$', '-', '+', '<', '(', ')', ':', ';'])):
            return range_text, range_text  # Return the same value for both min and max
        
        # Default fallback
        return "000000000000000", "000000000000000"
    
    def _format_table_output(self, begin: str, demo_field: str, table_reference: str, range_data: List[Dict[str, str]]) -> List[str]:
        """Format the table output with proper structure"""
        output = []
        
        # Remove leading zeros from begin number for display
        begin_trimmed = str(int(begin))
        
        # Check if this is range-based data or non-range data
        is_range_data = any('$' in info['range_text'] or info['min_val'].isdigit() for info in range_data)
        
        # Special handling for DEMO-S-NET-INCOME
        if demo_field.upper() == "DEMO-S-NET-INCOME":
            output.append("************************************************************************")
            output.append("****")
            output.append(f"**** DEMO-S-NET-INCOME {begin_trimmed}A (TABLE A #1-TOP ON MAPPING)")
            output.append("****")
            output.append("****   COL 007-021 = PER-MAIN1-AWAY-ASST-VAL-AM(1)-(8) PLUS PER-MAIN1-AWYASST-VAL-AM-EXT(1)-(3) (LOW RANGE)")
            output.append("****   COL 024-038 = PER-MAIN1-AWAY-ASST-VAL-AM(1)-(8) PLUS PER-MAIN1-AWYASST-VAL-AM-EXT(1)-(3) (HIGH RANGE)")
            output.append("****   COL 041-041 = DEMO-S-NET-INCOME")
            output.append("****   COL 044-044 = WARNING REPORT (Y/N)")
            output.append("****")
            output.append("****  ***************  ***************  *  *")
            
            # Add range data with proper spacing based on column specifications
            for range_info in range_data:
                min_val = range_info['min_val']
                max_val = range_info['max_val']
                code = range_info['code']
                warning_report = range_info.get('warning_report', '')
                
                # Format with proper spacing: COL 007-021 (15 chars), COL 024-038 (15 chars), COL 041-041 (1 char), COL 044-044 (1 char)
                # Position 1-6: begin + table reference (e.g., "288A  ")
                # Position 7-21: low range (15 characters)
                # Position 22-23: spaces
                # Position 24-38: high range (15 characters)  
                # Position 39-40: spaces
                # Position 41: demo net income (1 character)
                # Position 42-43: spaces
                # Position 44: warning report (1 character)
                
                line = f"{begin_trimmed}A  {min_val:<15}  {max_val:<15}  {code}  {warning_report}"
                output.append(line)
            
            return output
        
        # Special handling for DEMO-NET-INCOME (without S)
        if demo_field.upper() == "DEMO-NET-INCOME":
            output.append("************************************************************************")
            output.append("****")
            output.append(f"**** DEMO-NET-INCOME {begin_trimmed}A (TABLE A #1-TOP ON MAPPING)")
            output.append("****")
            output.append("****   COL 007-021 = PER-MAIN1-AWAY-ASST-VAL-AM(1)-(8) PLUS PER-MAIN1-AWYASST-VAL-AM-EXT(1)-(3) (LOW RANGE)")
            output.append("****   COL 024-038 = PER-MAIN1-AWAY-ASST-VAL-AM(1)-(8) PLUS PER-MAIN1-AWYASST-VAL-AM-EXT(1)-(3) (HIGH RANGE)")
            output.append("****   COL 041-041 = DEMO-NET-INCOME")
            output.append("****   COL 044-044 = WARNING REPORT (Y/N)")
            output.append("****")
            output.append("****  ***************  ***************  *  *")
            
            # Add range data with proper spacing based on column specifications
            for range_info in range_data:
                min_val = range_info['min_val']
                max_val = range_info['max_val']
                code = range_info['code']
                warning_report = range_info.get('warning_report', '')
                
                # Format with proper spacing: COL 007-021 (15 chars), COL 024-038 (15 chars), COL 041-041 (1 char), COL 044-044 (1 char)
                # Position 1-6: begin + table reference (e.g., "288A  ")
                # Position 7-21: low range (15 characters)
                # Position 22-23: spaces
                # Position 24-38: high range (15 characters)  
                # Position 39-40: spaces
                # Position 41: demo net income (1 character)
                # Position 42-43: spaces
                # Position 44: warning report (1 character)
                
                line = f"{begin_trimmed}A  {min_val:<15}  {max_val:<15}  {code}  {warning_report}"
                output.append(line)
            
            return output
        
        # Special handling for DEMO-INV-TIME-HORIZON type items
        if "DEMO-INV-TIME-HORIZON" in demo_field.upper():
            output.append("************************************************************************")
            output.append("****")
            output.append(f"**** {demo_field.upper()}  {begin_trimmed} (TABLE D)")
            output.append("****")
            output.append("****   COL 007-014 = PER-MAIN1-INOB-TM-HORIZ-DT (BEGIN RANGE)")
            output.append("****   COL 017-024 = PER-MAIN1-INOB-TM-HORIZ-DT (END RANGE)")
            output.append(f"****   COL 027-027 = {demo_field.upper()}")
            output.append("****   COL 030-030 = WARNING REPORT(Y/N)")
            output.append("****")
            output.append("****  ********  ********  *  *")
            
            # Add range data with proper spacing based on column specifications
            for range_info in range_data:
                min_val = range_info['min_val']
                max_val = range_info['max_val']
                code = range_info['code']
                warning_report = range_info.get('warning_report', '')
                
                # Format with proper spacing: COL 007-014 (8 chars), COL 017-024 (8 chars), COL 027-027 (1 char), COL 030-030 (1 char)
                # Position 1-6: begin + table reference (e.g., "1078  ")
                # Position 7-14: begin range (8 characters)
                # Position 15-16: spaces
                # Position 17-24: end range (8 characters)
                # Position 25-26: spaces
                # Position 27: demo inv time horizon (1 character)
                # Position 28-29: spaces
                # Position 30: warning report (1 character)
                
                line = f"{begin_trimmed}  {min_val:<8}  {max_val:<8}  {code}  {warning_report}"
                output.append(line)
            
            return output
        
        # Special handling for DEMO-MARITAL-STATUS type items
        if "DEMO-MARITAL-STATUS" in demo_field.upper():
            output.append("************************************************************************")
            output.append("****")
            output.append(f"**** RECORD TYPE {begin_trimmed}A ({demo_field.upper().replace('-', ' ')})")
            output.append("****")
            output.append("****   COL 007-007 = PER-PRIM1-MARITAL-STATUS")
            output.append(f"****   COL 010-010 = {demo_field.upper()}        ({begin_trimmed})")
            output.append("****")
            output.append("****  *  *")
            
            # For DEMO-MARITAL-STATUS, extract data as-is (not as numeric ranges)
            # Use the table reference from the original table_reference parameter
            table_ref = table_reference.upper() if table_reference else "A"
            
            # Add range data with proper spacing based on column specifications
            for range_info in range_data:
                # For DEMO-MARITAL-STATUS, use the original range_text and code as-is
                range_text = range_info.get('range_text', '').strip()
                code = range_info.get('code', '').strip()
                
                # Handle special case: # should be treated as default (empty spaces) for DEMO-MARITAL-STATUS
                if range_text == '#':
                    col1_value = ''  # Empty for default
                else:
                    col1_value = range_text.strip()  # Extra strip to remove any remaining spaces
                
                if code == '#':
                    col2_value = ''  # Empty for default
                else:
                    col2_value = code.strip()  # Extra strip to remove any remaining spaces
                
                # Format with proper spacing: COL 007-007 (1 char), COL 010-010 (1 char)
                # Position 1-6: begin + table reference (e.g., "301C  ")
                # Position 7: PER-PRIM1-MARITAL-STATUS (1 character)
                # Position 8-9: spaces
                # Position 10: DEMO-MARITAL-STATUS (1 character)
                
                # Only add the 2nd column if it has a value, otherwise just the 1st column
                if col2_value:
                    line = f"{begin_trimmed}{table_ref}  {col1_value}  {col2_value}"
                else:
                    line = f"{begin_trimmed}{table_ref}  {col1_value}"
                output.append(line)
            
            return output
        
        # Special handling for DEMO-TAX-BRACKET type items
        if "DEMO-TAX-BRACKET" in demo_field.upper():
            output.append("************************************************************************")
            output.append("****")
            output.append(f"**** RECORD TYPE {begin_trimmed}A  ({demo_field.upper().replace('-', ' ')})")
            output.append("****")
            output.append("****   COL 007-010 = PER-PRIM1-TAX-STATUS-CODE")
            output.append(f"****   COL 013-014 = {demo_field.upper()}           ({begin_trimmed})")
            output.append("****   COL 017-017 = WARNING REPORT (N/Y)")
            output.append("****")
            output.append("****  ****  **  *")
            
            # For DEMO-TAX-BRACKET, extract data as-is (not as numeric ranges)
            # Use the table reference from the original table_reference parameter
            table_ref = table_reference.upper() if table_reference else "A"
            
            # Add range data with proper spacing based on column specifications
            for range_info in range_data:
                # For DEMO-TAX-BRACKET, use the original range_text and code as-is
                range_text = range_info.get('range_text', '').strip()
                code = range_info.get('code', '').strip()
                warning_report = range_info.get('warning_report', '').strip()
                
                # Handle special case: "space" should be treated as actual space for DEMO-TAX-BRACKET
                if range_text.lower() == 'space':
                    col1_value = '    '  # 4 spaces for "space"
                else:
                    # Limit PER-PRIM1-TAX-STATUS-CODE to 4 characters
                    col1_value = range_text.strip()[:4].ljust(4)  # Cut to 4 chars and pad with spaces
                
                if code.lower() == 'space':
                    col2_value = '  '  # 2 spaces for "space"
                else:
                    # Limit DEMO-TAX-BRACKET to 2 characters
                    col2_value = code.strip()[:2].ljust(2)  # Cut to 2 chars and pad with spaces
                
                # Handle warning report: only show Y/N if there's a value, otherwise space
                if warning_report and warning_report.upper() in ['Y', 'N']:
                    warning_value = warning_report.upper()
                else:
                    warning_value = ' '  # Space if no warning report or invalid value
                
                # Format with proper spacing: COL 007-010 (4 chars), COL 013-014 (2 chars), COL 017-017 (1 char)
                # Position 1-6: begin + table reference (e.g., "305C  ")
                # Position 7-10: PER-PRIM1-TAX-STATUS-CODE (4 characters)
                # Position 11-12: spaces
                # Position 13-14: DEMO-TAX-BRACKET (2 characters)
                # Position 15-16: spaces
                # Position 17: WARNING REPORT (1 character)
                
                line = f"{begin_trimmed}{table_ref}  {col1_value}  {col2_value}  {warning_value}"
                output.append(line)
            
            return output
        
        # Special handling for DEMO-LIQ-NEEDS-TIMEFRM-1 type items
        if "DEMO-LIQ-NEEDS-TIMEFRM-1" in demo_field.upper():
            output.append("************************************************************************")
            output.append("****")
            output.append(f"**** RECORD TYPE {begin_trimmed} ({demo_field.upper()})")
            output.append("****")
            output.append(f"**** {demo_field.upper()}   {begin_trimmed}  (TABLE D)")
            output.append("****")
            output.append("****   COL 007-010 = PER-MAIN1-LQDITY-NEEDS-CD")
            output.append(f"****   COL 013-013 = {demo_field.upper()}")
            output.append("****   COL 015-015 = WARNING REPORT(Y/N)")
            output.append("****")
            output.append("****  ****  * *")
            
            # Add range data with proper spacing based on column specifications
            for range_info in range_data:
                min_val = range_info['min_val']
                max_val = range_info['max_val']
                code = range_info['code']
                warning_report = range_info.get('warning_report', '')
                
                # Handle warning report: only show Y/N if there's a value, otherwise space
                if warning_report and warning_report.upper() in ['Y', 'N']:
                    warning_value = warning_report.upper()
                else:
                    warning_value = ' '  # Space if no warning report or invalid value
                
                # Format with proper spacing: COL 007-010 (4 chars), COL 013-013 (1 char), COL 015-015 (1 char)
                # Position 1-6: begin (e.g., "1092  ")
                # Position 7-10: PER-MAIN1-LQDITY-NEEDS-CD (4 characters)
                # Position 11-12: spaces
                # Position 13: DEMO-LIQ-NEEDS-TIMEFRM-1 (1 character)
                # Position 14: space
                # Position 15: WARNING REPORT (1 character)
                
                line = f"{begin_trimmed}  {min_val:<4}  {code} {warning_value}"
                output.append(line)
            
            return output
        
        if is_range_data:
            # Range-based data format
            output.append("************************************************************************")
            output.append("****")
            output.append(f"**** {demo_field} {begin_trimmed}{table_reference.upper()} (TABLE {table_reference.upper()} #1-TOP ON MAPPING)")
            output.append("****  ***************  ***************  *  *")
            
            # Add range data
            for range_info in range_data:
                min_val = range_info['min_val']
                max_val = range_info['max_val']
                code = range_info['code']
                
                # Format the line - handle special cases for N codes
                if code.upper() == 'N' or code == ' ':
                    line = f"{begin_trimmed}{table_reference.upper()}  {min_val}  {max_val}     N"
                else:
                    line = f"{begin_trimmed}{table_reference.upper()}  {min_val}  {max_val}  {code}"
                
                output.append(line)
        else:
            # Non-range data format
            output.append("************************************************************************")
            output.append("****")
            output.append(f"**** RECORD TYPE {begin_trimmed}{table_reference.upper()} ({demo_field})")
            output.append("****")
            output.append("****  *  *  *")
            
            # Add non-range data
            for range_info in range_data:
                min_val = range_info['min_val']
                max_val = range_info['max_val']
                code = range_info['code']
                warning_report = range_info.get('warning_report', '')
                
                # For non-range data, format as: 301A  U or 301A  M  A or 301A  M  A  B
                if code and code.strip() and code != ' ':
                    if warning_report and warning_report.strip() and warning_report != ' ':
                        line = f"{begin_trimmed}{table_reference.upper()}  {min_val}  {code}  {warning_report}"
                    else:
                        line = f"{begin_trimmed}{table_reference.upper()}  {min_val}  {code}"
                else:
                    if warning_report and warning_report.strip() and warning_report != ' ':
                        line = f"{begin_trimmed}{table_reference.upper()}  {min_val}  {warning_report}"
                    else:
                        line = f"{begin_trimmed}{table_reference.upper()}  {min_val}"
                
                output.append(line)
        
        # Add closing separator
        output.append("****")
        output.append("************************************************************************")
        
        return output

# Test function
if __name__ == "__main__":
    converter = ExcelToTextConverter()
    try:
        result = converter.convert_excel_to_text("Input1.xlsx")
        print("SUCCESS: Conversion completed")
        print("Complete output:")
        lines = result.split('\n')
        for i, line in enumerate(lines):
            print(f"{i+1:3d}: {line}")
        print(f"\nTotal lines: {len(lines)}")
    except Exception as e:
        print(f"ERROR: {e}")
