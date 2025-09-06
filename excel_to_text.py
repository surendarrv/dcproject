#!/usr/bin/env python3
"""
Excel to Text Converter Application
A focused application that reads Excel files and converts them to text format.
"""

import os
import sys
import argparse
from pathlib import Path
from typing import Optional, Dict, Any, List, Union
import pandas as pd
import logging

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


class ExcelToTextConverter:
    """A streamlined class for converting Excel files to text files."""
    
    def __init__(self):
        """Initialize the converter."""
        logger.info("Initialized ExcelToTextConverter")
    
    def read_excel_file(self, excel_path: str, sheet_name: Optional[str] = None) -> pd.DataFrame:
        """Read an Excel file and return as pandas DataFrame."""
        try:
            excel_path = Path(excel_path)
            if not excel_path.exists():
                raise FileNotFoundError(f"Excel file not found: {excel_path}")
            
            if sheet_name:
                df = pd.read_excel(excel_path, sheet_name=sheet_name)
                logger.info(f"Successfully read sheet '{sheet_name}' from Excel file: {excel_path}")
            else:
                df = pd.read_excel(excel_path)
                logger.info(f"Successfully read Excel file: {excel_path}")
            
            return df
        except Exception as e:
            logger.error(f"Error reading Excel file {excel_path}: {str(e)}")
            raise
    
    def filter_data(self, df: pd.DataFrame, 
                   columns_to_keep: Optional[List[str]] = None,
                   row_filters: Optional[Dict[str, Union[str, List[str]]]] = None,
                   exclude_empty: bool = True) -> pd.DataFrame:
        """
        Filter DataFrame to extract only relevant data.
        
        Args:
            df: Input DataFrame
            columns_to_keep: List of column names to keep
            row_filters: Dict of column:value pairs to filter rows
            exclude_empty: Whether to exclude rows with empty values
        
        Returns:
            Filtered DataFrame
        """
        filtered_df = df.copy()
        
        # Filter columns
        if columns_to_keep:
            available_columns = [col for col in columns_to_keep if col in filtered_df.columns]
            if available_columns:
                filtered_df = filtered_df[available_columns]
                logger.info(f"Filtered to columns: {available_columns}")
        
        # Filter rows based on conditions
        if row_filters:
            for column, value in row_filters.items():
                if column in filtered_df.columns:
                    if isinstance(value, list):
                        filtered_df = filtered_df[filtered_df[column].isin(value)]
                    else:
                        filtered_df = filtered_df[filtered_df[column] == value]
                    logger.info(f"Applied row filter: {column} = {value}")
        
        # Remove empty rows
        if exclude_empty:
            original_rows = len(filtered_df)
            filtered_df = filtered_df.dropna(how='all')
            logger.info(f"Removed empty rows: {original_rows} -> {len(filtered_df)}")
        
        logger.info(f"Data filtering complete: {len(df)} rows -> {len(filtered_df)} rows")
        return filtered_df
    
    def convert_to_text(self, df: pd.DataFrame, 
                       output_format: str = "table",
                       delimiter: str = " | ",
                       include_headers: bool = True) -> str:
        """
        Convert DataFrame to text format.
        
        Args:
            df: DataFrame to convert
            output_format: Format type ('table', 'list', 'csv')
            delimiter: Delimiter for table format
            include_headers: Whether to include column headers
        
        Returns:
            Formatted text string
        """
        if df.empty:
            return "No data to convert."
        
        text_lines = []
        
        if output_format == "table":
            # Table format with delimiters
            if include_headers:
                headers = delimiter.join(str(col) for col in df.columns)
                text_lines.append(headers)
                text_lines.append("-" * len(headers))
            
            for _, row in df.iterrows():
                row_text = delimiter.join(str(val) if pd.notna(val) else "" for val in row)
                text_lines.append(row_text)
        
        elif output_format == "list":
            # List format - each record on separate lines
            for i, (_, row) in enumerate(df.iterrows(), 1):
                text_lines.append(f"Record {i}:")
                for col, val in row.items():
                    if pd.notna(val):
                        text_lines.append(f"  {col}: {val}")
                text_lines.append("")  # Empty line between records
        
        elif output_format == "csv":
            # CSV format
            if include_headers:
                text_lines.append(",".join(str(col) for col in df.columns))
            
            for _, row in df.iterrows():
                row_text = ",".join(f'"{str(val)}"' if pd.notna(val) else '""' for val in row)
                text_lines.append(row_text)
        
        return "\n".join(text_lines)
    
    def convert_excel_to_text(self, 
                             excel_path: str,
                             output_path: str,
                             sheet_name: Optional[str] = None,
                             columns_to_keep: Optional[List[str]] = None,
                             row_filters: Optional[Dict[str, Union[str, List[str]]]] = None,
                             output_format: str = "table",
                             delimiter: str = " | ",
                             include_headers: bool = True,
                             exclude_empty: bool = True) -> None:
        """
        Convert Excel file to text file with optional filtering.
        
        Args:
            excel_path: Path to input Excel file
            output_path: Path to output text file
            sheet_name: Specific sheet to read (None for first sheet)
            columns_to_keep: List of column names to keep
            row_filters: Row filtering conditions
            output_format: Text output format ('table', 'list', 'csv')
            delimiter: Delimiter for table format
            include_headers: Include column headers
            exclude_empty: Exclude empty rows
        """
        try:
            print(f"🔄 Reading Excel file: {excel_path}")
            
            # Read Excel file
            df = self.read_excel_file(excel_path, sheet_name)
            print(f"📊 Found {len(df)} rows and {len(df.columns)} columns")
            
            # Filter data if needed
            if columns_to_keep or row_filters or exclude_empty:
                print("🔍 Applying filters...")
                df = self.filter_data(df, columns_to_keep, row_filters, exclude_empty)
            
            # Convert to text
            print(f"📝 Converting to {output_format} format...")
            text_content = self.convert_to_text(df, output_format, delimiter, include_headers)
            
            # Write to text file
            output_file_path = Path(output_path)
            output_file_path.parent.mkdir(parents=True, exist_ok=True)
            
            with open(output_file_path, 'w', encoding='utf-8') as f:
                f.write(text_content)
            
            print(f"✅ Successfully converted Excel to text: {output_file_path}")
            print(f"📈 Final output: {len(df)} rows processed")
            
        except Exception as e:
            logger.error(f"Error in Excel to text conversion: {str(e)}")
            print(f"❌ Error: {str(e)}")
            raise
    
    def get_excel_info(self, excel_path: str) -> Dict[str, Any]:
        """Get information about an Excel file."""
        try:
            excel_path = Path(excel_path)
            if not excel_path.exists():
                raise FileNotFoundError(f"Excel file not found: {excel_path}")
            
            excel_file = pd.ExcelFile(excel_path)
            info = {
                'file_path': str(excel_path),
                'sheet_names': excel_file.sheet_names,
                'num_sheets': len(excel_file.sheet_names)
            }
            
            # Get info for each sheet
            sheet_info = {}
            for sheet_name in excel_file.sheet_names:
                df = pd.read_excel(excel_path, sheet_name=sheet_name)
                sheet_info[sheet_name] = {
                    'rows': len(df),
                    'columns': len(df.columns),
                    'column_names': list(df.columns)
                }
            
            info['sheets'] = sheet_info
            logger.info(f"Retrieved Excel file info: {excel_path}")
            return info
            
        except Exception as e:
            logger.error(f"Error getting Excel file info: {str(e)}")
            return {}


def main():
    """Main function with command-line interface."""
    parser = argparse.ArgumentParser(
        description="Convert Excel files to text format",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Basic conversion
  python excel_to_text.py input.xlsx output.txt
  
  # Specify sheet and format
  python excel_to_text.py input.xlsx output.txt --sheet "Sheet1" --format list
  
  # Filter specific columns
  python excel_to_text.py input.xlsx output.txt --columns "Name,Age,Email"
  
  # Get Excel file information
  python excel_to_text.py input.xlsx --info
        """
    )
    
    parser.add_argument('excel_file', help='Path to input Excel file')
    parser.add_argument('output_file', nargs='?', help='Path to output text file')
    parser.add_argument('--sheet', '-s', help='Sheet name to read (default: first sheet)')
    parser.add_argument('--format', '-f', choices=['table', 'list', 'csv'], 
                       default='table', help='Output format (default: table)')
    parser.add_argument('--columns', '-c', help='Comma-separated list of columns to keep')
    parser.add_argument('--delimiter', '-d', default=' | ', help='Delimiter for table format')
    parser.add_argument('--no-headers', action='store_true', help='Exclude column headers')
    parser.add_argument('--info', '-i', action='store_true', help='Show Excel file information only')
    
    args = parser.parse_args()
    
    # Validate input file
    if not Path(args.excel_file).exists():
        print(f"❌ Error: Excel file not found: {args.excel_file}")
        sys.exit(1)
    
    converter = ExcelToTextConverter()
    
    # Show file info if requested
    if args.info:
        print("📋 Excel File Information:")
        print("=" * 40)
        info = converter.get_excel_info(args.excel_file)
        if info:
            print(f"File: {info['file_path']}")
            print(f"Sheets: {info['num_sheets']}")
            for sheet_name, sheet_info in info['sheets'].items():
                print(f"\nSheet '{sheet_name}':")
                print(f"  Rows: {sheet_info['rows']}")
                print(f"  Columns: {sheet_info['columns']}")
                print(f"  Column names: {', '.join(sheet_info['column_names'])}")
        return
    
    # Validate output file
    if not args.output_file:
        print("❌ Error: Output file is required unless using --info")
        sys.exit(1)
    
    # Parse columns if provided
    columns_to_keep = None
    if args.columns:
        columns_to_keep = [col.strip() for col in args.columns.split(',')]
    
    # Convert Excel to text
    try:
        converter.convert_excel_to_text(
            excel_path=args.excel_file,
            output_path=args.output_file,
            sheet_name=args.sheet,
            columns_to_keep=columns_to_keep,
            output_format=args.format,
            delimiter=args.delimiter,
            include_headers=not args.no_headers
        )
    except Exception as e:
        print(f"❌ Conversion failed: {str(e)}")
        sys.exit(1)


if __name__ == "__main__":
    main()