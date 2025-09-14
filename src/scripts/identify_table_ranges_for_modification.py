"""
Tool 1: Identify Table Ranges for Modification

This tool uses LLM analysis to identify table ranges and extract table-range specific global context.
Global context per table range is identified only once and preserved throughout the orchestrator workflow.

Key Features:
- Uses existing LLM prompts for comprehensive analysis
- Applies 100% consistency rule for global items per table range
- Directly modifies state object as per orchestrator architecture
- Handles period normalization and operation classification
"""

import sys
import os
import json
import re
from typing import Dict, Any, List
from dotenv import load_dotenv
from openai import AzureOpenAI

# Add parent directory to path to import existing modules
sys.path.append(os.path.join(os.path.dirname(__file__), '..'))

# Import existing functionality
from prompts import create_llm_analysis_system_prompt, create_llm_analysis_user_prompt

# Load environment variables
load_dotenv()


def _is_valid_period_header(cell_content: str) -> bool:
    """
    Validate if cell content is actually a period header vs numeric data
    
    Args:
        cell_content (str): Cell content to validate
        
    Returns:
        bool: True if it's a valid period header
    """
    if not cell_content:
        return False
    
    cell_content = str(cell_content).strip().upper()
    
    # Reject pure numeric values
    if re.match(r'^-?\d*\.?\d+$', cell_content):
        return False
    
    # Reject very long decimal numbers (likely data)
    if '.' in cell_content and len(cell_content.split('.')[1]) > 2:
        return False
    
    # Accept quarter patterns
    if re.search(r'Q[1-4]', cell_content):
        return True
    
    # Accept FY patterns
    if re.search(r'FY\s*\d{2}', cell_content):
        return True
    
    # Accept CY patterns  
    if re.search(r'CY\s*\d{4}', cell_content):
        return True
    
    # Reject everything else
    return False


def normalize_period_for_database(display_period):
    """
    Convert any period format to database-compatible format
    Handles all common Excel display formats
    """
    if not display_period:
        return display_period
    
    # Remove common separators and spaces for pattern matching
    cleaned = re.sub(r'[^\w]', '', str(display_period).upper())
    
    # Quarter patterns (highest priority)
    if re.match(r'Q\d+\d{2}', cleaned):
        # Q325 -> Q3 FY25, Q226 -> Q2 FY26
        quarter = cleaned[0:2]
        year = cleaned[2:]
        return f"{quarter} FY{year}"
    
    # Quarter with space patterns (Q2 26, Q4 25)
    elif re.match(r'Q\d+\s*\d{2}', str(display_period).upper()):
        parts = re.findall(r'Q(\d+)\s*(\d{2})', str(display_period).upper())
        if parts:
            quarter, year = parts[0]
            return f"Q{quarter} FY{year}"
    
    # Quarter with FY patterns (Q2FY26, Q1 FY25)
    elif re.match(r'Q\d+\s*FY\s*\d{2}', str(display_period).upper()):
        parts = re.findall(r'Q(\d+)\s*FY\s*(\d{2})', str(display_period).upper())
        if parts:
            quarter, year = parts[0]
            return f"Q{quarter} FY{year}"
    
    # Financial year patterns (FY25, FY 25)
    elif re.match(r'FY\s*\d{2}', str(display_period).upper()):
        year = re.findall(r'FY\s*(\d{2})', str(display_period).upper())[0]
        return f"FY{year}"
    
    # Calendar year patterns (2024, CY24, CY 24)
    elif re.match(r'(CY\s*)?\d{4}', str(display_period).upper()):
        year = re.findall(r'(\d{4})', str(display_period))[0]
        return f"CY{year}"
    
    # Default: return as-is 
    return str(display_period)


def extract_period_mapping_from_excel(excel_data: str, table_range: str) -> Dict[str, str]:
    """
    Extract column-to-period mapping from Excel markdown data for a specific table range
    
    Args:
        excel_data (str): Excel data in markdown format
        table_range (str): Table range like "A5:D15"
        
    Returns:
        Dict[str, str]: Column letter to normalized period mapping
    """
    period_mapping = {}
    
    try:
        # Parse table range to get column bounds
        range_parts = table_range.split(':')
        if len(range_parts) != 2:
            return period_mapping
            
        start_cell = range_parts[0]  # A5
        end_cell = range_parts[1]    # D15
        
        # Extract column letters (A, B, C, D...)
        start_col = re.match(r'([A-Z]+)', start_cell).group(1)
        end_col = re.match(r'([A-Z]+)', end_cell).group(1)
        
        # Convert to column numbers for iteration
        def col_to_num(col):
            num = 0
            for char in col:
                num = num * 26 + (ord(char) - ord('A') + 1)
            return num
        
        def num_to_col(num):
            col = ""
            while num > 0:
                num -= 1
                col = chr(num % 26 + ord('A')) + col
                num //= 26
            return col
        
        start_col_num = col_to_num(start_col)
        end_col_num = col_to_num(end_col)
        
        # Look for period patterns in Excel data - improved validation
        period_indicators = [
            r'Q[1-4]\s*\d{2}',      # Q1 25, Q2 26
            r'Q[1-4]\s*FY\s*\d{2}', # Q1 FY25, Q2 FY26  
            r'FY\s*\d{2}',          # FY25, FY 26
            r'CY\s*\d{4}',          # CY2024, CY 2024
        ]
        
        # Split Excel data into lines for parsing
        lines = excel_data.split('\n')
        
        # Scan the first few rows above and within the table for period headers
        for line_idx, line in enumerate(lines[:20]):  # Check first 20 lines for headers
            # Skip empty lines or non-data lines
            if '|' not in line:
                continue
                
            # Look for period patterns in each column
            cells = line.split('|')
            
            for col_num in range(start_col_num, end_col_num + 1):
                col_letter = num_to_col(col_num)
                
                # Skip if we already found a period for this column
                if col_letter in period_mapping:
                    continue
                
                # Check if this column index exists in the line
                # CRITICAL FIX: Account for row number column at index 1 in markdown table
                cell_index = col_num + 1  # Offset by 1 to account for row number column
                if cell_index < len(cells):
                    cell_content = cells[cell_index].strip()
                    
                    # Skip empty cells, numbers, and obvious data values
                    if not cell_content or cell_content in ['', 'None', 'nan']:
                        continue
                    
                    # Skip if it's clearly numeric data (decimal numbers, percentages)
                    if re.match(r'^-?\d*\.?\d+$', cell_content):  # Pure numbers
                        continue
                    if re.match(r'^-?\d*\.?\d+%?$', cell_content):  # Numbers with %
                        continue
                    
                    # Check for period patterns
                    for pattern in period_indicators:
                        if re.search(pattern, cell_content, re.IGNORECASE):
                            # Additional validation: ensure it looks like a period, not data
                            if _is_valid_period_header(cell_content):
                                normalized_period = normalize_period_for_database(cell_content)
                                period_mapping[col_letter] = normalized_period
                                print(f"📅 Found period mapping: Column {col_letter} = '{cell_content}' → '{normalized_period}'")
                                break
        
        print(f"📊 Period mapping for {table_range}: {period_mapping}")
        return period_mapping
        
    except Exception as e:
        print(f"⚠️  Error extracting period mapping: {e}")
        return {}


def _initialize_sheet_period_mapping(state: Dict[str, Any], tables_list: List[Dict[str, Any]]):
    """
    Initialize the sheet-global period mapping from all identified tables
    
    Args:
        state: Agent state object
        tables_list: List of processed tables with period mappings
    """
    try:
        # Initialize sheet-global mappings
        if "sheet_period_mapping" not in state:
            state["sheet_period_mapping"] = {}
        if "sheet_columns_added" not in state:
            state["sheet_columns_added"] = []
        
        # Collect all period mappings from all tables
        all_periods = {}
        
        for table in tables_list:
            table_period_mapping = table.get("global_items", {}).get("period_mapping", {})
            for column, period in table_period_mapping.items():
                # Normalize the period for consistency
                normalized_period = normalize_period_for_database(period)
                all_periods[column] = normalized_period
                print(f"🌍 Found period in table {table.get('range', 'N/A')}: Column {column} = '{period}' → '{normalized_period}'")
        
        # Update the sheet-global period mapping
        state["sheet_period_mapping"].update(all_periods)
        
        print(f"🌍 Initialized SHEET-GLOBAL period mapping: {state['sheet_period_mapping']}")
        print(f"🌍 Found {len(all_periods)} existing period columns in the sheet")
        
    except Exception as e:
        print(f"⚠️  Failed to initialize sheet period mapping: {e}")
        # Don't raise - this is not critical enough to halt the process


def get_azure_openai_client():
    """Initialize Azure OpenAI client - reused from agent.py"""
    try:
        azure_deployment = os.getenv('AZURE_DEPLOYMENT')
        if azure_deployment.startswith('https://'):
            azure_endpoint = azure_deployment
        else:
            azure_endpoint = f"https://{azure_deployment}.openai.azure.com/"
        
        client = AzureOpenAI(
            azure_endpoint=azure_endpoint,
            api_key=os.getenv('OPENAI_API_KEY'),
            api_version=os.getenv('OPENAI_API_VERSION')
        )
        return client
    except Exception as e:
        print(f"❌ Failed to initialize Azure OpenAI client: {e}")
        return None


def identify_table_ranges_for_modification(
    excel_data: str,
    user_question: str,
    operation_type: str,
    target_period: str,
    processed_tables: List[str],
    state: Dict[str, Any]
) -> None:
    """
    Identify table ranges and extract table-range specific global context.
    
    This tool directly modifies the state object with:
    - identified_tables: List of tables with preserved global context
    - operation_type: Determined operation type
    - target_period: Normalized target period
    - processing_status: Updated to "tables_identified"
    
    Args:
        excel_data (str): Current Excel data in markdown format
        user_question (str): User's natural language request
        operation_type (str): Current operation type (may be empty on first run)
        target_period (str): Current target period (may be empty on first run)
        processed_tables (List[str]): List of already processed table ranges
        state (Dict[str, Any]): Agent state object to modify directly
        
    Raises:
        Exception: If LLM analysis fails or returns invalid JSON
    """
    try:
        print(f"🔍 Tool 1: Identifying table ranges for modification...")
        print(f"📊 Excel data length: {len(excel_data)} characters")
        print(f"❓ User question: {user_question}")
        print(f"✅ Already processed: {len(processed_tables)} tables")
        
        # Get Azure OpenAI client
        client = get_azure_openai_client()
        if not client:
            raise Exception("Failed to initialize Azure OpenAI client")
        
        # Prepare metadata from state
        excel_metadata = state.get("excel_metadata", {})
        
        # Use existing LLM prompts
        system_prompt = create_llm_analysis_system_prompt()
        user_prompt = create_llm_analysis_user_prompt(excel_data, user_question, excel_metadata)
        
        print(f"🤖 Calling LLM for comprehensive table analysis...")
        
        # Make Azure OpenAI API call
        response = client.chat.completions.create(
            model=os.getenv('DEPLOYMENT_NAME'),
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt}
            ],
            temperature=0,  # Low temperature for consistent analysis
            max_tokens=4000,
            response_format={"type": "json_object"}
        )
        
        # Parse LLM response
        llm_response_text = response.choices[0].message.content
        llm_analysis = json.loads(llm_response_text)
        
        print(f"✅ LLM Analysis Complete!")
        print(f"🎯 Operation: {llm_analysis.get('operation_analysis', {}).get('operation_type', 'unknown')}")
        print(f"📅 Target Period: {llm_analysis.get('operation_analysis', {}).get('target_period', 'unknown')}")
        print(f"📋 Tables Found: {len(llm_analysis.get('table_analysis', {}).get('identified_tables', []))}")
        
        # Extract key information from LLM analysis
        operation_analysis = llm_analysis.get("operation_analysis", {})
        table_analysis = llm_analysis.get("table_analysis", {})
        table_global_contexts = llm_analysis.get("table_global_contexts", {})
        validation = llm_analysis.get("validation", {})
        
        # Process identified tables and add global context + modification requirements
        processed_tables_list = []
        identified_tables = table_analysis.get("identified_tables", [])
        
        for table_info in identified_tables:
            table_range = table_info.get("range", "")
            
            # Skip if already processed
            if table_range in processed_tables:
                print(f"⏩ Skipping already processed table: {table_range}")
                continue
            
            # Extract global context for this specific table range
            global_context = table_global_contexts.get(table_range, {})
            global_items = {}
            
            # Process global items with 100% consistency rule
            for item_name in ["company_name", "entity", "metric_type"]:
                item_data = global_context.get(item_name, {})
                if item_data.get("is_global", False):
                    global_items[item_name] = item_data.get("value", "")
                else:
                    global_items[item_name] = ""  # Not global - empty string
            
            # Extract period mapping for this table range (NEW!)
            period_mapping = extract_period_mapping_from_excel(excel_data, table_range)
            global_items["period_mapping"] = period_mapping
            
            # Determine modification requirements based on operation type
            operation_type_resolved = operation_analysis.get("operation_type", "")
            target_period_resolved = operation_analysis.get("target_period", "")
            
            modification_required = "none"
            needs_new_column = False
            needs_new_rows = False
            
            if operation_type_resolved == "add_column":
                modification_required = "add_column_after_last"
                needs_new_column = True
            elif operation_type_resolved == "add_metrics":
                modification_required = "add_rows_after_last"
                needs_new_rows = True
            # update_existing doesn't need modification
            
            # Create complete table information with preserved global context
            processed_table = {
                "range": table_range,
                "description": table_info.get("description", ""),
                "relevance_score": table_info.get("relevance_score", 0.0),
                "table_type": table_info.get("table_type", ""),
                "structure": table_info.get("structure", {}),
                "needs_new_column": needs_new_column,
                "needs_new_rows": needs_new_rows,
                "modification_required": modification_required,
                "global_items": global_items,  # Identified once, preserved forever
                "global_analysis_confidence": {
                    item: global_context.get(item, {}).get("confidence", 0.0)
                    for item in ["company_name", "entity", "metric_type"]
                },
                "global_sources": {
                    item: global_context.get(item, {}).get("source", "")
                    for item in ["company_name", "entity", "metric_type"]
                }
            }
            
            processed_tables_list.append(processed_table)
            print(f"📋 Processed table {table_range}: {global_items}")
        
        # DIRECTLY MODIFY STATE OBJECT
        state["identified_tables"] = processed_tables_list
        state["operation_type"] = operation_analysis.get("operation_type", "")
        state["target_period"] = operation_analysis.get("target_period", "")
        state["processing_status"] = "tables_identified"
        state["current_table_index"] = 0  # Start with first table
        
        # Initialize sheet-global period mapping from all identified tables
        _initialize_sheet_period_mapping(state, processed_tables_list)
        
        # Store full LLM analysis for reference
        state["llm_analysis"] = llm_analysis
        
        # Handle validation issues
        if not validation.get("feasible", True):
            state["errors"].extend(validation.get("errors", []))
            if validation.get("errors"):
                state["processing_status"] = "error"
                return
        
        state["warnings"].extend(validation.get("warnings", []))
        
        print(f"✅ Tool 1 Complete: Identified {len(processed_tables_list)} tables for processing")
        print(f"🎯 Operation: {state['operation_type']}")
        print(f"📅 Target Period: {state['target_period']}")
        print(f"📊 Status: {state['processing_status']}")
        
    except json.JSONDecodeError as e:
        error_msg = f"Failed to parse LLM response as JSON: {str(e)}"
        state["errors"].append(error_msg)
        state["processing_status"] = "error"
        print(f"❌ Tool 1 JSON Error: {error_msg}")
        raise Exception(error_msg)
        
    except Exception as e:
        error_msg = f"Error in identify_table_ranges_for_modification: {str(e)}"
        state["errors"].append(error_msg)
        state["processing_status"] = "error"
        print(f"❌ Tool 1 Error: {error_msg}")
        raise Exception(error_msg)


def validate_table_global_context(table_range: str, global_items: Dict[str, str]) -> Dict[str, Any]:
    """
    Validate the extracted global context for a table range
    
    Args:
        table_range (str): Excel table range (e.g., "A5:D15")
        global_items (Dict[str, str]): Extracted global items
        
    Returns:
        Dict[str, Any]: Validation results with warnings and recommendations
    """
    validation_result = {
        "valid": True,
        "warnings": [],
        "recommendations": []
    }
    
    # Check for empty global items (indicating non-global context)
    empty_items = [item for item, value in global_items.items() if not value]
    if empty_items:
        validation_result["warnings"].append(
            f"Non-global items in {table_range}: {empty_items}. Will use cell-specific extraction."
        )
    
    # Check for potential inconsistencies
    if global_items.get("company_name") and not global_items.get("entity"):
        validation_result["recommendations"].append(
            f"Consider setting entity = company_name for {table_range}"
        )
    
    return validation_result


if __name__ == "__main__":
    """
    Test the table identification tool with sample data
    """
    print("=== Testing Table Identification Tool ===")
    
    # Mock test data
    sample_excel_data = """
    |    | A                   | B      | C      | D      |
    |----|---------------------|--------|--------|--------|
    |  1 | HDFC Bank          |        | Q3 25  | Q4 25  |
    |  2 |                    |        |        |        |
    |  3 | Key Financial Metrics|       |        |        |
    |  4 | Loan Growth %      | 3%     | 6%     | 7%     |
    |  5 | Deposit Growth %   | 16%    | 14%    | 16%    |
    """
    
    sample_user_question = "fill data for Q1 26"
    
    # Mock state object
    test_state = {
        "excel_metadata": {
            "sheet_info": {
                "name": "Test Sheet",
                "rows": 5,
                "cols": 4
            }
        },
        "errors": [],
        "warnings": []
    }
    
    try:
        # Test the tool
        identify_table_ranges_for_modification(
            excel_data=sample_excel_data,
            user_question=sample_user_question,
            operation_type="",
            target_period="",
            processed_tables=[],
            state=test_state
        )
        
        print("✅ Test completed successfully!")
        print(f"Identified tables: {len(test_state.get('identified_tables', []))}")
        print(f"Operation type: {test_state.get('operation_type')}")
        print(f"Target period: {test_state.get('target_period')}")
        
    except Exception as e:
        print(f"❌ Test failed: {e}")
