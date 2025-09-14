"""
Excel Agent - Single Orchestrator Architecture

This module implements a LangGraph-based AI agent with a single orchestrator node
that dynamically calls specialized tools for Excel financial template filling.

Key Features:
- Single orchestrator node with LLM-driven decision making
- Dynamic tool calling based on current state
- Iterative table processing with global context preservation  
- Human-in-the-loop intervention with global toggle
- Error handling that halts entire process for data integrity
"""

from langgraph.graph import StateGraph
from typing import TypedDict, List, Dict, Any
import sys
import os
import json
import shutil
from datetime import datetime
from openai import AzureOpenAI
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Add parent directory to path to import from scripts
sys.path.append(os.path.join(os.path.dirname(__file__), '..'))

# Import existing utilities
from scripts.excel_to_markdown import parse_sheet_xlsx_with_mapping

# Import our three specialized tools
from scripts.identify_table_ranges_for_modification import identify_table_ranges_for_modification
from scripts.modify_excel_sheet import modify_excel_sheet
from scripts.cell_mapping_and_fill_current_table import cell_mapping_and_fill_current_table

# Import enhanced logging
sys.path.append(os.path.dirname(__file__))
from enhanced_logging import create_logger

# Global toggle for human intervention
ENABLE_HUMAN_INTERVENTION = False


def normalize_period_for_database(display_period):
    """
    Convert any period format to database-compatible format
    Handles all common Excel display formats
    """
    if not display_period:
        return display_period
    
    import re
    
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


def create_excel_copy(original_path: str) -> str:
    """
    Create a timestamped copy of the Excel file for processing
    
    Args:
        original_path (str): Path to the original Excel file
        
    Returns:
        str: Path to the created copy
        
    Raises:
        Exception: If file copy fails
    """
    try:
        # Extract directory and filename components
        directory = os.path.dirname(original_path)
        filename = os.path.basename(original_path)
        name, ext = os.path.splitext(filename)
        
        # Create timestamped copy name
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        copy_name = f"{name}_copy_{timestamp}{ext}"
        copy_path = os.path.join(directory, copy_name)
        
        # Create the copy
        shutil.copy2(original_path, copy_path)
        
        print(f"📋 Created Excel copy: {filename} → {copy_name}")
        print(f"📁 Copy location: {copy_path}")
        
        return copy_path
        
    except Exception as e:
        print(f"❌ Failed to create Excel copy: {e}")
        raise Exception(f"Could not create copy of {original_path}: {str(e)}")


class AgentState(TypedDict):
    """
    Comprehensive state structure for the orchestrator architecture.
    Matches the specification in agent-logic.md
    """
    # === INPUT & CURRENT EXCEL STATE ===
    excel_file_path: str                      # Path to Excel file (gets modified during processing)
    excel_data: str                           # Current parsed Excel state (refreshed each iteration)
    user_question: str                        # Original user request
    
    # === ORCHESTRATOR CONTROL ===
    identified_tables: List[Dict[str, Any]]   # All tables with global context (identified once)
    processed_tables: List[str]               # Table ranges already completed
    current_table_index: int                  # Which table we're currently processing
    current_table: Dict[str, Any]             # Current table being worked on
    
    # === SHEET-GLOBAL CONTEXT ===
    sheet_period_mapping: Dict[str, str]      # GLOBAL: Column -> Period mapping for entire sheet
    sheet_columns_added: List[str]            # Track which period columns have been added to sheet
    period_exists_globally: bool              # Track if target period already exists globally
    
    # === OPERATION CONTEXT ===
    operation_type: str                       # "add_column", "update_existing", "add_metrics"
    target_period: str                        # e.g., "Q2 FY26" (normalized format)
    
    # === ITERATION STATE ===
    processing_status: str                    # "start", "tables_identified", "excel_modified", "filling_data", "complete"
    current_iteration: int                    # Track iterations for debugging
    
    # === HUMAN-IN-THE-LOOP CONTROL ===
    human_intervention_enabled: bool          # Global toggle for human intervention
    pending_human_approval: bool              # Flag for awaiting human decision
    human_decision: Dict[str, Any]            # Human's decision response
    
    # === RESULTS & TRACKING ===
    table_processing_results: Dict[str, Any] # Results per table range
    total_cells_filled: int                   # Running count across all tables
    errors: List[str]                         # Errors encountered (halt on error)
    warnings: List[str]                       # Warnings or ambiguities
    
    # === METADATA ===
    excel_metadata: Dict[str, Any]            # Excel metadata (refreshed each iteration)
    llm_analysis: Dict[str, Any]              # LLM analysis results
    session_logger: Any                       # Single logger instance for entire session


def set_human_intervention_mode(enabled: bool):
    """Global function to enable/disable human intervention"""
    global ENABLE_HUMAN_INTERVENTION
    ENABLE_HUMAN_INTERVENTION = enabled  # FIX: Use the actual parameter value
    print(f"🤖 Human intervention {'ENABLED' if enabled else 'DISABLED'}")


def initialize_agent_state(excel_file_path: str, user_question: str) -> AgentState:
    """
    Initialize agent state with human intervention setting and default values
    
    Args:
        excel_file_path (str): Path to Excel file to process
        user_question (str): User's natural language request
        
    Returns:
        AgentState: Initialized state object
    """
    return {
        # Input & Excel State
        "excel_file_path": excel_file_path,
        "excel_data": "",
        "user_question": user_question,
        
        # Orchestrator Control
        "identified_tables": [],
        "processed_tables": [],
        "current_table_index": 0,
        "current_table": {},
        
        # Sheet-Global Context
        "sheet_period_mapping": {},  # Track all periods in entire sheet
        "sheet_columns_added": [],   # Track which period columns were added
        "period_exists_globally": False,  # Track if target period already exists globally
        
        # Operation Context
        "operation_type": "",
        "target_period": "",
        
        # Iteration State
        "processing_status": "start",
        "current_iteration": 0,
        
        # Human-in-the-Loop Control
        "human_intervention_enabled": ENABLE_HUMAN_INTERVENTION,
        "pending_human_approval": False,
        "human_decision": {},
        
        # Results & Tracking
        "table_processing_results": {},
        "total_cells_filled": 0,
        "errors": [],
        "warnings": [],
        
        # Metadata
        "excel_metadata": {},
        "llm_analysis": {},
        "session_logger": None  # Will be set when agent starts
    }


def get_azure_openai_client():
    """Initialize Azure OpenAI client with environment variables"""
    try:
        # Handle both full URL and deployment name formats
        azure_deployment = os.getenv('AZURE_DEPLOYMENT')
        if azure_deployment and azure_deployment.startswith('https://'):
            # Full URL provided
            azure_endpoint = azure_deployment
        else:
            # Just deployment name provided
            azure_endpoint = f"https://{azure_deployment}.openai.azure.com/"
        
        client = AzureOpenAI(
            azure_endpoint=azure_endpoint,
            api_key=os.getenv('OPENAI_API_KEY'),
            api_version=os.getenv('OPENAI_API_VERSION')
        )
        
        print(f"🔗 Azure OpenAI client initialized")
        print(f"   Endpoint: {azure_endpoint}")
        print(f"   API Version: {os.getenv('OPENAI_API_VERSION')}")
        print(f"   Deployment: {os.getenv('DEPLOYMENT_NAME')}")
        
        return client
    except Exception as e:
        print(f"❌ Failed to initialize Azure OpenAI client: {e}")
        print(f"   AZURE_DEPLOYMENT: {os.getenv('AZURE_DEPLOYMENT')}")
        print(f"   OPENAI_API_VERSION: {os.getenv('OPENAI_API_VERSION')}")
        print(f"   DEPLOYMENT_NAME: {os.getenv('DEPLOYMENT_NAME')}")
        return None


def re_parse_excel_state(excel_file_path: str) -> Dict[str, Any]:
    """
    Re-parse Excel file to get current state (always called at start of each iteration)
    
    Args:
        excel_file_path (str): Path to Excel file
        
    Returns:
        Dict[str, Any]: Parsed Excel data and metadata
    """
    try:
        parsed_result = parse_sheet_xlsx_with_mapping(excel_file_path)
        return {
            "excel_data": parsed_result["markdown"],
            "excel_metadata": parsed_result["metadata"]
        }
    except Exception as e:
        print(f"❌ Failed to re-parse Excel file: {e}")
        return {
            "excel_data": "",
            "excel_metadata": {}
        }


def llm_reasoning_and_tool_decision(
    excel_data: str,
    user_question: str, 
    processing_status: str,
    processed_tables: List[str],
    identified_tables: List[Dict],
    current_table_index: int,
    sheet_period_mapping: Dict[str, str],
    sheet_columns_added: List[str],
    logger = None
) -> Dict[str, Any]:
    """
    LLM analyzes current state and decides next action with table-range context
    
    Args:
        excel_data (str): Current Excel data in markdown format
        user_question (str): User's original request
        processing_status (str): Current processing status
        processed_tables (List[str]): List of completed table ranges
        identified_tables (List[Dict]): All identified tables with context
        current_table_index (int): Index of current table being processed
        logger: Enhanced logger instance
        
    Returns:
        Dict[str, Any]: Tool decision with reasoning and parameters
    """
    try:
        client = get_azure_openai_client()
        if not client:
            raise Exception("Failed to initialize Azure OpenAI client")
        
        # Get current table info for better decision making
        current_table = None
        current_period_mapping = {}
        
        if 0 <= current_table_index < len(identified_tables):
            current_table = identified_tables[current_table_index]
            current_period_mapping = current_table.get("global_items", {}).get("period_mapping", {})
        
        system_prompt = """
You are an Excel modification orchestrator. Analyze the current state and decide the next action.

AVAILABLE TOOLS:
1. identify_table_ranges_for_modification - When no tables identified yet (first run)
2. modify_excel_sheet - When current table needs structural changes (add column/row)
3. cell_mapping_and_fill_current_table - When current table ready for data filling

CRITICAL DECISION LOGIC:
- If processing_status == "start": Use identify_table_ranges_for_modification
- If current table already has the target period column: Use cell_mapping_and_fill_current_table
- If current table needs the target period column added: Use modify_excel_sheet
- If current table is completely processed: Move to next table or complete

IMPORTANT RULES:
1. ALWAYS check if the target period (e.g., "Q2 FY26") already exists as a column header
2. Do NOT add duplicate columns - if "Q2 FY26" already exists, proceed to filling
3. Sequential table processing - complete one table before moving to next
4. If all tables are processed, return "complete"

PERIOD MAPPING RULES - CRITICAL FOR DECISION:
- SHEET-GLOBAL PERIODS: Period columns apply to ENTIRE sheet, not individual tables
- If period_exists_globally = True: ALWAYS use cell_mapping_and_fill_current_table
- If period_exists_globally = False: ALWAYS use modify_excel_sheet
- NEVER IGNORE the period_exists_globally flag - it is the definitive decision criterion
- Period columns are shared across all tables in the sheet
- Adding duplicate period columns is a critical error that must be avoided

EXCEL ANALYSIS:
- Look for column headers that match the target period
- Check if the current table range already includes the target period
- Verify if modifications have already been made

FOR modify_excel_sheet TOOL:
- Specify the exact cell where the new period header should be placed
- Look at the Excel structure to find the correct header row and column
- For table A5:D15, if adding column E, the header should typically go in E1 or E2
- Analyze existing headers to determine the correct row (where Q1, Q2, Q3, Q4 appear)

Return JSON: {
    "tool_name": "tool_to_call",  // CRITICAL: If period_exists_globally=True, use "cell_mapping_and_fill_current_table"
    "reasoning": "MUST reference period_exists_globally flag in reasoning",
    "parameters": {
        "target_cell": "E2"  // ONLY for modify_excel_sheet when period_exists_globally=False
    },
    "confidence": 0.9
}
"""
        
        # Enhanced user prompt with better context
        current_table_info = ""
        if current_table:
            current_table_info = f"""
CURRENT TABLE DETAILS:
- Range: {current_table.get('range', 'N/A')}
- Description: {current_table.get('description', 'N/A')}
- Needs new column: {current_table.get('needs_new_column', False)}
- Global items: {current_table.get('global_items', {})}
- Period mapping: {current_period_mapping}
"""
        
        # Extract target period for analysis and normalization - FIXED LOGIC
        target_period = None
        normalized_target_period = None
        if identified_tables and current_table_index < len(identified_tables):
            current_table_data = identified_tables[current_table_index]
            
            # CRITICAL FIX: Extract complete period from user question (not just last word)
            import re
            # Look for period patterns in user question
            period_patterns = [
                r'(Q[1-4]\s*\d{2})',        # Q1 25, Q2 26, etc.
                r'(Q[1-4]\s*FY\s*\d{2})',   # Q1 FY25, Q2 FY26, etc.
                r'(FY\s*\d{2})',            # FY25, FY26, etc.
                r'(CY\s*\d{4})'             # CY2024, CY2025, etc.
            ]
            
            for pattern in period_patterns:
                match = re.search(pattern, user_question, re.IGNORECASE)
                if match:
                    target_period = match.group(1)
                    print(f"🎯 Extracted target period from user question: '{target_period}'")
                    break
            
            # Fallback if no pattern found
            if not target_period:
                target_period = "Q2 FY26"  # Default fallback
                print(f"⚠️  No period pattern found in '{user_question}', using fallback: '{target_period}'")
                
            normalized_target_period = normalize_period_for_database(target_period)
            print(f"🔄 Normalized target period: '{target_period}' → '{normalized_target_period}'")
        
        # Check if target period already exists in SHEET-GLOBAL period mapping
        period_exists_globally = False
        period_found_in_column = None
        
        print(f"🔍 Period Detection Debug:")
        print(f"   Target period (raw): '{target_period}'")
        print(f"   Target period (normalized): '{normalized_target_period}'")
        print(f"   Sheet period mapping: {sheet_period_mapping}")
        print(f"   Sheet columns added: {sheet_columns_added}")
        
        # First check sheet-global mapping (most important)
        if sheet_period_mapping and normalized_target_period:
            print(f"🔍 Checking sheet-global mapping for '{normalized_target_period}'...")
            for col, period in sheet_period_mapping.items():
                normalized_existing_period = normalize_period_for_database(period)
                print(f"   Column {col}: '{period}' → '{normalized_existing_period}' (match: {normalized_existing_period == normalized_target_period})")
                if normalized_existing_period == normalized_target_period:
                    period_exists_globally = True
                    period_found_in_column = col
                    print(f"🌍 ✅ Target period '{normalized_target_period}' found GLOBALLY in column {col} as '{period}'")
                    break
        
        # Also check if it was added during this session
        if normalized_target_period in sheet_columns_added:
            period_exists_globally = True
            print(f"🌍 ✅ Target period '{normalized_target_period}' was added during this session")
        
        print(f"🔍 Final period detection result: period_exists_globally = {period_exists_globally}")
        
        # Store the period detection result in state for debugging and loop detection
        # This will be passed in the next iteration if we modify the state parameter
        
        # Secondary check: current table's period mapping (for completeness)
        period_exists_in_table = False
        if current_period_mapping and normalized_target_period:
            for col, period in current_period_mapping.items():
                normalized_existing_period = normalize_period_for_database(period)
                if normalized_existing_period == normalized_target_period:
                    period_exists_in_table = True
                    print(f"📋 Target period '{normalized_target_period}' found in current table column {col} as '{period}'")
                    break
        
        user_prompt = f"""
CURRENT STATE:
- User Request: {user_question}
- Processing Status: {processing_status}
- Current Table Index: {current_table_index}
- Processed Tables: {len(processed_tables)} completed
- Total Tables: {len(identified_tables)}
- Processed Table Ranges: {processed_tables}
- Target Period: {target_period}
- Normalized Target Period: {normalized_target_period}
- Target Period Exists GLOBALLY: {period_exists_globally}
- Target Period Exists in Current Table: {period_exists_in_table}
- Sheet Period Mapping: {sheet_period_mapping}
- Columns Added This Session: {sheet_columns_added}

{current_table_info}

EXCEL DATA ANALYSIS:
Look carefully at the Excel data below to see if the target period already exists as a column header:
{excel_data[:8000]}

CRITICAL DECISION LOGIC - FOLLOW EXACTLY:
1. **MANDATORY CHECK**: Period exists globally = {period_exists_globally}
2. **IF period_exists_globally = True**: USE cell_mapping_and_fill_current_table
3. **IF period_exists_globally = False**: USE modify_excel_sheet
4. If period exists globally, DO NOT use modify_excel_sheet - use cell_mapping_and_fill_current_table instead
5. Only use modify_excel_sheet if target period is completely missing from sheet_period_mapping
6. Remember: Period columns are SHEET-GLOBAL, not table-specific!
7. Current table period mapping (local): {current_period_mapping}
8. Pay attention to the table range: {current_table.get('range', 'N/A') if current_table else 'N/A'}

TASK: Determine if the target period column already exists in the current table range.
If it exists, proceed to data filling. If not, add the column first.

ANALYZE AND DECIDE NEXT ACTION:
"""
        
        print(f"🤖 LLM Decision: Analyzing current state...")
        print(f"📊 Status: {processing_status}")
        print(f"📋 Tables: {len(processed_tables)}/{len(identified_tables)} processed")
        print(f"📊 Current table index: {current_table_index}")
        if current_table:
            print(f"📋 Current table: {current_table.get('range', 'N/A')}")
        
        # Prepare input for logging
        llm_input = {
            "user_question": user_question,
            "processing_status": processing_status,
            "current_table_index": current_table_index,
            "processed_tables": processed_tables,
            "identified_tables": [table.get('range') for table in identified_tables],
            "current_table": current_table,
            "excel_data": excel_data[:500] + "..." if len(excel_data) > 500 else excel_data
        }
        
        # Make Azure OpenAI API call
        response = client.chat.completions.create(
            model=os.getenv('DEPLOYMENT_NAME'),
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt}
            ],
            temperature=0,
            max_tokens=4000,
            response_format={"type": "json_object"}
        )
        
        # Parse LLM response
        llm_response_text = response.choices[0].message.content
        reasoning_result = json.loads(llm_response_text)
        
        print(f"🎯 LLM Decision: {reasoning_result.get('tool_name', 'unknown')}")
        print(f"💭 Reasoning: {reasoning_result.get('reasoning', 'N/A')}")
        
        # Log LLM decision if logger available
        if logger:
            logger.log_llm_decision(0, llm_input, reasoning_result)
        
        return reasoning_result
        
    except Exception as e:
        print(f"❌ LLM reasoning failed: {e}")
        if logger:
            logger.log_error(0, f"LLM reasoning failed: {str(e)}")
        
        # Default fallback decision
        if processing_status == "start":
            return {
                "tool_name": "identify_table_ranges_for_modification",
                "reasoning": "Fallback: Starting with table identification",
                "parameters": {},
                "confidence": 0.5
            }
        else:
            return {
                "tool_name": "complete",
                "reasoning": "Fallback: Cannot determine next action",
                "parameters": {},
                "confidence": 0.1
            }


def request_human_approval(reasoning_result: Dict[str, Any], state: AgentState) -> Dict[str, Any]:
    """
    Request human approval before tool execution
    
    Args:
        reasoning_result (Dict[str, Any]): LLM's proposed action
        state (AgentState): Current agent state
        
    Returns:
        Dict[str, Any]: Human decision with approval and modifications
    """
    print("\n" + "="*50)
    print("🤖 AGENT DECISION REQUIRES APPROVAL")
    print("="*50)
    print(f"📊 Current Excel State: {len(state['excel_data'])} chars")
    print(f"🎯 Proposed Tool: {reasoning_result['tool_name']}")
    print(f"💭 Reasoning: {reasoning_result['reasoning']}")
    print(f"⚙️  Parameters: {json.dumps(reasoning_result['parameters'], indent=2)}")
    
    if reasoning_result['tool_name'] == 'modify_excel_sheet':
        print("⚠️  WARNING: This will modify the Excel file!")
    
    print("\nOptions:")
    print("1. ✅ Approve and proceed")
    print("2. ❌ Reject and halt")
    print("3. 🔧 Modify parameters")
    
    while True:
        choice = input("\nYour decision (1/2/3): ").strip()
        
        if choice == "1":
            return {"approved": True, "modifications": {}}
        
        elif choice == "2":
            return {"approved": False, "reason": "User rejected"}
        
        elif choice == "3":
            print("Enter parameter modifications (JSON format):")
            try:
                modifications = json.loads(input())
                return {"approved": True, "modifications": modifications}
            except json.JSONDecodeError:
                print("Invalid JSON. Please try again.")
                continue
        
        else:
            print("Invalid choice. Please enter 1, 2, or 3.")


def apply_human_modifications(reasoning_result: Dict[str, Any], human_decision: Dict[str, Any]) -> Dict[str, Any]:
    """Apply human modifications to the reasoning result"""
    modifications = human_decision.get("modifications", {})
    
    if modifications:
        # Apply parameter modifications
        reasoning_result["parameters"].update(modifications)
        reasoning_result["human_modified"] = True
        print(f"✅ Applied human modifications: {modifications}")
    
    return reasoning_result


def call_selected_tool(tool_name: str, parameters: Dict[str, Any], state: AgentState) -> None:
    """
    Dynamically call the selected tool with parameters
    
    Args:
        tool_name (str): Name of tool to call
        parameters (Dict[str, Any]): Tool parameters
        state (AgentState): Agent state to modify directly
        
    Raises:
        Exception: If tool call fails
    """
    if tool_name == "identify_table_ranges_for_modification":
        identify_table_ranges_for_modification(
            excel_data=state["excel_data"],
            user_question=state["user_question"],
            operation_type=state.get("operation_type", ""),
            target_period=state.get("target_period", ""),
            processed_tables=state.get("processed_tables", []),
            state=state
        )
    
    elif tool_name == "modify_excel_sheet":
        # Get current table info
        current_table = get_current_table(state)
        if not current_table:
            raise Exception("No current table to modify")
        
        # Ensure we use "add_column" for adding new period columns
        modification_type = parameters.get("modification_type", state.get("operation_type", "add_column"))
        if not modification_type or modification_type == "":
            modification_type = "add_column"  # Default to add_column for period additions
        
        modify_excel_sheet(
            excel_file_path=state["excel_file_path"],
            table_range=current_table["range"],
            modification_type=modification_type,
            target_period=state["target_period"],
            position=parameters.get("position", "after_last"),
            state=state,
            target_cell=parameters.get("target_cell")
        )
    
    elif tool_name == "cell_mapping_and_fill_current_table":
        # Get current table info
        current_table = get_current_table(state)
        if not current_table:
            raise Exception("No current table to fill")
        
        cell_mapping_and_fill_current_table(
            excel_data=state["excel_data"],
            table_range=current_table["range"],
            global_items=current_table.get("global_items", {}),
            target_period=state["target_period"],
            operation_type=state["operation_type"],
            state=state
        )
    
    elif tool_name == "complete":
        state["processing_status"] = "complete"
        print("✅ All processing complete!")
    
    else:
        raise Exception(f"Unknown tool: {tool_name}")


def get_current_table(state: AgentState) -> Dict[str, Any]:
    """
    Get the current table being processed
    
    Args:
        state (AgentState): Current agent state
        
    Returns:
        Dict[str, Any]: Current table info or empty dict
    """
    identified_tables = state.get("identified_tables", [])
    current_index = state.get("current_table_index", 0)
    
    if 0 <= current_index < len(identified_tables):
        return identified_tables[current_index]
    
    return {}


def all_tables_processed(state: AgentState) -> bool:
    """
    Check if all identified tables have been processed
    
    Args:
        state (AgentState): Current agent state
        
    Returns:
        bool: True if all tables processed
    """
    identified_tables = state.get("identified_tables", [])
    processed_tables = state.get("processed_tables", [])
    
    if not identified_tables:
        return False
    
    # Check if all identified table ranges are in processed list
    identified_ranges = [table.get("range", "") for table in identified_tables]
    return all(table_range in processed_tables for table_range in identified_ranges)


def detect_infinite_loop(state: AgentState, tool_name: str, max_consecutive: int = 2) -> bool:
    """
    Detect if we're stuck in an infinite loop of the same tool
    
    Args:
        state (AgentState): Current agent state
        tool_name (str): Tool being proposed
        max_consecutive (int): Max consecutive calls of same tool before flagging
        
    Returns:
        bool: True if infinite loop detected
    """
    # Initialize action history if not exists
    if "action_history" not in state:
        state["action_history"] = []
    
    # Don't add current tool to history yet - we're just checking
    current_history = state["action_history"]
    
    print(f"🔍 Loop detection: Current history: {current_history[-5:] if len(current_history) > 5 else current_history}")
    print(f"🔍 Loop detection: Proposed tool: {tool_name}")
    
    # Check for consecutive identical actions in recent history
    if len(current_history) >= max_consecutive:
        recent_actions = current_history[-max_consecutive:]
        if all(action == tool_name for action in recent_actions):
            print(f"⚠️  INFINITE LOOP DETECTED: {tool_name} called {max_consecutive} times consecutively")
            print(f"📜 Recent action history: {current_history[-5:]}")
            return True
    
    # Special case: if we see the same tool proposed multiple times recently
    if len(current_history) >= 1 and current_history[-1] == tool_name:
        consecutive_count = 1
        for i in range(len(current_history) - 2, -1, -1):
            if current_history[i] == tool_name:
                consecutive_count += 1
            else:
                break
        
        print(f"🔍 Loop detection: Found {consecutive_count} consecutive {tool_name} calls")
        
        if consecutive_count >= max_consecutive:
            print(f"⚠️  INFINITE LOOP DETECTED: {tool_name} already called {consecutive_count} times in a row")
            print(f"📜 Action history: {current_history}")
            return True
    
    # Enhanced detection: Check for modify_excel_sheet loops with same operation
    if tool_name == "modify_excel_sheet" and len(current_history) >= 2:
        modify_count = sum(1 for action in current_history[-3:] if action == "modify_excel_sheet")
        if modify_count >= 2:
            print(f"⚠️  POTENTIAL LOOP: modify_excel_sheet called {modify_count} times in last 3 iterations")
            print(f"📊 Current state - Period exists globally: {state.get('period_exists_globally', 'Unknown')}")
            print(f"📊 Current state - Sheet period mapping: {state.get('sheet_period_mapping', {})}")
            print(f"📊 Current state - Target period: {state.get('target_period', 'Unknown')}")
            print(f"📊 This indicates the LLM is not recognizing that the period already exists!")
            return True
    
    return False


def add_to_action_history(state: AgentState, tool_name: str):
    """
    Add executed tool to action history (call after successful tool execution)
    
    Args:
        state (AgentState): Current agent state
        tool_name (str): Tool that was executed
    """
    # Initialize action history if not exists
    if "action_history" not in state:
        state["action_history"] = []
    
    # Add current action to history
    state["action_history"].append(tool_name)
    
    # Keep only recent history (last 15 actions)
    state["action_history"] = state["action_history"][-15:]
    
    print(f"📜 Action history updated: {state['action_history'][-5:]}")  # Show last 5


def force_state_transition(state: AgentState, current_tool: str) -> Dict[str, Any]:
    """
    Force a state transition when infinite loop is detected
    
    Args:
        state (AgentState): Current agent state
        current_tool (str): Tool that's causing the loop
        
    Returns:
        Dict[str, Any]: Forced decision to break the loop
    """
    print(f"🔧 FORCING STATE TRANSITION to break infinite loop...")
    
    # If stuck on modify_excel_sheet, force move to cell filling
    if current_tool == "modify_excel_sheet":
        current_table = get_current_table(state)
        if current_table:
            print(f"🔧 Forcing transition from modification to cell filling for table: {current_table.get('range', 'N/A')}")
            return {
                "tool_name": "cell_mapping_and_fill_current_table",
                "reasoning": "FORCED: Breaking infinite loop - proceeding to cell filling",
                "parameters": {},
                "confidence": 0.8,
                "forced_transition": True
            }
    
    # If stuck on cell filling, force move to next table
    elif current_tool == "cell_mapping_and_fill_current_table":
        current_index = state.get("current_table_index", 0)
        identified_tables = state.get("identified_tables", [])
        
        if current_index < len(identified_tables) - 1:
            # Move to next table
            state["current_table_index"] = current_index + 1
            current_table = identified_tables[current_index]
            state["processed_tables"].append(current_table.get("range", ""))
            
            print(f"🔧 Forcing move to next table: index {current_index + 1}")
            return {
                "tool_name": "modify_excel_sheet",
                "reasoning": "FORCED: Breaking infinite loop - moving to next table",
                "parameters": {},
                "confidence": 0.8,
                "forced_transition": True
            }
        else:
            # All tables processed
            print(f"🔧 Forcing completion - all tables processed")
            return {
                "tool_name": "complete",
                "reasoning": "FORCED: Breaking infinite loop - completing process",
                "parameters": {},
                "confidence": 0.8,
                "forced_transition": True
            }
    
    # Default: force completion
    return {
        "tool_name": "complete",
        "reasoning": "FORCED: Breaking infinite loop - cannot determine safe transition",
        "parameters": {},
        "confidence": 0.5,
        "forced_transition": True
    }


def orchestrator_node(state: AgentState) -> AgentState:
    """
    Single orchestrator node that:
    1. Always parses Excel first to get current state
    2. Uses LLM reasoning to decide next action based on table-range global context
    3. Optional human intervention before tool calling (global toggle)
    4. Dynamically calls appropriate tools (tools directly modify state)
    5. Handles errors by halting entire process
    6. Continues iteration until all tables processed
    
    Args:
        state (AgentState): Current agent state
        
    Returns:
        AgentState: Updated agent state
    """
    # Use the session logger (created once per session)
    logger = state.get("session_logger")
    current_iteration = state.get("current_iteration", 0) + 1
    
    # If no logger exists, create one (should not happen in normal flow)
    if logger is None:
        logger = create_logger()
        state["session_logger"] = logger
        print("⚠️  Created emergency logger - this should not happen in normal flow")
    
    try:
        print(f"\n🔄 === Orchestrator Iteration {current_iteration} ===")
        
        # Log iteration start with comprehensive state snapshot
        logger.log_iteration_start(current_iteration, state)
        
        # Store state before modifications for change tracking
        state_before = state.copy()
        
        # === STEP 1: ALWAYS PARSE EXCEL FIRST ===
        print(f"📖 Step 1: Re-parsing Excel file...")
        current_excel_state = re_parse_excel_state(state["excel_file_path"])
        state["excel_data"] = current_excel_state["excel_data"]
        state["excel_metadata"] = current_excel_state["excel_metadata"]
        
        # Log Excel parsing with full markdown for debugging
        logger.log_excel_parsing(current_iteration, len(state["excel_data"]), state["excel_data"], full_markdown=state["excel_data"])
        print(f"✅ Parsed Excel: {len(state['excel_data'])} characters")
        
        # === STEP 2: LLM REASONING + TOOL DECISION ===
        print(f"🧠 Step 2: LLM reasoning for next action...")
        reasoning_result = llm_reasoning_and_tool_decision(
            state["excel_data"],
            state["user_question"], 
            state.get("processing_status", "start"),
            state.get("processed_tables", []),
            state.get("identified_tables", []),
            state.get("current_table_index", 0),
            state.get("sheet_period_mapping", {}),
            state.get("sheet_columns_added", []),
            logger
        )
        
        # === STEP 2.5: INFINITE LOOP DETECTION ===
        proposed_tool = reasoning_result.get("tool_name", "")
        if detect_infinite_loop(state, proposed_tool):
            print(f"🚨 INFINITE LOOP DETECTED! Forcing state transition...")
            reasoning_result = force_state_transition(state, proposed_tool)
            logger.log_error(current_iteration, f"Infinite loop detected for tool: {proposed_tool}", {
                "action_history": state.get("action_history", []),
                "forced_decision": reasoning_result
            })
        
        # Check for completion
        if reasoning_result.get("tool_name") == "complete":
            state["processing_status"] = "complete"
            print("✅ Processing complete!")
            logger.save_session_log()
            return state
        
        # === STEP 3: HUMAN-IN-THE-LOOP (OPTIONAL) ===
        if state.get("human_intervention_enabled", False):
            print(f"👤 Step 3: Requesting human approval...")
            human_decision = request_human_approval(reasoning_result, state)
            
            # Log human intervention
            logger.log_human_intervention(current_iteration, reasoning_result, human_decision)
            
            if not human_decision.get("approved", False):
                state["processing_status"] = "human_rejected"
                print("❌ Human rejected the action")
                logger.save_session_log()
                return state
            # Apply any human modifications to the reasoning_result
            reasoning_result = apply_human_modifications(reasoning_result, human_decision)
        else:
            print(f"🤖 Step 3: Skipping human intervention (disabled)")
        
        # === STEP 4: DYNAMIC TOOL CALLING ===
        print(f"🔧 Step 4: Calling tool '{reasoning_result['tool_name']}'...")
        tool_success = False
        tool_output = {}
        
        try:
            # Prepare tool input for logging
            tool_input = {
                "tool_name": reasoning_result["tool_name"],
                "parameters": reasoning_result.get("parameters", {}),
                "excel_file_path": state["excel_file_path"],
                "current_table_index": state.get("current_table_index", 0)
            }
            
            call_selected_tool(reasoning_result["tool_name"], reasoning_result.get("parameters", {}), state)
            tool_success = True
            tool_output = {"status": "success", "message": "Tool completed successfully"}
            
            # Add to action history after successful execution
            add_to_action_history(state, reasoning_result["tool_name"])
            
            print(f"✅ Tool '{reasoning_result['tool_name']}' completed successfully")
            
            # Log successful tool execution
            logger.log_tool_execution(current_iteration, reasoning_result["tool_name"], tool_input, tool_output, True)
            
        except Exception as e:
            # Halt entire process on tool failure
            error_msg = f"Tool {reasoning_result['tool_name']} failed: {str(e)}"
            state["errors"].append(error_msg)
            state["processing_status"] = "error"
            
            tool_output = {"status": "error", "error": str(e)}
            
            # Log failed tool execution
            logger.log_tool_execution(current_iteration, reasoning_result["tool_name"], tool_input, tool_output, False)
            logger.log_error(current_iteration, error_msg)
            
            print(f"❌ {error_msg}")
            logger.save_session_log()
            return state
        
        # === STEP 5: DETERMINE NEXT ITERATION ===
        if all_tables_processed(state):
            state["processing_status"] = "complete"
            print("🎉 All tables processed successfully!")
        
        state["current_iteration"] = current_iteration
        
        # Set current table for next iteration
        current_table = get_current_table(state)
        state["current_table"] = current_table
        
        # Log state changes
        logger.log_state_changes(current_iteration, state_before, state)
        
        print(f"📊 Iteration {state['current_iteration']} complete")
        print(f"📋 Total cells filled: {state.get('total_cells_filled', 0)}")
        
        # Save session log after each iteration
        logger.save_session_log()
        
        return state
        
    except Exception as e:
        error_msg = f"Orchestrator node failed: {str(e)}"
        state["errors"].append(error_msg)
        state["processing_status"] = "error"
        
        # Log orchestrator failure
        logger.log_error(current_iteration, error_msg, {"exception_type": type(e).__name__})
        logger.save_session_log()
    
        print(f"❌ {error_msg}")
    return state


def create_orchestrator_graph() -> StateGraph:
    """
    Create the LangGraph StateGraph with single orchestrator node
    
    Returns:
        StateGraph: Configured graph ready for execution
    """
    # Create the graph
    graph = StateGraph(AgentState)
    
    # Add the single orchestrator node
    graph.add_node("orchestrator", orchestrator_node)
    
    # Set entry point
    graph.set_entry_point("orchestrator")
    
    # Add conditional edges for looping
    def should_continue(state: AgentState) -> str:
        """Determine if processing should continue"""
        status = state.get("processing_status", "")
        
        if status in ["complete", "error", "human_rejected"]:
            return "end"
        else:
            return "orchestrator"
    
    graph.add_conditional_edges(
        "orchestrator",
        should_continue,
        {
            "orchestrator": "orchestrator",  # Continue processing
            "end": "__end__"                # End processing
        }
    )
    
    return graph.compile()


def run_excel_agent(excel_file_path: str, user_question: str, enable_human_intervention: bool = False) -> AgentState:
    """
    Main entry point to run the Excel agent
    
    Args:
        excel_file_path (str): Path to Excel file to process
        user_question (str): User's natural language request
        enable_human_intervention (bool): Enable human intervention mode
        
    Returns:
        AgentState: Final state after processing
    """
    print(f"🚀 Starting Excel Agent")
    print(f"📁 Original File: {excel_file_path}")
    print(f"❓ Question: {user_question}")
    print(f"👤 Human intervention: {'ENABLED' if enable_human_intervention else 'DISABLED'}")
    
    # Create a copy of the Excel file to preserve the original
    try:
        working_file_path = create_excel_copy(excel_file_path)
        print(f"📋 Working on copy: {working_file_path}")
    except Exception as e:
        print(f"❌ Failed to create Excel copy, using original: {e}")
        working_file_path = excel_file_path
    
    # Set human intervention mode
    set_human_intervention_mode(enable_human_intervention)
    
    # Create a single session logger for the entire run
    session_logger = create_logger()
    
    # Initialize state with the working copy path
    initial_state = initialize_agent_state(working_file_path, user_question)
    initial_state["session_logger"] = session_logger
    
    # Create and run the graph
    graph = create_orchestrator_graph()
    
    try:
        # Execute the graph
        final_state = graph.invoke(initial_state)
        
        # Print summary
        print(f"\n🏁 === PROCESSING COMPLETE ===")
        print(f"📁 Original File: {excel_file_path}")
        print(f"📋 Working Copy: {working_file_path}")
        print(f"📊 Final Status: {final_state.get('processing_status', 'unknown')}")
        print(f"📋 Tables Processed: {len(final_state.get('processed_tables', []))}")
        print(f"📝 Cells Filled: {final_state.get('total_cells_filled', 0)}")
        print(f"🔄 Iterations: {final_state.get('current_iteration', 0)}")
        
        if final_state.get("errors"):
            print(f"❌ Errors: {len(final_state['errors'])}")
            for error in final_state["errors"]:
                print(f"   - {error}")
        
        if final_state.get("warnings"):
            print(f"⚠️  Warnings: {len(final_state['warnings'])}")
            for warning in final_state["warnings"]:
                print(f"   - {warning}")
        
        return final_state
        
    except Exception as e:
        print(f"❌ Graph execution failed: {e}")
        print(f"📁 Original File: {excel_file_path}")
        print(f"📋 Working Copy: {working_file_path}")
        initial_state["errors"].append(f"Graph execution failed: {str(e)}")
        initial_state["processing_status"] = "error"
        return initial_state


if __name__ == "__main__":
    """
    Test the orchestrator agent with sample data
    """
    print("=== Testing Excel Agent Orchestrator ===")
    
    # Test configuration
    test_excel_file = "docs/sample_inputs/itus-banking-sample.xlsx"
    test_user_question = "fill data for Q1 25"
    test_human_intervention = False  # Disabled to test automatic loop detection and fixes
    
    # Check if test file exists
    if not os.path.exists(test_excel_file):
        print(f"❌ Test file not found: {test_excel_file}")
        print("Please provide a valid Excel file path to test the agent.")
        sys.exit(1)
    
    try:
        # Run the agent
        result_state = run_excel_agent(
            excel_file_path=test_excel_file,
            user_question=test_user_question,
            enable_human_intervention=test_human_intervention
        )
        
        print("✅ Test completed successfully!")
        
    except Exception as e:
        print(f"❌ Test failed: {e}")