# XL Fill Plugin - AI Agent Logic Documentation

## Purpose of this file:

This document provides comprehensive logic and workflow instructions for building a LangGraph-based AI agent that intelligently fills Excel financial templates. The agent uses a **single orchestrator node** with dynamic tool calling and iterative table processing, utilizing the existing xl_fill_plugin API backend.

**📄 For current API implementation details, see [Current API Implementation](./current-api-implementation.md)**

---

## Overview

The AI agent processes Excel files to intelligently fill financial templates using a **dynamic orchestrator architecture** that processes tables iteratively:

**Core Workflow**: **Single Orchestrator Node** with iterative table processing:
1. **Always Parse First** - Re-parse Excel to get current state at each iteration
2. **LLM Reasoning** - Analyze current state and decide next action using table-range wise global context
3. **Human-in-the-Loop** (Optional) - Allow human intervention before tool calling with global toggle
4. **Dynamic Tool Calling** - Call appropriate tools based on LLM reasoning
5. **State Update** - Tools directly modify state based on results
6. **Repeat** - Continue until all identified tables are processed

### Key Architectural Principles

#### **Orchestrator-Driven Processing**
- **Single Node**: One orchestrator node handles all decisions and tool calling
- **Always Parse First**: Every iteration starts by re-parsing the Excel file to get current state
- **Dynamic Tool Selection**: LLM decides which tools to call based on current Excel state
- **Sequential Table Processing**: Process one table range at a time, handle Excel modifications mid-process
- **Table-Range Global Context**: Apply global vs cell-specific logic independently for each table range
- **Global Context Persistence**: Global context per table range is identified only once and preserved
- **Error Handling**: Halt entire process on tool failures for data integrity
- **Human-in-the-Loop**: Optional human intervention before tool calling with global toggle control

#### **Spatial Mapping Patterns**
- **Quarters/Periods**: Always mapped column-wise in Excel sheets (horizontal progression)
  - Examples: Q1→Q2→Q3 moving left to right across columns
  - Can appear in any row, agent must scan entire sheet
  - Formats vary: "Q3 25", "Q1 FY26", "FY24", "CY23", etc.
  - **KEY INSIGHT**: Periods are **column-global** - all cells in column C correspond to the same period

- **Metrics**: Always mapped row-wise in Excel sheets (vertical progression)  
  - Examples: Revenue→EBITDA→Net Income moving top to bottom down rows
  - Can appear in any column, agent must scan entire sheet
  - May include embedded information of company name, entity or metric_type like "Deposit Growth % (Standalone)" or "HDFC Bank Retail Banking Market Share"

#### **Period Mapping Strategy**
- **Column-Global Context**: Each column has one period that applies to ALL data cells in that column
  - Column C = "Q4 25" (all C5, C6, C7... are Q4 25 data)
  - Column D = "Q1 26" (all D5, D6, D7... are Q1 26 data)
  - Column E = "Q2 FY26" (all E5, E6, E7... are Q2 FY26 data)
- **Period Normalization**: Convert display formats to database-compatible formats
  - Display: "Q2 26" → Database: "Q2 FY26"
  - Display: "Q4 25" → Database: "Q4 FY25"
- **Consistency Rule**: Use normalized format everywhere for exact matching

#### **Context Mapping Strategy**
- **Global Items**: Values that are 100% consistent across ALL cells in a specific table range
  - `company_name`: Company identifier (e.g., "HDFC Bank", "ICICI Bank")
  - `entity`: Subsidiary/Business entity/segment of the company (e.g., "HDFC Bank", "Retail Banking", "ICICI Life Insurance")  
  - `metric_type`: Consolidation level ("Consolidated", "Standalone", or empty string)
  - `period_mapping`: Column-to-period mapping (e.g., {"C": "Q4 FY25", "D": "Q1 FY26", "E": "Q2 FY26"})
  - **Critical**: If ANY cell in the table range has different values of company name, entity or metric_type, the item is NOT global for that table range.

- **Cell-Specific Items**: Values that vary per individual cell location
  - `metric`: Financial metric name extracted from row context and headers
  - `quarter`: Time period derived from global period_mapping based on cell's column

#### **Decision Logic for Global vs Cell-Specific**
The agent must apply this logic for each table range independently:
- **100% consistency rule**: For an item to be global, it must be identical across every single cell in the table range
- **Multi-table independence**: Different table ranges can have completely different global items
- **Fallback strategy**: When in doubt, treat as cell-specific (empty string for global items)

---

## LangGraph Agent Workflow

### Agent State Structure

The orchestrator maintains comprehensive state for iterative table processing with human-in-the-loop capabilities and Excel live monitoring:

```python
class AgentState(TypedDict):
    # === INPUT & CURRENT EXCEL STATE ===
    excel_file_path: str                      # Path to Excel file (gets modified during processing)
    sheet_name: str                           # Name of Excel sheet to work on
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
    
    # === EXCEL LIVE MONITORING & MANAGEMENT ===
    excel_manager: Any                        # ExcelManager instance for xlwings operations
    session_logger: Any                       # Single logger instance for entire session
    
    # === RESULTS & TRACKING ===
    table_processing_results: Dict[str, Any] # Results per table range
    total_cells_filled: int                   # Running count across all tables
    errors: List[str]                         # Errors encountered (halt on error)
    warnings: List[str]                       # Warnings or ambiguities
    
    # === METADATA ===
    excel_metadata: Dict[str, Any]            # Excel metadata (refreshed each iteration)
    llm_analysis: Dict[str, Any]              # LLM analysis results
```

### Orchestrator Node Architecture

#### Mermaid diagram

graph TD
    A[Excel File + User Question] --> B[Orchestrator Node]
    B --> C{LLM Decision}
    C -->|First Run| D[Tool 1: Identify Tables]
    C -->|Need Modification| E[Tool 2: Modify Excel]
    C -->|Ready to Fill| F[Tool 3: Fill Cells]
    D --> G[Update State]
    E --> G
    F --> G
    G --> H{All Tables Done?}
    H -->|No| B
    H -->|Yes| I[Complete]

#### **Core Orchestrator Loop with Human-in-the-Loop and Excel Live Monitoring**

```python
def orchestrator_node(state: AgentState) -> AgentState:
    """
    Single orchestrator node that:
    1. Always parses Excel first to get current state
    2. Excel Live Monitoring: Refreshes Excel workbook for visual inspection
    3. Uses LLM reasoning to decide next action based on table-range global context
    4. Optional human intervention before tool calling (global toggle)
    5. Dynamically calls appropriate tools (tools directly modify state)
    6. Handles errors by halting entire process
    7. Continues iteration until all tables processed
    """
    
    # === STEP 1: ALWAYS PARSE EXCEL FIRST ===
    print(f"📖 Step 1: Re-parsing Excel file...")
    
    # === STEP 1.5: EXCEL LIVE MONITORING REFRESH ===
    # Refresh Excel workbook if manager exists and is open (only if live monitoring is enabled)
    if MONITOR_EXCEL_LIVE and state.get("excel_manager") and state["excel_manager"].is_open:
        print(f"🔄 Refreshing Excel workbook for visual inspection...")
        state["excel_manager"].refresh_excel()
        state["excel_manager"].ensure_visible()
    
    current_excel_state = re_parse_excel_state(state["excel_file_path"], state.get("sheet_name"))
    state["excel_data"] = current_excel_state["excel_data"]
    state["excel_metadata"] = current_excel_state["excel_metadata"]
    
    # === STEP 2: LLM REASONING + TOOL DECISION ===
    reasoning_result = llm_reasoning_and_tool_decision(
        state["excel_data"],
        state["user_question"], 
        state.get("processing_status", "start"),
        state.get("processed_tables", []),
        state.get("identified_tables", []),
        state.get("current_table_index", 0),
        state.get("sheet_period_mapping", {}),
        state.get("sheet_columns_added", []),
        state.get("session_logger")
    )
    
    # === STEP 3: HUMAN-IN-THE-LOOP (OPTIONAL) ===
    if state.get("human_intervention_enabled", False):
        human_decision = request_human_approval(reasoning_result, state)
        if not human_decision.get("approved", False):
            state["processing_status"] = "human_rejected"
            return state
        # Apply any human modifications to the reasoning_result
        reasoning_result = apply_human_modifications(reasoning_result, human_decision)
    
    # === STEP 4: DYNAMIC TOOL CALLING ===
    try:
        call_selected_tool(reasoning_result["tool_name"], reasoning_result["parameters"], state)
    except Exception as e:
        # Halt entire process on tool failure
        state["errors"].append(f"Tool {reasoning_result['tool_name']} failed: {str(e)}")
        state["processing_status"] = "error"
        return state
    
    # === STEP 5: DETERMINE NEXT ITERATION ===
    if all_tables_processed(state):
        state["processing_status"] = "complete"
    
    state["current_iteration"] = state.get("current_iteration", 0) + 1
    return state
```

### Available Tools

The orchestrator uses three specialized tools that directly modify the state object:

#### **Tool 1: `identify_table_ranges_for_modification`**
**Purpose**: Identify table ranges and extract table-range specific global context (called once per table)

```python
def identify_table_ranges_for_modification(
    excel_data: str,
    user_question: str,
    operation_type: str,
    target_period: str,
    processed_tables: List[str],
    state: AgentState  # Tool directly modifies state
) -> None:
    """
    Identifies tables and applies the 100% consistency rule for global items per table range.
    Global context per table range is identified only once and preserved.
    
    DIRECTLY MODIFIES STATE:
    - state["identified_tables"] = List of tables with global context
    - state["operation_type"] = Determined operation type
    - state["target_period"] = Normalized target period
    - state["processing_status"] = "tables_identified"
    
    Table Structure Example:
    {
        "range": "A5:D15",
        "description": "Key Financial Metrics",
        "needs_new_column": True,
        "needs_new_rows": False,
        "modification_required": "add_column_after_D",
        "global_items": {                    # Identified once, preserved forever
            "company_name": "HDFC Bank",     # Global if 100% consistent in this range
            "entity": "HDFC Bank",           # Global if 100% consistent in this range  
            "metric_type": "",               # Empty if not 100% consistent in this range
            "period_mapping": {              # Column-wise period mapping (NEW!)
                "C": "Q4 25",                # Column C contains Q4 25 data
                "D": "Q1 26",                # Column D contains Q1 26 data  
                "E": "Q2 FY26"               # Column E contains Q2 FY26 data (normalized)
            }
        },
        "relevance_score": 0.9
    }
    """
```

#### **Tool 2: `modify_excel_sheet`**
**Purpose**: Physically modify the Excel file to accommodate new data with live monitoring support

```python
def modify_excel_sheet(
    excel_file_path: str,
    table_range: str,
    modification_type: str,  # "add_column", "add_row", "insert_column"
    target_period: str,
    position: str,
    state: AgentState  # Tool directly modifies state
) -> None:
    """
    Modifies Excel file and updates state with changes.
    Supports both openpyxl (for structural changes) and xlwings (for live monitoring).
    On failure, raises exception to halt entire process.
    
    DIRECTLY MODIFIES STATE:
    - Updates table range in state["identified_tables"] 
    - Updates sheet_period_mapping with new period columns
    - Updates sheet_columns_added tracking
    - state["processing_status"] = "excel_modified"
    - Forces re-parse on next iteration (excel_data will be refreshed)
    
    MODIFIES EXCEL FILE:
    - Uses openpyxl for structural changes (adding columns/rows)
    - Adds columns/rows as needed
    - Sets appropriate headers
    - Preserves formatting
    
    EXCEL LIVE MONITORING INTEGRATION:
    - After openpyxl modifications, forces Excel to reload from disk if excel_manager exists
    - Calls excel_manager.reload_from_disk() to display structural changes
    - Maintains real-time visual feedback during modifications
    """
```

#### **Tool 3: `cell_mapping_and_fill_current_table`**
**Purpose**: Apply global context and cell-specific mapping for the current table range with real-time Excel updates

```python
def cell_mapping_and_fill_current_table(
    excel_data: str,
    table_range: str,
    global_items: Dict[str, Any],  # Table-range specific global context (preserved)
    target_period: str,
    operation_type: str,
    state: AgentState  # Tool directly modifies state
) -> None:
    """
    Uses preserved table-range global context + cell-specific extraction to map and fill cells.
    Calls xl_fill_plugin API and updates Excel file with results using hybrid update strategy.
    
    DIRECTLY MODIFIES STATE:
    - state["table_processing_results"][table_range] = Results
    - state["total_cells_filled"] += cells_filled
    - state["processed_tables"].append(table_range)
    - state["current_table_index"] += 1
    - state["processing_status"] = "next_table" or "complete"
    
    HYBRID EXCEL UPDATE STRATEGY:
    - **Real-time Updates**: Uses xlwings when excel_manager.is_open for immediate visibility
      * Cell values updated instantly
      * Hyperlinks added and styled (blue, underlined)
      * Comments with API matching details preserved
      * Full support for all original openpyxl features in real-time
    - **File-based Updates**: Falls back to openpyxl when xlwings not available
      * Traditional file modification approach
      * Same features but requires file reload for visibility
    
    EXCEL FILE MODIFICATIONS:
    - Fills cells with API results and hyperlinks
    - Adds standardized comments: "company_name | entity | metric_type | metric | time_period | document_year"
    - Different cell values for different statuses:
      * "filled": Actual numerical/text value from database
      * "no_data": "N/A" (no matches found)
      * "llm_no_match": "NO MATCH" (LLM rejected all matches when ALLOW_NO_MATCH=true)
      * "error": "ERROR" (API or processing error)
    - Preserves formatting and adds source traceability
    
    Cell Mapping Structure:
    {
        "E5": {
            "company_name": "HDFC Bank",      # From preserved global_items
            "entity": "HDFC Bank",            # From preserved global_items
            "metric_type": "",                # From preserved global_items
            "metric": "Loan Growth %",        # Cell-specific from row context
            "quarter": "Q2 FY26",            # From global period_mapping["E"] (normalized)
            "source_info": {...}
        }
    }
    """
```

### Human-in-the-Loop Implementation

#### **Global Toggle Control**
```python
# In agent.py
ENABLE_HUMAN_INTERVENTION = os.getenv('ENABLE_HUMAN_INTERVENTION', 'false').lower() == 'true'  # Global toggle via environment

def set_human_intervention_mode(enabled: bool):
    """Global function to enable/disable human intervention"""
    global ENABLE_HUMAN_INTERVENTION
    ENABLE_HUMAN_INTERVENTION = enabled

def initialize_agent_state(excel_file_path: str, user_question: str) -> AgentState:
    """Initialize state with human intervention setting"""
    return {
        "excel_file_path": excel_file_path,
        "user_question": user_question,
        "human_intervention_enabled": ENABLE_HUMAN_INTERVENTION,
        "pending_human_approval": False,
        "human_decision": {},
        # ... other state fields
    }
```

#### **Human Approval Process**
```python
def request_human_approval(reasoning_result: Dict[str, Any], state: AgentState) -> Dict[str, Any]:
    """
    Request human approval before tool execution
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
```

### LLM Reasoning and Tool Decision

#### **Orchestrator Decision Process**
```python
def llm_reasoning_and_tool_decision(
    excel_data: str,
    user_question: str, 
    processing_status: str,
    processed_tables: List[str],
    identified_tables: List[Dict]
) -> Dict[str, Any]:
    """
    LLM analyzes current state and decides next action with table-range context
    """
    
    system_prompt = """
    You are an Excel modification orchestrator. Analyze the current state and decide the next action.
    
    AVAILABLE TOOLS:
    1. identify_table_ranges_for_modification - When no tables identified yet (first run)
    2. modify_excel_sheet - When current table needs structural changes (add column/row)
    3. cell_mapping_and_fill_current_table - When current table ready for data filling
    
    DECISION LOGIC:
    - If processing_status == "start": Use identify_table_ranges_for_modification
    - If identified_tables[current_index] needs modification: Use modify_excel_sheet
    - If identified_tables[current_index] ready for filling: Use cell_mapping_and_fill_current_table
    - If all tables processed: Return "complete"
    
    GLOBAL CONTEXT RULES:
    - Global context per table range is identified once and preserved
    - Never re-evaluate global context after Excel modifications
    - Sequential table processing - one table at a time
    
    PERIOD MAPPING RULES:
    - Check period_mapping in current table's global_items for existing periods
    - Normalize period formats: "Q2 26" equals "Q2 FY26" 
    - If target period exists in period_mapping, proceed to cell filling
    - Only add new columns if target period genuinely missing from table
    
    Return JSON: {
        "tool_name": "tool_to_call",
        "reasoning": "why this tool is needed",
        "parameters": {"param1": "value1"},
        "confidence": 0.9
    }
    """
    
    user_prompt = f"""
    CURRENT STATE:
    - User Request: {user_question}
    - Processing Status: {processing_status}
    - Processed Tables: {len(processed_tables)} completed
    - Identified Tables: {len(identified_tables)} total
    - Current Excel Preview: {excel_data[:1500]}...
    
    ANALYZE AND DECIDE NEXT ACTION:
    """
    
    # Make LLM call and return structured decision
    # Implementation details for Azure OpenAI call
```

### Excel Live Monitoring and Real-time Updates

The agent incorporates comprehensive Excel live monitoring capabilities using xlwings for macOS compatibility and real-time visual feedback during processing.

#### **Architecture Overview**

The system uses a **hybrid approach** combining the strengths of both xlwings and openpyxl:

- **xlwings**: Real-time updates, visual inspection, and immediate feedback
- **openpyxl**: Structural modifications (adding columns/rows) and complex operations

#### **Environment Variable Configuration**

```python
# Global toggles for Excel live monitoring
MONITOR_EXCEL_LIVE = os.getenv('MONITOR_EXCEL_LIVE', 'false').lower() == 'true'
ENABLE_HUMAN_INTERVENTION = os.getenv('ENABLE_HUMAN_INTERVENTION', 'false').lower() == 'true'
GET_BEST_5 = os.getenv('GET_BEST_5', 'false').lower() == 'true'
ALLOW_NO_MATCH = os.getenv('ALLOW_NO_MATCH', 'false').lower() == 'true'
```

**MONITOR_EXCEL_LIVE** Configuration:
- **Purpose**: Controls whether Excel is opened visually and refreshed during agent execution
- **Default**: `false` (agent runs without opening Excel visually for faster execution)
- **When `true`**: Excel opens visibly, refreshes during iterations, provides real-time visual feedback
- **When `false`**: Agent runs in background mode without visual Excel display

#### **ExcelManager Class Integration**

```python
class ExcelManager:
    """
    Manages Excel file operations using xlwings for macOS compatibility.
    Enhanced for M1/M4 Mac compatibility with real-time updates.
    """
    
    def __init__(self):
        self.app = None
        self.workbook = None
        self.file_path = None
        self.is_open = False
    
    def open_excel_file(self, file_path: str, display: bool = True) -> bool:
        """Open Excel file for visual inspection and refreshing"""
        # Connects to or starts Excel application
        # Makes Excel visible for inspection
        # Opens the workbook and brings Excel to front
    
    def refresh_excel(self, sheet_name: str = None) -> bool:
        """
        Refresh data connections and calculations in Excel workbook.
        macOS M1/M4 compatible with enhanced processing time and display refresh.
        """
        # Platform-specific refresh methods (macOS vs Windows)
        # Extended processing time for M1/M4 chips
        # Force screen updates and recalculation
        # Save workbook after refresh
    
    def reload_from_disk(self) -> bool:
        """Force reload Excel workbook from disk to pick up openpyxl changes"""
        # Workaround for openpyxl/xlwings compatibility
        # Closes current workbook and reopens from disk
        # Maintains visual inspection capability
    
    def update_cell(self, sheet_name: str, cell_ref: str, value: any) -> bool:
        """Update a single cell using xlwings for real-time updates"""
    
    def update_cells_batch(self, sheet_name: str, cell_updates: dict) -> bool:
        """Update multiple cells in batch using xlwings"""
```

#### **Agent Initialization with Live Monitoring**

```python
def run_excel_agent(excel_file_path: str, user_question: str, sheet_name: str = "Main") -> AgentState:
    """
    Main entry point with integrated Excel live monitoring
    """
    
    # Create Excel manager for visual inspection (only if live monitoring enabled)
    excel_manager = None
    if MONITOR_EXCEL_LIVE:
        excel_manager = ExcelManager()
        
        # Open Excel file for visual inspection at startup
        print(f"📊 Opening Excel file for visual inspection...")
        excel_opened = excel_manager.open_excel_file(working_file_path, display=True)
        
        if excel_opened:
            print(f"✅ Excel file is now open for visual inspection during agent execution")
        else:
            print(f"⚠️ Could not open Excel file for inspection, continuing without visual display")
    else:
        print(f"📊 Excel live monitoring disabled - running without visual display")
    
    # Initialize state with excel_manager
    initial_state = initialize_agent_state(working_file_path, user_question, sheet_name)
    initial_state["excel_manager"] = excel_manager
```

#### **Real-time Update Strategy**

**Problem Solved**: Incompatibility between openpyxl (file-based) and xlwings (in-memory) operations.

**Solution 1: Cell Data Updates (Primary)**
```python
def _update_excel_with_xlwings(excel_manager, processed_results, sheet_name, cell_mappings):
    """Update Excel using xlwings for real-time updates with full feature support"""
    
    sheet = excel_manager.workbook.sheets[sheet_name]
    
    for cell_ref, result in processed_results.items():
        cell_range = sheet.range(cell_ref)
        
        if result["status"] == "filled":
            # Set value with automatic type conversion
            cell_range.value = result["value"]
            
            # Add hyperlink if available (real-time styling)
            if result.get("source_url"):
                cell_range.add_hyperlink(result["source_url"])
                cell_range.font.color = (0, 0, 255)  # Blue
                cell_range.font.underline = True
            
            # Add comment with API matching details (real-time)
            if cell_mappings and cell_ref in cell_mappings:
                comment_text = create_standardized_comment(mapping, result)
                cell_range.note.text = comment_text
```

**Solution 2: Structural Changes (Secondary)**
```python
# In modify_excel_sheet.py - after openpyxl saves structural changes
if excel_manager and excel_manager.is_open:
    excel_manager.reload_from_disk()  # Force reload to display structural changes
```

#### **Smart Update Detection**

```python
def _update_excel_with_results(excel_file_path, processed_results, cell_mappings, excel_manager, sheet_name):
    """Smart update strategy selection based on available tools"""
    
    # Check if we're using xlwings for real-time updates
    use_xlwings = excel_manager and excel_manager.is_open
    
    if use_xlwings:
        print("🔄 Using xlwings for real-time Excel updates...")
        _update_excel_with_xlwings(excel_manager, processed_results, sheet_name, cell_mappings)
        return  # Exit early since xlwings handles everything
    else:
        print("📝 Using openpyxl for Excel updates...")
        # Traditional openpyxl approach with file-based updates
```

#### **Enhanced macOS M1/M4 Compatibility**

The refresh functionality includes specific optimizations for Apple Silicon:

```python
def refresh_excel(self, sheet_name: str = None) -> bool:
    """Enhanced macOS M1/M4 compatible refresh"""
    
    if platform.system() == "Darwin":  # macOS
        # Extended processing time for M1/M4 chips
        time.sleep(2)  # Increased from 1 to 2 seconds
        
        # Force display refresh for M1/M4
        self.app.screen_updating = False
        time.sleep(0.3)  # Brief pause for M1/M4 processing
        self.app.screen_updating = True
        
        # Additional save verification for macOS
        if save_successful:
            time.sleep(1)  # Brief pause after save
            print("✅ Save operation completed")
```

#### **Visual Inspection Benefits**

When `MONITOR_EXCEL_LIVE=true`:

1. **Real-time Changes**: See modifications as they happen during agent execution
2. **Debug Assistance**: Spot issues immediately with visual feedback
3. **Data Validation**: Verify results visually during processing
4. **Process Transparency**: Understand what the agent is doing step-by-step
5. **Interactive Troubleshooting**: Excel remains open for inspection if errors occur

#### **Performance Considerations**

- **xlwings updates**: Faster for individual cell updates, immediate visibility
- **xlwings hyperlinks/comments**: Slight overhead but still faster than file reload
- **openpyxl + reload**: Necessary for structural changes, brief delay during reload
- **Hybrid approach**: Optimal balance between performance and functionality
- **Feature completeness**: All original features (values, hyperlinks, comments) preserved in real-time

#### **Troubleshooting Integration**

```python
# Excel cleanup - keep workbook open for inspection unless there were errors
if MONITOR_EXCEL_LIVE and excel_manager and excel_manager.is_open:
    final_status = final_state.get('processing_status', 'unknown')
    if final_status == "complete":
        print(f"📊 Excel workbook remains open for final inspection")
        excel_manager.ensure_visible()
    else:
        print(f"⚠️ Process ended with status '{final_status}' - keeping Excel open for troubleshooting")
        excel_manager.ensure_visible()
```

### Orchestrator Workflow Summary

#### **Sequential Table Processing Flow**
1. **Iteration 1**: `identify_table_ranges_for_modification` → Global context identified once per table
2. **Iteration 2**: `modify_excel_sheet` → Add required columns/rows if needed  
3. **Iteration 3**: `cell_mapping_and_fill_current_table` → Fill data for current table
4. **Iteration 4+**: Repeat steps 2-3 for next table in sequence
5. **Final**: All tables processed → `processing_status = "complete"`

#### **Error Handling Strategy**
- **Tool Failure**: Halt entire process immediately
- **Human Rejection**: Stop processing and return control to user  
- **Excel Corruption**: Halt with detailed error message
- **API Failures**: Halt with specific error details

#### **State Management**
- **Always Parse First**: Excel state refreshed every iteration
- **Global Context Preservation**: Identified once per table, never re-evaluated
- **Sequential Processing**: One table at a time, no parallel processing
- **Direct State Modification**: Tools modify state directly, no return values

---

## Data Mapping Logic

### Global Items vs Cell-Specific Decision Tree

The agent must make precise decisions about what context items should be treated as global vs cell-specific for each table range. This is one of the most critical aspects of the entire system.

#### **Global Items Definition and Rules**

**Global Items** are context values that are **100% consistent across ALL cells** in a specific table range:

1. **`company_name`**: Company identifier
   - **Examples**: "HDFC Bank", "ICICI Bank", "Reliance Industries Ltd"
   - **Global When**: Single company mentioned in table headers, all data rows relate to same company
   - **Not Global When**: Multiple companies in different rows of the same table
   - **Edge Cases**: 
     - Company mentioned in sheet title but multiple subsidiaries in rows → Not global
     - Company abbreviations vs full names → Use most complete version if consistent

2. **`entity`**: Business entity/segment/subsidiary
   - **Examples**: 
     - Same as company: "HDFC Bank" (entity) = "HDFC Bank" (company)
     - Business segment: "Retail Banking", "Investment Banking"
     - Subsidiary: "HDFC Life Insurance", "ICICI Prudential"
     - Geographic: "HDFC Bank India", "Citibank Singapore"
   - **Global When**: All metrics in table relate to same business unit
   - **Not Global When**: Table covers multiple business units or subsidiaries
   - **Default Rule**: If no separate entity mentioned, entity = company_name

3. **`metric_type`**: Consolidation level
   - **Values**: "Consolidated", "Standalone", or "" (empty string)
   - **Global When**: 
     - Section header like "Standalone Results" covers entire table
     - All individual metrics consistently have same suffix (e.g., all end with "(Consolidated)")
   - **Not Global When**: 
     - Mix of consolidated and standalone metrics in same table
     - Some metrics have suffix, others don't
     - No clear indication anywhere
   - **Default**: Empty string when unclear

#### **Cell-Specific Items Rules**

**Cell-Specific Items** vary for each individual cell location:

1. **`metric`**: Financial metric name
   - **Extraction Strategy**: 
     - Primary: From row labels (usually leftmost columns)
     - Secondary: From section headers combined with row context
     - Tertiary: From complex row structures
   - **Cleaning Rules**: 
     - **DO NOT** remove parenthetical content: "Net Income (NII)" stays as-is
     - **DO NOT** remove percentage signs: "Growth %" stays as-is
     - **DO NOT** remove metric type indicators: "Revenue (Standalone)" stays as-is
   - **Concatenation Rules**: 
     - If table has section headers, combine: "Deposit Growth (Standalone)" + "HDFC Bank" = "HDFC Bank Deposit Growth (Standalone)"
     - If multiple companies per row: "HDFC Bank Retail Banking" + "Deposit Growth" = "HDFC Bank Retail Banking Deposit Growth"

2. **`quarter`**: Time period 
   - **Extraction Strategy**: From column headers, scanning upward from target cell
   - **Multiple Formats**: Must handle various display formats but normalize for database
   - **Display vs Database**: Maintain original format in Excel, convert for API calls

#### **Complex Decision Scenarios**

**Scenario 1: Mixed Metric Types in Same Table**
```markdown
| A | B | C | D |
|---|---|---|---|
| Revenue (Standalone) | 100 | 110 | 120 |
| EBITDA (Consolidated) | 80 | 85 | 90 |
```
**Decision**: metric_type = "" (not global), embedded in individual metrics

**Scenario 2: Multiple Entities, Same Company**
```markdown
| A | B | C | D |
|---|---|---|---|
| HDFC Retail Banking | 100 | 110 | 120 |
| HDFC Investment Banking | 80 | 85 | 90 |
```
**Decision**: company_name = "HDFC Bank" (global), entity = "" (not global)

**Scenario 3: Section Headers with Mixed Content**
```markdown
| A | B | C | D |
|---|---|---|---|
| Standalone Results | | | |
| HDFC Bank Revenue | 100 | 110 | 120 |
| ICICI Bank Revenue | 80 | 85 | 90 |
```
**Decision**: metric_type = "Standalone" (global), company_name = "" (not global)

### Period/Quarter Normalization

**Critical Requirement**: The database requires EXACT format matching. Period normalization is essential for successful data retrieval.

#### **Comprehensive Format Mapping**

**Input Format → Database Format:**

**Quarter Patterns:**
- "Q3 25" → "Q3 FY25"
- "Q2FY26" → "Q2 FY26"
- "Q1 FY 25" → "Q1 FY25"
- "3Q25" → "Q3 FY25"
- "Q4'24" → "Q4 FY24"

**Financial Year Patterns:**
- "FY25" → "FY25" (already correct)
- "FY 25" → "FY25"
- "2025" → "FY25" (if in financial context)
- "25" → "FY25" (if in quarter context)

**Calendar Year Patterns:**
- "2024" → "CY2024"
- "CY24" → "CY2024"
- "CY 24" → "CY2024"

#### **Enhanced Normalization Algorithm**

```python
def normalize_period_for_database(display_period):
    """
    Convert any period format to database-compatible format
    Handles all common Excel display formats
    """
    if not display_period:
        return display_period
    
    # Remove common separators and spaces
    cleaned = re.sub(r'[^\w]', '', display_period.upper())
    
    # Quarter patterns (highest priority)
    if re.match(r'Q\d+\d{2}', cleaned):
        # Q325 -> Q3 FY25, Q226 -> Q2 FY26
        quarter = cleaned[0:2]
        year = cleaned[2:]
        return f"{quarter} FY{year}"
    
    # Quarter with space patterns (Q2 26, Q4 25)
    elif re.match(r'Q\d+\s*\d{2}', display_period.upper()):
        parts = re.findall(r'Q(\d+)\s*(\d{2})', display_period.upper())
        if parts:
            quarter, year = parts[0]
            return f"Q{quarter} FY{year}"
    
    # Quarter with FY patterns (Q2FY26, Q1 FY25)
    elif re.match(r'Q\d+\s*FY\s*\d{2}', display_period.upper()):
        parts = re.findall(r'Q(\d+)\s*FY\s*(\d{2})', display_period.upper())
        if parts:
            quarter, year = parts[0]
            return f"Q{quarter} FY{year}"
    
    # Financial year patterns (FY25, FY 25)
    elif re.match(r'FY\s*\d{2}', display_period.upper()):
        year = re.findall(r'FY\s*(\d{2})', display_period.upper())[0]
        return f"FY{year}"
    
    # Calendar year patterns (2024, CY24, CY 24)
    elif re.match(r'(CY\s*)?\d{4}', display_period.upper()):
        year = re.findall(r'(\d{4})', display_period)[0]
        return f"CY{year}"
    
    # Default: return as-is with warning
    print(f"⚠️  Unrecognized period format: '{display_period}' - using as-is")
    return display_period

def maintain_display_format(original_period, normalized_period):
    """
    Keep original format for Excel display while using normalized for database
    """
    return {
        "display_format": original_period,    # "Q3 25" for Excel
        "database_format": normalized_period  # "Q3 FY25" for API
    }
```

#### **Period Validation**

**Before API Calls:**
```python
def validate_period_format(period):
    """
    Ensure period format will work with database
    """
    valid_patterns = [
        r"Q[1-4] FY\d{2}",    # Q1 FY25
        r"FY\d{2}",           # FY25
        r"CY\d{4}"            # CY2024
    ]
    
    for pattern in valid_patterns:
        if re.match(pattern, period):
            return True, period
    
    return False, f"Invalid period format: {period}"
```

### Context Concatenation Strategy

When global items are empty (not consistent across table), the agent must intelligently combine context elements into the `metric` field for database search.

#### **Concatenation Rules**

**Rule 1: Company + Metric (when company not global)**
```python
# Table has multiple companies
company_from_row = "HDFC Bank"
metric_from_row = "Deposit Growth %"
final_metric = f"{company_from_row} {metric_from_row}"
# Result: "HDFC Bank Deposit Growth %"
```

**Rule 2: Entity + Metric (when entity not global)**
```python
# Table has multiple business segments
entity_from_row = "Retail Banking"
metric_from_row = "Revenue"
final_metric = f"{entity_from_row} {metric_from_row}"
# Result: "Retail Banking Revenue"
```

**Rule 3: Section Header + Row Context**
```python
# Table has section headers
section_header = "Deposit Growth (Standalone)"
company_from_row = "HDFC Bank"
final_metric = f"{company_from_row} {section_header}"
# Result: "HDFC Bank Deposit Growth (Standalone)"
```

**Rule 4: Full Context Combination**
```python
# Complex table with multiple context levels
company = "ICICI Bank"
entity = "Life Insurance Division" 
section = "Premium Growth"
final_metric = f"{company} {entity} {section}"
# Result: "ICICI Bank Life Insurance Division Premium Growth"
```

#### **Search Optimization**

The concatenated strings work effectively because:
1. **Hybrid Search**: BM25 + Vector search handles various keyword combinations
2. **Semantic Understanding**: Vector embeddings understand financial terminology relationships
3. **Keyword Matching**: BM25 finds exact company/metric matches
4. **Flexible Matching**: System tolerates minor variations in naming

---

## Excel Structure Examples

### **Case 1: Company-Centric Structure** (HDFC Bank Example)

**Original Excel Structure:**
```markdown
|   A   |    B     |   C   |   D   |   E   |   F   |
|-------|----------|-------|-------|-------|-------|
|   1   |          | Q3 25 | Q4 25 | Q1 26 | [NEW] |
|   2   | HDFC Bank|       |       |       |       |
|   3   |          |       |       |       |       |
|   4   | Key Financial Metrics |   |       |       |
|   5   | Loan Growth %    | 3%    | 6%    | 7%    | [FILL]|
|   6   | Deposit Growth % | 16%   | 14%   | 16%   | [FILL]|
|   7   | NII Growth %     | 12%   | 15%   | 18%   | [FILL]|
```

**Agent Step-by-Step Analysis:**

1. **Table Identification:**
   - **Visual boundaries**: Empty rows at 3, column spacing patterns
   - **Detected range**: A1:F7 (includes headers and data)
   - **Data range**: A5:F7 (actual metric data)
   - **Header analysis**: Row 1 contains periods, Row 2 contains company, Row 4 contains section header

2. **Global Items Analysis:**
   ```python
   # Company Analysis
   company_scan_results = {
       "B2": "HDFC Bank",           # Found in header area
       "data_rows": []              # No company names in data rows (A5-A7)
   }
   # Decision: company_name = "HDFC Bank" (global, confidence: 0.95)
   
   # Entity Analysis  
   entity_scan_results = {
       "separate_entity_found": False,  # No business segments mentioned
       "default_to_company": True
   }
   # Decision: entity = "HDFC Bank" (global, confidence: 0.90)
   
   # Metric Type Analysis
   metric_type_scan_results = {
       "section_headers": ["Key Financial Metrics"],  # No consolidated/standalone
       "embedded_in_metrics": [],                     # No parenthetical indicators
       "separate_columns": []                         # No metric type columns
   }
   # Decision: metric_type = "" (global, confidence: 1.0)
   ```

3. **Final Cell Mappings:**
   ```python
   cell_mappings = {
       "F5": {
           "company_name": "HDFC Bank",      # From global items
           "entity": "HDFC Bank",            # From global items  
           "metric_type": "",                # From global items
           "metric": "Loan Growth %",        # Extracted from A5
           "quarter": "Q2 FY26"             # Normalized target period
       },
       "F6": {
           "company_name": "HDFC Bank",
           "entity": "HDFC Bank", 
           "metric_type": "",
           "metric": "Deposit Growth %",     # Extracted from A6
           "quarter": "Q2 FY26"
       }
   }
   ```

### **Case 2: Multi-Company Comparison Structure**

**Original Excel Structure:**
```markdown
|   A   |    B     |   C   |   D   |   E   |   F   |
|-------|----------|-------|-------|-------|-------|
|   1   |          | Q3 25 | Q4 25 | Q1 26 | [NEW] |
|   2   |          |       |       |       |       |
|   3   | Banking Sector Comparison |    |       |       |
|   4   | Deposit Growth % |       |       |       |       |
|   5   | HDFC Bank| 3%    | 6%    | 7%    | [FILL]|  
|   6   | ICICI Bank| 16%  | 14%   | 16%   | [FILL]|
|   7   | Axis Bank| 12%   | 16%   | 19%   | [FILL]|
```

**Agent Step-by-Step Analysis:**

1. **Table Identification:**
   - **Section detected**: A4:F7 (Deposit Growth comparison)
   - **Section header**: "Deposit Growth %" in A4
   - **Multiple companies**: Each row represents different bank

2. **Global Items Analysis:**
   ```python
   # Company Analysis
   companies_found = ["HDFC Bank", "ICICI Bank", "Axis Bank"]  # Multiple companies
   # Decision: company_name = "" (not global - multiple values)
   
   # Entity Analysis (follows company decision)
   # Decision: entity = "" (not global - follows company)
   
   # Metric Type Analysis
   section_header = "Deposit Growth %"  # No consolidation indicator
   # Decision: metric_type = "" (not global - no clear indication)
   ```

3. **Cell-Specific Context Extraction:**
   ```python
   # Strategy: Combine section header with row company
   for cell in ["F5", "F6", "F7"]:
       row_company = extract_from_row(cell)     # "HDFC Bank", "ICICI Bank", etc.
       section_metric = "Deposit Growth %"      # From A4 header
       combined_metric = f"{row_company} {section_metric}"
   ```

4. **Final Cell Mappings:**
   ```python
   cell_mappings = {
       "F5": {
           "company_name": "",                       # Not global - empty
           "entity": "",                             # Not global - empty
           "metric_type": "",                        # Not global - empty  
           "metric": "HDFC Bank Deposit Growth %",   # Combined context
           "quarter": "Q2 FY26"
       },
       "F6": {
           "company_name": "",
           "entity": "",
           "metric_type": "",
           "metric": "ICICI Bank Deposit Growth %",  # Combined context
           "quarter": "Q2 FY26"
       },
       "F7": {
           "company_name": "",
           "entity": "",
           "metric_type": "",
           "metric": "Axis Bank Deposit Growth %",   # Combined context
           "quarter": "Q2 FY26"
       }
   }
   ```

### **Case 3: Complex Metric Type Segregation**

**Original Excel Structure:**
```markdown
|   A   |    B     |   C   |   D   |   E   |   F   |
|-------|----------|-------|-------|-------|-------|
|   1   | HDFC Bank Results | Q3 25 | Q4 25 | Q1 26 | [NEW] |
|   2   |          |       |       |       |       |
|   3   | Standalone Results |     |       |       |       |
|   4   | Revenue  | 1000  | 1100  | 1200  | [FILL]|
|   5   | EBITDA   | 800   | 850   | 900   | [FILL]|
|   6   |          |       |       |       |       |
|   7   | Consolidated Results |   |       |       |       |
|   8   | Revenue  | 1500  | 1600  | 1700  | [FILL]|
|   9   | EBITDA   | 1200  | 1250  | 1300  | [FILL]|
```

**Agent Step-by-Step Analysis:**

1. **Table Identification:**
   - **Separate sections detected**: 
     - Section 1: A3:F5 (Standalone Results)
     - Section 2: A7:F9 (Consolidated Results)
   - **Section headers**: "Standalone Results" and "Consolidated Results"
   - **Same company**: "HDFC Bank" in A1 (sheet title)

2. **Global Items Analysis by Section:**
   
   **Section 1 Analysis (A3:F5):**
   ```python
   # Company Analysis
   company_in_sheet_title = "HDFC Bank"
   companies_in_section = []  # No companies in data rows
   # Decision: company_name = "HDFC Bank" (global, from sheet title)
   
   # Entity Analysis
   entity_references = []  # No separate entities mentioned
   # Decision: entity = "HDFC Bank" (global, same as company)
   
   # Metric Type Analysis
   section_header = "Standalone Results"  # Clear indicator
   # Decision: metric_type = "Standalone" (global, confidence: 0.95)
   ```

   **Section 2 Analysis (A7:F9):**
   ```python
   # Same company and entity analysis as Section 1
   # Metric Type Analysis
   section_header = "Consolidated Results"  # Clear indicator
   # Decision: metric_type = "Consolidated" (global, confidence: 0.95)
   ```

3. **Cell Mappings by Section:**
   ```python
   # Section 1 (Standalone) Cell Mappings
   standalone_mappings = {
       "F4": {
           "company_name": "HDFC Bank",     # Global from sheet title
           "entity": "HDFC Bank",           # Global (same as company)
           "metric_type": "Standalone",     # Global from section header
           "metric": "Revenue",             # From A4
           "quarter": "Q2 FY26"
       },
       "F5": {
           "company_name": "HDFC Bank",
           "entity": "HDFC Bank", 
           "metric_type": "Standalone",
           "metric": "EBITDA",              # From A5
           "quarter": "Q2 FY26"
       }
   }
   
   # Section 2 (Consolidated) Cell Mappings
   consolidated_mappings = {
       "F8": {
           "company_name": "HDFC Bank",     # Global from sheet title
           "entity": "HDFC Bank",           # Global (same as company)
           "metric_type": "Consolidated",   # Global from section header  
           "metric": "Revenue",             # From A8
           "quarter": "Q2 FY26"
       },
       "F9": {
           "company_name": "HDFC Bank",
           "entity": "HDFC Bank",
           "metric_type": "Consolidated",
           "metric": "EBITDA",              # From A9
           "quarter": "Q2 FY26"
       }
   }
   ```

---

## Agent Implementation Notes

### Error Handling Strategy

The agent should handle errors through intelligent prompting, graceful degradation, and user interaction. Error handling is built into each node of the LangGraph workflow.

#### **Node-Specific Error Handling**

**Node 1 (Sheet Analysis) Errors:**
```python
error_scenarios = {
    "file_corruption": {
        "symptoms": "Excel parsing fails, image analysis fails",
        "response": "Request user to resave file in standard Excel format",
        "fallback": "Use text-only analysis if image unavailable"
    },
    "unsupported_format": {
        "symptoms": "Non-standard Excel features, complex macros",
        "response": "Extract basic data, warn about lost formatting",
        "fallback": "Focus on text content extraction"
    },
    "empty_sheet": {
        "symptoms": "No data detected in analysis",
        "response": "Ask user to verify correct sheet/tab selected",
        "fallback": "Prompt for manual data range specification"
    }
}
```

**Node 2 (Intent Interpretation) Errors:**
```python
error_scenarios = {
    "ambiguous_request": {
        "symptoms": "Multiple possible interpretations",
        "response": "Present options to user: 'Did you mean add column or update existing?'",
        "fallback": "Default to most conservative interpretation (update existing)"
    },
    "impossible_request": {
        "symptoms": "Request conflicts with sheet structure",
        "response": "Explain constraints: 'Cannot add Q1 26 - already exists in column E'",
        "fallback": "Suggest alternative: 'Would you like to update Q1 26 instead?'"
    },
    "unclear_period": {
        "symptoms": "Period format not recognizable",
        "response": "Ask for clarification: 'Did you mean Q2 FY26 or Q2 FY25?'",
        "fallback": "Use pattern matching with confidence score"
    }
}
```

**Node 3 (Table Identification) Errors:**
```python
error_scenarios = {
    "no_tables_detected": {
        "symptoms": "Complex layout, no clear boundaries",
        "response": "Ask user to select table range manually",
        "fallback": "Treat entire sheet as single table"
    },
    "multiple_valid_tables": {
        "symptoms": "Several tables could match user intent",
        "response": "Show table previews, ask user to choose",
        "fallback": "Process all relevant tables"
    },
    "fragmented_data": {
        "symptoms": "Data scattered across multiple areas",
        "response": "Highlight detected areas, ask for confirmation",
        "fallback": "Process largest coherent section"
    }
}
```

**Node 4 (Global Items Extraction) Errors:**
```python
error_scenarios = {
    "conflicting_global_items": {
        "symptoms": "Same item appears differently across table",
        "response": "Show conflict: 'Found both HDFC Bank and HDFC Bank Ltd'",
        "resolution": "Ask user which to use, or default to most complete version"
    },
    "unclear_entity_distinction": {
        "symptoms": "Cannot distinguish company from entity",
        "response": "Present analysis: 'Treating as same entity - confirm?'",
        "fallback": "Default entity = company_name"
    },
    "mixed_metric_types": {
        "symptoms": "Some metrics consolidated, others standalone",
        "response": "Explain decision: 'Treating metric_type as cell-specific due to mixed types'",
        "fallback": "Set metric_type = '' for all cells"
    }
}
```

**Node 6 (API Orchestration) Errors:**
```python
error_scenarios = {
    "no_data_found": {
        "symptoms": "API returns empty results for cell",
        "response": "Report: 'No data found for HDFC Bank Loan Growth Q2 FY26'",
        "fallback": "Leave cell empty with comment explaining missing data"
    },
    "api_timeout": {
        "symptoms": "API calls exceed timeout",
        "response": "Retry with longer timeout, process remaining cells",
        "fallback": "Skip failed cells, report in summary"
    },
    "partial_failures": {
        "symptoms": "Some cells succeed, others fail",
        "response": "Continue processing, compile failure report",
        "fallback": "Fill successful cells, highlight failures"
    }
}
```

#### **User Communication Patterns**

**Proactive Communication:**
```python
communication_triggers = {
    "low_confidence_decisions": "Confidence < 0.7 on global items",
    "data_quality_issues": "Missing data for >20% of requested cells", 
    "structural_ambiguity": "Multiple valid table interpretations",
    "format_inconsistencies": "Mixed period formats or naming conventions"
}

response_templates = {
    "decision_explanation": "I identified {company_name} as global because it appears consistently across all {num_rows} data rows. However, I'm {confidence}% confident. Would you like to confirm?",
    "progress_update": "Processing {current_cell}/{total_cells} cells. Found data for {success_count}, missing data for {missing_count}.",
    "completion_summary": "Successfully filled {filled_count} cells. {missing_count} cells had no available data. {error_count} cells encountered errors."
}
```

### Tool Requirements

The agent needs access to these specialized tools with specific capabilities:

#### **1. Excel Parser Tool**
```python
tool_capabilities = {
    "input": "Excel file (.xlsx, .xls)",
    "output": "Structured markdown with preserved coordinates",
    "features": [
        "Convert Excel to markdown table format",
        "Preserve exact cell coordinates (A1, B2, etc.)",
        "Maintain empty cells and spacing",
        "Extract basic formatting hints (bold, merged cells)",
        "Handle multiple sheets (if needed)"
    ],
    "error_handling": [
        "Graceful degradation for unsupported features",
        "Fallback to text-only extraction",
        "Clear error messages for corrupted files"
    ]
}
```

#### **2. LLM Analysis Tool**
```python
tool_capabilities = {
    "input": "Excel markdown + user question + metadata",
    "output": "Structured analysis JSON with operation classification, table identification, and global context",
    "features": [
        "Intent classification and period normalization",
        "Table boundary detection and structure analysis",
        "Global context extraction with 100% consistency rule",
        "Feasibility validation and error detection",
        "Confidence scoring for all decisions"
    ],
    "llm_requirements": [
        "Large context window (>32k tokens) for complete Excel data",
        "Structured JSON output format enforcement",
        "Financial domain knowledge for metric recognition",
        "Excel structure understanding and pattern recognition"
    ]
}
```

#### **3. DataFrame Processor Tool** 
```python
tool_capabilities = {
    "input": "LLM analysis output + pandas DataFrame + operation parameters",
    "output": "Modified DataFrame + cell coordinate mappings",
    "operations": [
        "add_column: Insert new period column with headers",
        "update_existing: Identify existing cells to update", 
        "add_metrics: Insert new metric rows",
        "restructure: Handle complex layout changes"
    ],
    "coordinate_tracking": [
        "Return exact Excel cell references for all changes",
        "Maintain original formatting positions",
        "Track header vs data cell distinctions"
    ]
}
```

#### **4. API Client Tool**
```python
tool_capabilities = {
    "endpoints": [
        "/xl-fill-values: Existing context-based processing",
        "/get-values: Direct parameter lookup",
        "Discovery endpoints: /companies, /metrics, /periods"
    ],
    "features": [
        "Parallel request processing with concurrency control",
        "Automatic retry logic with exponential backoff",
        "Response validation and error classification",
        "Metadata extraction and confidence scoring"
    ],
    "error_handling": [
        "Graceful handling of partial failures",
        "Detailed error reporting per cell",
        "Alternative search suggestions for failed lookups"
    ]
}
```

### Human-in-the-Loop Integration

The agent should support human intervention at strategic decision points while maintaining workflow efficiency.

#### **Intervention Points**

**1. Table Range Selection:**
```python
intervention_scenarios = {
    "confidence_threshold": "Auto-detection confidence < 0.8",
    "user_preference": "User explicitly requests manual selection",
    "conflict_resolution": "Multiple equally valid interpretations exist"
}

intervention_interface = {
    "preview_mode": "Show detected tables with colored highlights",
    "selection_tool": "Allow user to draw/select ranges directly",
    "confirmation_prompt": "Present clear before/after comparison"
}
```

**2. Global Items Confirmation:**
```python
confirmation_triggers = {
    "ambiguous_entities": "Entity != company_name with confidence < 0.9",
    "mixed_contexts": "Borderline decision between global vs cell-specific",
    "user_domain_knowledge": "Financial expert might have better context"
}

confirmation_interface = {
    "decision_summary": "Show agent's reasoning with confidence scores",
    "alternative_options": "Present other viable interpretations", 
    "impact_preview": "Show how decision affects cell mappings"
}
```

**3. Data Validation:**
```python
validation_checkpoints = {
    "pre_api_calls": "Show cell mappings before database queries",
    "post_processing": "Display filled values before final application",
    "error_resolution": "Present alternatives for failed lookups"
}

validation_interface = {
    "cell_mapping_preview": "Table showing all cell->context mappings",
    "data_preview": "Show actual values that will be filled",
    "confidence_indicators": "Visual indicators for data quality"
}
```

**4. Error Resolution:**
```python
error_resolution_strategies = {
    "suggestion_mode": "Present multiple resolution options",
    "learning_mode": "Learn from user corrections for future improvement",
    "fallback_mode": "Graceful degradation when automation fails"
}

resolution_interface = {
    "error_explanation": "Clear description of what went wrong",
    "suggested_fixes": "Ranked list of potential solutions",
    "manual_override": "Allow direct specification of problematic fields"
}
```

### Performance Optimization

#### **Caching Strategy**
```python
caching_opportunities = {
    "period_normalization": "Cache normalization patterns per session",
    "company_entity_mapping": "Cache recognized company-entity relationships", 
    "table_structure_patterns": "Learn common Excel template patterns",
    "api_response_caching": "Cache recent database lookups"
}
```

#### **Parallel Processing**
```python
parallelization_points = {
    "api_calls": "Process multiple cells simultaneously (max 10 concurrent)",
    "table_analysis": "Analyze multiple tables in parallel",
    "validation_checks": "Parallel validation of different aspects",
    "image_processing": "Async image analysis while processing text"
}
```

### Quality Assurance

#### **Validation Framework**
```python
validation_layers = {
    "input_validation": "Verify Excel file integrity and format",
    "logic_validation": "Check global items decisions for consistency",
    "output_validation": "Verify all cell mappings have required fields",
    "api_validation": "Validate API responses before application",
    "final_validation": "Check Excel file integrity after updates"
}
```

#### **Confidence Scoring**
```python
confidence_factors = {
    "table_detection": "Visual + structural consistency",
    "global_items": "Percentage consistency across table range",
    "period_normalization": "Pattern match confidence",
    "api_responses": "Database match scores and metadata quality"
}

confidence_thresholds = {
    "auto_proceed": "> 0.9 - High confidence, proceed automatically",
    "user_confirmation": "0.7-0.9 - Medium confidence, ask user confirmation", 
    "manual_intervention": "< 0.7 - Low confidence, require manual review"
}
```

This comprehensive logic documentation provides everything needed to build a sophisticated LangGraph agent that can intelligently analyze Excel sheets, make context-aware decisions, and reliably fill financial data while maintaining transparency and user control throughout the process.

---

## Key Configuration Parameters

| Parameter | Value | Purpose |
|-----------|-------|---------|
| `rrf_k` | 60 | RRF ranking parameter for hybrid search |
| `vector_threshold` | 0.67 | Minimum vector similarity score |
| `hybrid_threshold` | 0.033 | Minimum hybrid score (calculated) |
| `limit` | 5-10 | Max results per search type |
| `max_concurrent` | 10 | Parallel processing limit |

---

## Database Schema Reference

### Required Fields in COMPANY_WISE_METRICS_DATA3:
- `company_name`: Company identifier
- `entity`: Business entity/subsidiary  
- `metric_type`: "Consolidated" or "Standalone"
- `metric`: Financial metric name
- `time_period`: Quarter (e.g., "Q2 FY26")
- `value`: Numeric metric value
- `value_in`: Units (e.g., "INR Cr")
- `units`: Value type (e.g., "Absolute", "Percentage")
- `document_year`: Source document year
- `source_url`: Reference document link
- `metric_embeddings`: Vector embeddings for semantic search
- `sparse_vector`: BM25 keyword index

---

## LLM Configuration and Service Toggle

The agent supports two LLM services with seamless switching via environment variables:

### Available LLM Services

1. **Azure OpenAI** (Default)
   - Uses `AzureOpenAI` client
   - Model specified by `DEPLOYMENT_NAME` environment variable
   - Supports JSON response format enforcement
   - Configuration via environment variables:
     - `AZURE_DEPLOYMENT`: Azure deployment endpoint or name
     - `OPENAI_API_KEY`: Azure OpenAI API key
     - `OPENAI_API_VERSION`: API version (e.g., "2024-02-01")
     - `DEPLOYMENT_NAME`: Model deployment name

2. **Ollama Turbo Service** (Optional)
   - Uses `ollama.Client` with Turbo service
   - Model: `gpt-oss:20b`
   - Automatic Turbo mode when using hosted service
   - Configuration via environment variables:
     - `USE_OLLAMA`: Set to `"true"` to enable Ollama (default: `"false"`)
     - `OLLAMA_HOST`: Ollama service endpoint (default: `"https://ollama.com"`)
     - `OLLAMA_API_KEY`: Ollama API key for Turbo service

### LLM Service Toggle

The agent automatically detects which service to use based on the `USE_OLLAMA` environment variable:

```python
# Toggle between services
USE_OLLAMA=false  # Uses Azure OpenAI (default)
USE_OLLAMA=true   # Uses Ollama gpt-oss:20b Turbo
```

### Unified LLM Interface

All LLM calls go through the `get_llm_response()` function which:
- Handles service switching transparently
- Manages JSON response format requirements
- Provides consistent error handling
- Supports temperature and token limit controls

### JSON Response Handling

The agent requires JSON responses from LLMs for tool decisions:
- **Azure OpenAI**: Uses `response_format={"type": "json_object"}`
- **Ollama**: Enhances system prompt with JSON enforcement instructions

---

This documentation provides the complete logic for building an AI agent that can intelligently understand Excel structures, extract context, and fill financial data using the xl_fill_plugin backend API.

---

## LLM-Enhanced Best Match Selection

### Overview

The agent now supports intelligent best match selection using LLM reasoning when multiple top results are available from the xl_fill_plugin API. This feature enhances accuracy by allowing the AI to evaluate and choose the most contextually appropriate match from multiple options.

### Configuration

#### Environment Variables
```python
GET_BEST_5 = os.getenv('GET_BEST_5', 'false').lower() == 'true'
ALLOW_NO_MATCH = os.getenv('ALLOW_NO_MATCH', 'false').lower() == 'true'
ENABLE_HUMAN_INTERVENTION = os.getenv('ENABLE_HUMAN_INTERVENTION', 'false').lower() == 'true'
```

**GET_BEST_5**:
- **Purpose**: Controls whether to request top 5 results from xl_fill_plugin API and use LLM-based selection
- **Default**: `false` (maintains existing behavior)

**ALLOW_NO_MATCH**:
- **Purpose**: Allows LLM to reject all top 5 matches when fundamentally incompatible with target context
- **Default**: `false` (LLM must select best available match - legacy behavior)
- **Dependency**: Only active when `GET_BEST_5=true`

**ENABLE_HUMAN_INTERVENTION**:
- **Purpose**: Enables human-in-the-loop approval before each tool execution
- **Default**: `false` (automated execution without human approval)
- **Impact**: When `true`, agent pauses before each tool call for user approval/modification

**Location**: `/excel-agent/.env` file

### Enhanced API Call Flow

#### Standard Flow (GET_BEST_5=false)
```python
payload = {
    "company_name": "HDFC Bank",
    "entity": "HDFC Bank", 
    "metric": "net interest margin",
    "metric_type": "",
    "quarter": "Q1 FY25",
    "get_best_5": False  # Default behavior
}
# Returns: Single best match in matched_values
```

#### Enhanced Flow (GET_BEST_5=true)
```python
payload = {
    "company_name": "HDFC Bank",
    "entity": "HDFC Bank",
    "metric": "net interest margin", 
    "metric_type": "",
    "quarter": "Q1 FY25",
    "get_best_5": True  # Request top 5 results
}
# Returns: Single best match in matched_values + top 5 in top_5_matches
```

### LLM-Based Selection Process

#### Selection Function
```python
def llm_select_best_match(top_5_matches: List[Dict[str, Any]], cell_mapping: Dict[str, Any]) -> Dict[str, Any]:
    """
    Use LLM to select the best match from top 5 results based on context compatibility
    
    Args:
        top_5_matches: Top 5 matches from API (ranked by hybrid search)
        cell_mapping: Original cell mapping context for comparison
        
    Returns:
        Selected best match from the top 5
    """
```

#### LLM Selection Criteria
The LLM evaluates matches based on:
1. **Company Name Match**: Exact or close match to target company
2. **Entity Consistency**: Entity alignment with context
3. **Metric Type Compatibility**: Standalone/Consolidated match
4. **Metric Relevance**: Semantic similarity to target metric
5. **Time Period Alignment**: Exact or closest time period match

#### No Match Selection (ALLOW_NO_MATCH=true)
When enabled, the LLM can select `"no_match"` instead of picking from ranks 1-5:

**Selection Options**:
- **Ranks 1-5**: Select best available match from top options
- **"no_match"**: Reject all matches when fundamentally incompatible

**No Match Criteria** (LLM rejects all matches when):
- **Company Mismatch**: Target is "HDFC Bank" but all matches are different companies like "Reliance Industries"
- **Metric Type Incompatibility**: Target is "standalone" but all matches are "consolidated" 
- **Time Period Gap**: Target is "Q1 FY26" but all matches are from years like 2020
- **Entity Mismatch**: Target is "Retail Banking" but all matches are different like "Insurance"
- **Metric Mismatch**: Target is "Net Interest Margin" but all matches are different like "Net Interest Income"

**Conservative Approach**: LLM only uses "no_match" for fundamental incompatibilities, not minor variations

#### LLM Prompt Structure

**Standard Prompt (ALLOW_NO_MATCH=false)**:
```python
system_prompt = """
You are an expert financial data analyst. Select the BEST match from the provided options 
based on contextual relevance and accuracy.

SELECTION CRITERIA (in order of importance):
1. Company name exact/close match
2. Entity consistency  
3. Metric type compatibility
4. Metric semantic relevance
5. Time period alignment

Return JSON with:
{
    "selected_rank": <1-5>,
    "reasoning": "Explain why this match is best",
    "confidence": <0.1-1.0>
}
"""
```

**Enhanced Prompt (ALLOW_NO_MATCH=true)**:
```python
system_prompt = """
You are an expert financial data analyst. Select the BEST match from the provided options 
based on contextual relevance and accuracy.

SELECTION CRITERIA (in order of importance):
1. Company name exact/close match
2. Entity consistency  
3. Metric type compatibility
4. Metric semantic relevance
5. Time period alignment

NO MATCH OPTION:
- If NONE of the 5 matches are suitable, you can select "no_match"
- Use this when:
  * Company names are completely different
  * Metric types are incompatible 
  * Time periods are too far apart
  * Entity/metric mismatches are too significant
- Be selective - only use "no_match" for fundamental incompatibilities

Return JSON with:
{
    "selected_rank": <1-5 or "no_match">,
    "reasoning": "Explain selection or why no match is suitable",
    "confidence": <0.1-1.0>
}
"""
```

**User Prompt Template** (same for both modes):
```python
user_prompt = f"""
TARGET CONTEXT:
Company: {cell_mapping.get('company_name')}
Entity: {cell_mapping.get('entity')}
Metric: {cell_mapping.get('metric')}
Metric Type: {cell_mapping.get('metric_type')}
Time Period: {cell_mapping.get('quarter')}

TOP 5 OPTIONS:
{formatted_options}
"""
```

### Response Processing Enhancement

#### Updated Cell Processing Logic
```python
# Enhanced response processing in cell_mapping_and_fill_current_table
if api_result["status"] == "success":
    api_data = api_result["data"]
    matched_values = api_data.get("matched_values", {})
    top_5_matches = api_data.get("top_5_matches", [])
    
    # Check for LLM-enhanced selection opportunity
    if GET_BEST_5 and top_5_matches and len(top_5_matches) > 1:
        # Use LLM to select best match from top 5
        best_match = llm_select_best_match(top_5_matches, cell_mapping_context)
        
        # Check if LLM selected "no_match" (only possible when ALLOW_NO_MATCH=True)
        if best_match.get("no_match_selected", False):
            processed_results[cell_ref] = {
                "value": "",
                "status": "llm_no_match",
                "reason": "LLM determined no suitable match from available options",
                "llm_selected": True,
                "llm_reasoning": best_match.get("llm_reasoning", ""),
                "llm_confidence": best_match.get("llm_confidence", 0),
                "total_alternatives": len(top_5_matches)
            }
            cells_failed += 1
            continue
        
        # Use LLM-selected match for final value
        value = best_match.get("value", "")
        source_data = best_match
    else:
        # Standard processing: use first (best) match from matched_values
        best_match = list(matched_values.values())[0]
        value = best_match.get("value", "")
        source_data = best_match
```

#### Excel File Updates

**Cell Values for Different Statuses**:
- **"filled"**: Actual numerical/text value from database
- **"no_data"**: `"N/A"` (no matches found in database)
- **"llm_no_match"**: `"NO MATCH"` (LLM rejected all available matches)
- **"error"**: `"ERROR"` (API or processing error)

**Excel Comments** (pipe-separated format):
```python
# Standard filled cell comment
"HDFC Bank | HDFC Bank | Standalone | Net Interest Margin | Q1 FY25 | 2024"

# LLM no match comment (includes reasoning)
"LLM NO MATCH (5 alternatives checked): Company mismatch - target HDFC Bank but all matches are Reliance | HDFC Bank | Retail Banking | Net Interest Margin | Q1 FY25 | N/A"
```

### Benefits

1. **Improved Accuracy**: LLM reasoning considers full context beyond just similarity scores
2. **Semantic Understanding**: Better handling of company name variations and metric synonyms
3. **Contextual Relevance**: Considers entity type, metric type, and time period holistically
4. **Quality Control**: Can reject unsuitable matches when `ALLOW_NO_MATCH=true`
5. **Data Integrity**: Prevents insertion of fundamentally incompatible data
6. **Backward Compatibility**: Existing behavior preserved when both flags are false
7. **Transparent Selection**: LLM provides reasoning for all decisions including rejections

### Usage Patterns

#### Configuration Options

**Legacy Mode (default)**:
```bash
# In /excel-agent/.env
GET_BEST_5=false
ALLOW_NO_MATCH=false
ENABLE_HUMAN_INTERVENTION=false
# OR simply omit all variables
```
- Uses first match from API
- No LLM evaluation
- Fully automated execution
- Fastest performance

**Enhanced Selection Mode**:
```bash
# In /excel-agent/.env
GET_BEST_5=true
ALLOW_NO_MATCH=false
ENABLE_HUMAN_INTERVENTION=false
```
- LLM selects from top 5 matches
- LLM must choose one of the 5 options
- Improved accuracy, legacy fallback behavior
- Automated execution

**Maximum Quality Mode**:
```bash
# In /excel-agent/.env
GET_BEST_5=true
ALLOW_NO_MATCH=true
ENABLE_HUMAN_INTERVENTION=false
```
- LLM selects from top 5 OR rejects all
- Highest data quality control
- May result in more empty cells
- Automated execution

**Human-in-the-Loop Mode** (any combination + human oversight):
```bash
# In /excel-agent/.env
GET_BEST_5=true
ALLOW_NO_MATCH=true
ENABLE_HUMAN_INTERVENTION=true
```
- Maximum quality with human oversight
- Agent pauses before each tool execution
- User can approve, reject, or modify actions
- Slowest but most controlled execution

#### Monitor Selection Process
```python
# LLM selection provides detailed logging:
print(f"🤖 LLM analyzing {len(top_5_matches)} top matches for best selection...")
print(f"🧠 LLM selected: {selected_rank} with confidence {confidence:.2f}")
print(f"💭 Reasoning: {reasoning}")

# No match selection (when ALLOW_NO_MATCH=true):
print("🚫 LLM determined no suitable match from top 5 results")
print(f"🚫 Cell {cell_ref}: LLM determined no suitable match from top 5 results")
```

#### Performance Considerations
- **Additional API Call**: Each cell with multiple matches requires LLM evaluation
- **Enhanced Accuracy**: Trade-off between speed and precision
- **Configurable**: Can be disabled for high-speed scenarios

---

=======LIST OF 10 RECENT CHANGES========
1. **MAJOR FEATURE**: Integrated Excel Live Monitoring and Real-time Updates - agent now supports visual Excel inspection and real-time updates using xlwings/openpyxl hybrid approach. Includes MONITOR_EXCEL_LIVE environment toggle, ExcelManager class for macOS M1/M4 compatibility, real-time cell updates with hyperlinks and comments, smart update detection, and visual troubleshooting capabilities. Solves openpyxl/xlwings incompatibility with dual-mode operation.
2. **ENHANCED**: Updated agent state structure to include excel_manager, session_logger, and sheet-global context tracking (sheet_period_mapping, sheet_columns_added, period_exists_globally) for comprehensive state management during live monitoring and real-time updates.
3. **ENHANCED**: Modified orchestrator loop to include Excel workbook refresh at each iteration when live monitoring enabled, providing visual feedback and real-time display updates during agent execution.
4. **ENHANCED**: Updated tool documentation for modify_excel_sheet and cell_mapping_and_fill_current_table to include live monitoring integration, real-time update strategies, and hybrid xlwings/openpyxl approach with detailed status handling.
5. **CONFIGURATION**: Moved ENABLE_HUMAN_INTERVENTION to environment variables - human-in-the-loop mode can now be configured via .env file with ENABLE_HUMAN_INTERVENTION=true/false instead of hardcoded value. Provides consistent configuration approach with other agent toggles and allows runtime control without code modification.
6. **MAJOR FEATURE**: Implemented ALLOW_NO_MATCH toggle for LLM match selection - agent can now reject all top 5 results when fundamentally incompatible with target context (company mismatch, metric type incompatibility, time period gaps, entity/metric mismatches). Includes conditional prompt generation, defensive fallback logic, new "llm_no_match" status with Excel cell value "NO MATCH", and detailed LLM reasoning in comments. Fully backward compatible with default ALLOW_NO_MATCH=false.
7. **MAJOR FEATURE**: Implemented LLM-Enhanced Best Match Selection with GET_BEST_5 toggle - agent now requests top 5 results from xl_fill_plugin API and uses LLM reasoning to select the most contextually appropriate match based on company name, entity, metric type, metric relevance, and time period alignment. Includes temperature=0 for consistent selection and detailed logging for transparency.
8. **ADDED**: Integrated Ollama LLM service support with USE_OLLAMA toggle - agent now supports both Azure OpenAI and Ollama (gpt-oss:20b Turbo) with unified LLM interface, automatic JSON response handling, and seamless switching via environment variables.
9. **ENHANCED**: Updated cell comment format to standardized structure: `company_name | entity | metric_type | metric | time_period | document_year` - comments now extract values from API response `matched_values` and preserve cell hyperlinks for data traceability.
10. **CRITICAL FIX**: Fixed contradictory LLM decision logic in system and user prompts - was causing agent to skip table identification step and jump straight to modification, resulting in "No current table to modify" errors. Now properly enforces processing_status == "start" → identify_table_ranges_for_modification sequence.
