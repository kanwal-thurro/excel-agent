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

The orchestrator maintains comprehensive state for iterative table processing with human-in-the-loop capabilities:

```python
class AgentState(TypedDict):
    # === INPUT & CURRENT EXCEL STATE ===
    excel_file_path: str                      # Path to Excel file (gets modified during processing)
    excel_data: str                           # Current parsed Excel state (refreshed each iteration)
    user_question: str                        # Original user request
    
    # === ORCHESTRATOR CONTROL ===
    identified_tables: List[Dict[str, Any]]   # All tables with global context (identified once)
    processed_tables: List[str]               # Table ranges already completed
    current_table_index: int                  # Which table we're currently processing
    current_table: Dict[str, Any]             # Current table being worked on
    
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

#### **Core Orchestrator Loop with Human-in-the-Loop**

```python
def orchestrator_node(state: AgentState) -> AgentState:
    """
    Single orchestrator node that:
    1. Always parses Excel first to get current state
    2. Uses LLM reasoning to decide next action based on table-range global context
    3. Optional human intervention before tool calling (global toggle)
    4. Dynamically calls appropriate tools (tools directly modify state)
    5. Handles errors by halting entire process
    6. Continues iteration until all tables processed
    """
    
    # === STEP 1: ALWAYS PARSE EXCEL FIRST ===
    current_excel_state = re_parse_excel_state(state["excel_file_path"])
    state["excel_data"] = current_excel_state["excel_data"]
    state["excel_metadata"] = current_excel_state["excel_metadata"]
    
    # === STEP 2: LLM REASONING + TOOL DECISION ===
    reasoning_result = llm_reasoning_and_tool_decision(
        state["excel_data"],
        state["user_question"], 
        state.get("processing_status", "start"),
        state.get("processed_tables", []),
        state.get("identified_tables", [])
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
**Purpose**: Physically modify the Excel file to accommodate new data

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
    On failure, raises exception to halt entire process.
    
    DIRECTLY MODIFIES STATE:
    - Updates table range in state["identified_tables"] 
    - state["processing_status"] = "excel_modified"
    - Forces re-parse on next iteration (excel_data will be refreshed)
    
    MODIFIES EXCEL FILE:
    - Adds columns/rows as needed
    - Sets appropriate headers
    - Preserves formatting
    """
```

#### **Tool 3: `cell_mapping_and_fill_current_table`**
**Purpose**: Apply global context and cell-specific mapping for the current table range only

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
    Calls xl_fill_plugin API and updates Excel file with results.
    
    DIRECTLY MODIFIES STATE:
    - state["table_processing_results"][table_range] = Results
    - state["total_cells_filled"] += cells_filled
    - state["processed_tables"].append(table_range)
    - state["current_table_index"] += 1
    - state["processing_status"] = "next_table" or "complete"
    
    MODIFIES EXCEL FILE:
    - Fills cells with API results
    - Preserves formatting
    
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
ENABLE_HUMAN_INTERVENTION = False  # Global toggle at script level

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

This documentation provides the complete logic for building an AI agent that can intelligently understand Excel structures, extract context, and fill financial data using the xl_fill_plugin backend API.

=======LIST OF 10 RECENT CHANGES========
1. **CRITICAL FIX**: Fixed period extraction logic in `llm_reasoning_and_tool_decision()` - was extracting only "25" instead of "Q1 25" from user questions, causing infinite loops because target period never matched added columns.
2. **FIX**: Fixed human intervention toggle in `set_human_intervention_mode()` - was always setting to False regardless of parameter value.
3. **ENHANCED**: Improved infinite loop detection with better debugging and enhanced detection for modify_excel_sheet loops.
4. **ADDED**: Enhanced period detection debugging with comprehensive logging to track period matching process.
5. **ADDED**: New state field `period_exists_globally` to track period detection results for debugging.
6. **ENHANCED**: Added pattern-based period extraction using regex to handle various period formats (Q1 25, Q1 FY25, FY25, CY2024).
7. **IMPROVED**: Better consecutive tool call detection with enhanced logging and state information.
8. **ADDED**: Specific detection for modify_excel_sheet loops that occur when LLM doesn't recognize existing periods.
9. **FIXED**: Test configuration now defaults to disabled human intervention for automated testing.
10. **ENHANCED**: Added comprehensive debugging output for period normalization and matching process to help troubleshoot future issues.
11. **ADDED**: Excel file copying functionality - agent now creates timestamped copies to preserve original sample files before processing.