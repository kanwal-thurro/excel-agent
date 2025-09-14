"""
LLM Prompt Engineering Module

This module contains ALL prompt templates and prompt generation functions 
for the Excel AI Agent's LLM-based operations, organized by functionality.

Key Sections:
1. Table Analysis Prompts (for identify_table_ranges_for_modification)
2. Orchestrator Decision Prompts (for agent.py orchestrator)  
3. Match Selection Prompts (for cell_mapping_and_fill_current_table)
4. Human Interaction Prompts
5. Utility Functions and Examples
"""

# =============================================================================
# 1. TABLE ANALYSIS PROMPTS (Used by identify_table_ranges_for_modification)
# =============================================================================

def create_llm_analysis_system_prompt():
    """
    System prompt for comprehensive Excel table analysis and structure identification.
    Used by identify_table_ranges_for_modification tool.
    
    Returns:
        str: Complete system prompt for LLM table analysis
    """
    return """You are an expert financial Excel analyst specializing in data pipeline automation. Your role is to systematically analyze Excel financial templates for automated data filling operations.

CONTEXT & INTENT:
You are part of an AI system that fills Excel financial templates with database values. Your analysis determines:
1. WHAT operation the user wants (add new data vs update existing)
2. WHERE in the Excel structure data should be placed (precise cell ranges)  
3. HOW to contextualize database queries (global vs cell-specific parameters)

Your analysis feeds into downstream nodes that:
- Generate precise cell mappings for database API calls
- Apply 100% consistency rules for global context
- Fill only DATA rows (never section headers or titles)

SYSTEMATIC APPROACH:
You must be EXHAUSTIVE and SYSTEMATIC - scan every row, find every data section, exclude all headers.
The downstream pipeline depends on complete, accurate boundary detection.

CORE RESPONSIBILITIES:
1. Classify user operation intent with high accuracy
2. Systematically scan ENTIRE sheet to find ALL data sections  
3. Exclude section headers from data ranges (headers are for context only)
4. Extract global context items using 100% consistency rule
5. Normalize periods to database-compatible formats
6. Validate feasibility and detect potential issues

CRITICAL RULES:
- Data ranges contain ONLY rows with fillable metric values
- Section headers are context sources, NOT data targets
- Global items must be 100% consistent across ALL cells in table range
- Period normalization: Q3 25 → Q3 FY25, 2024 → CY2024, FY 25 → FY25
- Scan entire sheet - do NOT stop after finding 2-3 sections
- If any ambiguity exists, flag as warning and provide reasoning
- Always return valid JSON following the exact structure specified
- Be conservative: when in doubt, mark items as not global

OUTPUT FORMAT:
Return a structured JSON object with four main sections:
- operation_analysis: Intent classification and period normalization
- table_analysis: Complete section identification with proper boundaries
- table_global_contexts: Global items extraction with per-table consistency validation
- validation: Feasibility check and warnings

Focus on completeness, accuracy and systematic analysis. Your output drives automated data filling."""


def create_llm_analysis_user_prompt(excel_markdown, user_question, metadata):
    """
    User prompt template for Excel table analysis with complete instructions.
    Used by identify_table_ranges_for_modification tool.
    
    Args:
        excel_markdown (str): Complete Excel data in markdown format
        user_question (str): User's natural language request
        metadata (dict): Excel metadata (dimensions, merged ranges, etc.)
        
    Returns:
        str: Complete user prompt with data and analysis requirements
    """
    return f"""Analyze the following Excel data and user request:

USER REQUEST: "{user_question}"

EXCEL DATA:
{excel_markdown}

METADATA:
Sheet: {metadata.get('sheet_info', {}).get('name', 'Unknown')}
Dimensions: {metadata.get('sheet_info', {}).get('rows', 0)} rows × {metadata.get('sheet_info', {}).get('cols', 0)} columns
Merged Ranges: {metadata.get('sheet_info', {}).get('merged_ranges', [])}

REQUIRED ANALYSIS:

1. OPERATION CLASSIFICATION:
   - Classify as: "add_column", "update_existing", or "add_metrics"
   - Extract target period and normalize (Q3 25 → Q3 FY25, etc.)
   - Assess confidence and provide reasoning

2. TABLE IDENTIFICATION:
   - Scan the ENTIRE sheet systematically to find ALL data sections
   - Use empty rows as section separators (consecutive empty rows indicate new sections)
   - EXCLUDE section headers from data ranges (e.g., "Key Financial Metrics" is header, not data)
   - Data ranges should contain ONLY rows with actual metric data
   - Look for patterns: [Section Header] → [Data Rows] → [Empty Row] → [Next Section]
   - Identify header rows, data rows, period columns, metric columns for each section
   - Score relevance of each section to user request

3. GLOBAL CONTEXT EXTRACTION (100% CONSISTENCY RULE FOR EACH TABLE):
   - FOR EACH IDENTIFIED TABLE SEPARATELY, analyze:
   - company_name: Is there ONE company consistent across ALL data cells in THIS table?
   - entity: Is there ONE entity/business unit across ALL data cells in THIS table?  
   - metric_type: Is "Consolidated"/"Standalone" consistent across ALL data cells in THIS table?
   - CRITICAL: Each table has independent global context - different tables can have different global items
   - If ANY variation exists within a table, mark as not global (empty string) for THAT table
   - IMPORTANT: Use the EXACT table ranges (e.g., "A5:D15", "A18:D27") as keys in table_global_contexts

4. VALIDATION:
   - Can this request be fulfilled with the current data structure?
   - Are there any ambiguities or potential issues?
   - Provide specific warnings or suggestions

CRITICAL TABLE DETECTION EXAMPLES:
❌ WRONG: "A4:D15" (includes section header "Key Financial Metrics" in A4)
✅ CORRECT: "A5:D15" (excludes header, includes only data rows)

❌ WRONG: Missing sections after row 27
✅ CORRECT: Scan entire sheet, find ALL sections separated by empty rows

SYSTEMATIC SCANNING APPROACH:
1. Start from row 1, scan to the last row with data
2. Identify pattern: [Section Header] → [Data Rows] → [Empty Row Separator]
3. For each section found:
   - EXCLUDE the section header row from data range
   - INCLUDE only rows with actual metric/data values
   - STOP at empty row (section separator)
4. Continue scanning after each empty row for additional sections
5. Do NOT stop after finding just 2-3 sections - scan the ENTIRE sheet

EXAMPLE PATTERN RECOGNITION:
- Row 4: "Key Financial Metrics" → SECTION HEADER (exclude from range)
- Rows 5-15: Actual metric data → DATA RANGE "A5:D15"
- Row 16: Empty → SECTION SEPARATOR  
- Row 17: "Key Operational Ratios" → SECTION HEADER (exclude from range)
- Rows 18-27: Actual ratio data → DATA RANGE "A18:D27"
- Row 28: Empty → SECTION SEPARATOR
- Row 29: "Loan Book by Segments..." → SECTION HEADER (exclude from range)
- Rows 30-34: Loan data → DATA RANGE "A30:D34"
- Continue this pattern through the entire sheet...

Return JSON following this EXACT structure:
{{
  "operation_analysis": {{
    "operation_type": "add_column|update_existing|add_metrics",
    "target_period": "normalized_period_format",
    "original_period": "user_original_format", 
    "confidence": 0.0-1.0,
    "reasoning": "explanation"
  }},
  "table_analysis": {{
    "identified_tables": [
      {{
        "range": "A5:D15",
        "description": "Key Financial Metrics",
        "relevance_score": 0.0-1.0,
        "table_type": "financial_metrics|comparison|summary|loan_book|macro_data",
        "structure": {{
          "header_rows": [1, 2],
          "data_rows": [5, 6, 7, 8, 9],
          "period_columns": [2, 3, 4],
          "metric_columns": [1]
        }}
      }}
    ],
    "primary_table": "most_relevant_range",
    "confidence": 0.0-1.0,
    "total_sections_found": "number_of_distinct_data_sections",
    "scanning_completeness": "ensure_entire_sheet_analyzed"
  }},
  "table_global_contexts": {{
    "A5:D15": {{
      "company_name": {{
        "value": "company_or_empty",
        "is_global": true|false,
        "confidence": 0.0-1.0,
        "source": "cell_reference_or_description",
        "reasoning": "explanation_specific_to_this_table"
      }},
      "entity": {{
        "value": "entity_or_empty",
        "is_global": true|false,
        "confidence": 0.0-1.0,
        "source": "cell_reference_or_description", 
        "reasoning": "explanation_specific_to_this_table"
      }},
      "metric_type": {{
        "value": "Consolidated|Standalone|empty",
        "is_global": true|false,
        "confidence": 0.0-1.0,
        "source": "cell_reference_or_description",
        "reasoning": "explanation_specific_to_this_table"
      }}
    }},
    "A18:D27": {{
      "company_name": {{ "value": "...", "is_global": true|false, "confidence": 0.0-1.0, "source": "...", "reasoning": "..." }},
      "entity": {{ "value": "...", "is_global": true|false, "confidence": 0.0-1.0, "source": "...", "reasoning": "..." }},
      "metric_type": {{ "value": "...", "is_global": true|false, "confidence": 0.0-1.0, "source": "...", "reasoning": "..." }}
    }}
  }},
  "validation": {{
    "feasible": true|false,
    "warnings": ["warning1", "warning2"],
    "errors": ["error1", "error2"],
    "suggestions": ["suggestion1", "suggestion2"]
  }}
}}"""


# =============================================================================
# 2. ORCHESTRATOR DECISION PROMPTS (Used by agent.py)
# =============================================================================

def create_orchestrator_system_prompt():
    """
    System prompt for the main orchestrator decision-making logic.
    Used by llm_reasoning_and_tool_decision in agent.py.
    
    Returns:
        str: Complete system prompt for orchestrator decisions
    """
    return """
You are an Excel modification orchestrator. Analyze the current state and decide the next action.

AVAILABLE TOOLS:
1. identify_table_ranges_for_modification - When no tables identified yet (first run)
2. modify_excel_sheet - When current table needs structural changes (add column/row)
3. cell_mapping_and_fill_current_table - When current table ready for data filling

CRITICAL DECISION LOGIC - FOLLOW IN EXACT ORDER:
1. **FIRST PRIORITY: Check processing_status**
   - If processing_status == "start": Use identify_table_ranges_for_modification (ALWAYS)
   - No other rules apply when status is "start" - tables must be identified first

2. **SECOND PRIORITY: Once tables are identified**
   - If period_exists_globally = True: Use cell_mapping_and_fill_current_table
   - If period_exists_globally = False: Use modify_excel_sheet

3. **THIRD PRIORITY: Completion check**
   - If all tables processed: Return "complete"

IMPORTANT RULES:
1. ALWAYS check if the target period (e.g., "Q2 FY26") already exists as a column header
2. Do NOT add duplicate columns - if "Q2 FY26" already exists, proceed to filling
3. Sequential table processing - complete one table before moving to next
4. If all tables are processed, return "complete"

PERIOD MAPPING RULES - APPLY ONLY AFTER TABLE IDENTIFICATION:
- SHEET-GLOBAL PERIODS: Period columns apply to ENTIRE sheet, not individual tables
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


def create_orchestrator_user_prompt(
    user_question: str,
    processing_status: str,
    current_table_index: int,
    processed_tables: list,
    identified_tables: list,
    current_table: dict,
    target_period: str,
    normalized_target_period: str,
    period_exists_globally: bool,
    period_exists_in_table: bool,
    sheet_period_mapping: dict,
    sheet_columns_added: list,
    current_period_mapping: dict,
    excel_data: str
):
    """
    User prompt template for orchestrator decision-making.
    Used by llm_reasoning_and_tool_decision in agent.py.
    
    Returns:
        str: Complete user prompt for orchestrator decisions
    """
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
    
    return f"""
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

CRITICAL DECISION LOGIC - FOLLOW IN EXACT ORDER:
1. **FIRST PRIORITY**: Processing Status = {processing_status}
   - IF processing_status == "start": USE identify_table_ranges_for_modification (ALWAYS - no other logic applies)
   
2. **SECOND PRIORITY** (only after tables identified): Period exists globally = {period_exists_globally}
   - IF period_exists_globally = True: USE cell_mapping_and_fill_current_table
   - IF period_exists_globally = False: USE modify_excel_sheet
   
3. **IMPORTANT**: Do NOT apply period logic when status is "start" - tables must be identified first
4. Current table period mapping (local): {current_period_mapping}
5. Pay attention to the table range: {current_table.get('range', 'N/A') if current_table else 'N/A'}

TASK: Determine if the target period column already exists in the current table range.
If it exists, proceed to data filling. If not, add the column first.

ANALYZE AND DECIDE NEXT ACTION:
"""


# =============================================================================
# 3. MATCH SELECTION PROMPTS (Used by cell_mapping_and_fill_current_table)
# =============================================================================

def create_match_selection_system_prompt():
    """
    System prompt for LLM-based best match selection from top 5 database results.
    Used by llm_select_best_match in cell_mapping_and_fill_current_table.
    
    Returns:
        str: Complete system prompt for match selection
    """
    return """
You are an expert financial data analyst tasked with selecting the best match from multiple database results.

CONTEXT:
You will be given:
1. Target context (what was requested)
2. Top 5 matches from a vector/semantic search system (ranked by another AI model)

YOUR TASK:
Select the SINGLE BEST match based on:
1. **Company Name Match**: Exact or closest company name match
2. **Entity Match**: Exact or closest entity match  
3. **Metric Type Match**: Exact or closest metric type match
4. **Metric Match**: Exact or closest metric name match
5. **Time Period Match**: Exact or closest time period match

IMPORTANT RULES:
- The provided ranks (1-5) are from another AI model's scoring, but you should make your own independent judgment
- Prioritize EXACT matches over partial matches
- Consider business context (e.g., "HDFC Bank Limited" vs "HDFC Bank" are the same entity)
- If time periods are very close (e.g., Q1 FY25 vs Q1 FY26), prefer the exact match
- Value differences are less important than context matching

Return JSON with:
{
    "selected_rank": <1-5>,
    "reasoning": "Explain why this match is best based on the 5 criteria",
    "confidence": <0.1-1.0>
}
"""


def create_match_selection_user_prompt(cell_mapping: dict, matches_summary: list):
    """
    User prompt template for LLM-based match selection.
    Used by llm_select_best_match in cell_mapping_and_fill_current_table.
    
    Args:
        cell_mapping (dict): Original cell mapping context
        matches_summary (list): List of top 5 matches with metadata
        
    Returns:
        str: Complete user prompt for match selection
    """
    user_prompt = f"""
TARGET CONTEXT (what was requested):
- Company: {cell_mapping.get('company_name', 'N/A')}
- Entity: {cell_mapping.get('entity', 'N/A')}
- Metric Type: {cell_mapping.get('metric_type', 'N/A')}
- Metric: {cell_mapping.get('metric', 'N/A')}
- Time Period: {cell_mapping.get('quarter', 'N/A')}

TOP 5 MATCHES FROM DATABASE:
"""
    
    for i, match in enumerate(matches_summary):
        user_prompt += f"""
Rank {match['rank']}:
- Company: {match['company_name']}
- Entity: {match['entity']}
- Metric Type: {match['metric_type']}
- Metric: {match['metric']}
- Time Period: {match['time_period']}
- Value: {match['value']}
- Hybrid Score: {match['hybrid_score']:.4f}
- Document Year: {match['document_year']}
"""
    
    user_prompt += "\nAnalyze each match against the target context and select the best one."
    
    return user_prompt


# =============================================================================
# 4. HUMAN INTERACTION PROMPTS (Used by agent.py)
# =============================================================================

def create_human_approval_prompt(reasoning_result: dict, state: dict):
    """
    Create formatted prompt for human approval before tool execution.
    Used by request_human_approval in agent.py.
    
    Args:
        reasoning_result (dict): LLM's proposed action
        state (dict): Current agent state
        
    Returns:
        str: Formatted approval prompt
    """
    import json
    
    prompt = f"""
{"="*50}
🤖 AGENT DECISION REQUIRES APPROVAL
{"="*50}
📊 Current Excel State: {len(state.get('excel_data', ''))} chars
🎯 Proposed Tool: {reasoning_result['tool_name']}
💭 Reasoning: {reasoning_result['reasoning']}
⚙️  Parameters: {json.dumps(reasoning_result['parameters'], indent=2)}
"""
    
    if reasoning_result['tool_name'] == 'modify_excel_sheet':
        prompt += "\n⚠️  WARNING: This will modify the Excel file!"
    
    prompt += """

Options:
1. ✅ Approve and proceed
2. ❌ Reject and halt
3. 🔧 Modify parameters

Your decision (1/2/3): """
    
    return prompt


def create_parameter_modification_prompt():
    """
    Prompt for requesting parameter modifications from user.
    Used by request_human_approval in agent.py.
    
    Returns:
        str: Parameter modification prompt
    """
    return "Enter parameter modifications (JSON format):"


# =============================================================================
# 5. UTILITY FUNCTIONS AND EXAMPLES
# =============================================================================

def create_period_normalization_examples():
    """
    Provide examples for period normalization patterns
    
    Returns:
        dict: Mapping of common period formats to normalized formats
    """
    return {
        # Quarter patterns
        "Q3 25": "Q3 FY25",
        "Q2FY26": "Q2 FY26", 
        "Q1 FY 25": "Q1 FY25",
        "3Q25": "Q3 FY25",
        "Q4'24": "Q4 FY24",
        
        # Financial year patterns
        "FY25": "FY25",  # Already correct
        "FY 25": "FY25",
        "2025": "FY25",  # In financial context
        "25": "FY25",    # In quarter context
        
        # Calendar year patterns
        "2024": "CY2024",
        "CY24": "CY2024",
        "CY 24": "CY2024"
    }


def create_global_item_examples():
    """
    Provide examples for global item consistency checking
    
    Returns:
        dict: Examples of global vs non-global scenarios
    """
    return {
        "global_company_examples": [
            "Single company in header (B2: 'HDFC Bank'), all metrics relate to HDFC Bank",
            "Company name appears consistently across table without variation"
        ],
        "non_global_company_examples": [
            "Multiple companies: 'HDFC Bank' in row 5, 'ICICI Bank' in row 6",
            "Company varies by business segment or subsidiary"
        ],
        "global_entity_examples": [
            "All metrics for same business unit: 'Retail Banking'",
            "Entity same as company: 'HDFC Bank' throughout"
        ],
        "non_global_entity_examples": [
            "Mixed entities: 'Retail Banking', 'Investment Banking', 'Life Insurance'",
            "Company varies, so entity definitely varies"
        ],
        "global_metric_type_examples": [
            "Section header 'Standalone Results' covers entire table",
            "All metrics consistently have '(Consolidated)' suffix"
        ],
        "non_global_metric_type_examples": [
            "Mix of 'Revenue (Standalone)' and 'EBITDA (Consolidated)'",
            "No clear consolidated/standalone indicators anywhere"
        ]
    }


def create_table_detection_patterns():
    """
    Provide patterns for table boundary detection
    
    Returns:
        dict: Common patterns for identifying table structures
    """
    return {
        "primary_indicators": [
            "Two or more consecutive empty rows",
            "Two or more consecutive empty columns", 
            "Transition from numbers to text or vice versa",
            "Clear visual separations",
            "Section headers or titles",
            "Significant formatting changes"
        ],
        "secondary_indicators": [
            "Font size changes (headers typically larger)",
            "Alignment patterns (headers center, data right-aligned)",
            "Color pattern differences",
            "Consistent border styles within tables"
        ],
        "structure_patterns": [
            "Header rows contain period information (Q1, Q2, etc.)",
            "Metric rows contain financial terms (Revenue, EBITDA, etc.)",
            "Data cells contain numeric values",
            "Empty cells used for spacing and organization"
        ]
    }


# =============================================================================
# 6. PROMPT FACTORY FUNCTIONS
# =============================================================================

def get_prompt_for_operation(operation_type: str, context: dict = None):
    """
    Factory function to get appropriate prompts based on operation type.
    
    Args:
        operation_type (str): Type of operation (table_analysis, orchestrator_decision, match_selection, human_approval)
        context (dict): Additional context for prompt generation
        
    Returns:
        tuple: (system_prompt, user_prompt_template)
    """
    if operation_type == "table_analysis":
        return (
            create_llm_analysis_system_prompt(),
            create_llm_analysis_user_prompt
        )
    elif operation_type == "orchestrator_decision":
        return (
            create_orchestrator_system_prompt(),
            create_orchestrator_user_prompt
        )
    elif operation_type == "match_selection":
        return (
            create_match_selection_system_prompt(),
            create_match_selection_user_prompt
        )
    elif operation_type == "human_approval":
        return (
            None,  # No system prompt needed
            create_human_approval_prompt
        )
    else:
        raise ValueError(f"Unknown operation type: {operation_type}")


def validate_prompt_completeness():
    """
    Validate that all required prompts are available and properly structured.
    
    Returns:
        dict: Validation results
    """
    validation_results = {
        "missing_prompts": [],
        "available_prompts": [],
        "status": "valid"
    }
    
    prompt_functions = [
        "create_llm_analysis_system_prompt",
        "create_llm_analysis_user_prompt", 
        "create_orchestrator_system_prompt",
        "create_orchestrator_user_prompt",
        "create_match_selection_system_prompt",
        "create_match_selection_user_prompt",
        "create_human_approval_prompt"
    ]
    
    import sys
    current_module = sys.modules[__name__]
    
    for func_name in prompt_functions:
        if hasattr(current_module, func_name):
            validation_results["available_prompts"].append(func_name)
        else:
            validation_results["missing_prompts"].append(func_name)
    
    if validation_results["missing_prompts"]:
        validation_results["status"] = "incomplete"
    
    return validation_results


# =============================================================================
# 7. TESTING AND DEBUGGING
# =============================================================================

if __name__ == "__main__":
    """
    Test all prompt functions to ensure they work correctly
    """
    print("=== Testing LLM Prompts Module ===")
    
    # Test validation
    validation = validate_prompt_completeness()
    print(f"Prompt validation: {validation['status']}")
    print(f"Available prompts: {len(validation['available_prompts'])}")
    
    if validation['missing_prompts']:
        print(f"Missing prompts: {validation['missing_prompts']}")
    
    # Test prompt generation
    try:
        system_prompt = create_llm_analysis_system_prompt()
        print(f"✅ Table analysis system prompt: {len(system_prompt)} chars")
        
        orchestrator_prompt = create_orchestrator_system_prompt()
        print(f"✅ Orchestrator system prompt: {len(orchestrator_prompt)} chars")
        
        match_prompt = create_match_selection_system_prompt()
        print(f"✅ Match selection system prompt: {len(match_prompt)} chars")
        
        print("✅ All core prompts generated successfully!")
        
    except Exception as e:
        print(f"❌ Error testing prompts: {e}")
    
    # Test factory function
    try:
        system, user_template = get_prompt_for_operation("table_analysis")
        print(f"✅ Factory function works: {len(system)} chars")
        
    except Exception as e:
        print(f"❌ Error testing factory function: {e}")
