"""
LLM Prompt Engineering Module

This module contains all prompt templates and prompt generation functions 
for the Excel AI Agent's LLM-based analysis nodes.
"""


def create_llm_analysis_system_prompt():
    """
    Create the system prompt for Node 2: LLM Analysis
    
    This prompt establishes the LLM as an expert financial Excel analyst
    with specific responsibilities for analyzing Excel structures and user requests.
    
    Returns:
        str: Complete system prompt for LLM analysis
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
    Create the user prompt for Node 2: LLM Analysis
    
    This prompt provides the LLM with complete Excel data, user request,
    and detailed analysis instructions with exact JSON structure requirements.
    
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


# Additional prompt templates for future nodes can be added here
def create_cell_mapping_prompt_template():
    """
    Template for future Node 3: Cell Mapping prompts
    (Placeholder for deterministic cell mapping logic)
    """
    pass


def create_api_orchestration_prompt_template():
    """
    Template for future Node 4: API Orchestration prompts
    (Placeholder for API call optimization logic)
    """
    pass
