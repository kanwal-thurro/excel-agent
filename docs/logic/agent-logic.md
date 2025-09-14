# XL Fill Plugin - AI Agent Logic Documentation

## Purpose of this file:

This document provides comprehensive logic and workflow instructions for building a LangGraph-based AI agent that intelligently fills Excel financial templates. The agent is multimodal (text + image processing) and uses the existing xl_fill_plugin API backend.

**📄 For current API implementation details, see [Current API Implementation](./current-api-implementation.md)**

---

## Overview

The AI agent processes Excel files to intelligently fill financial templates. The agent automates Excel financial data filling through a sophisticated 5-node workflow:

1. **Analyzing** Excel sheets using text parsing to understand structural layout of the excel sheet with cell values.
2. **LLM Analysis** - Comprehensive analysis using Large Language Model to: interpret user requests, identify table ranges, extract global context items, and determine operation parameters
3. **Generating** explicit cell mappings that combine global and cell-specific context for database queries using deterministic algorithms
4. **Orchestrating** backend API calls using the existing xl_fill_plugin endpoints with explicit context mappings
5. **Updating** the original Excel file while preserving formatting and structure

### Key Principles

#### **Spatial Mapping Patterns**
- **Quarters/Periods**: Always mapped column-wise in Excel sheets (horizontal progression)
  - Examples: Q1→Q2→Q3 moving left to right across columns
  - Can appear in any row, agent must scan entire sheet
  - Formats vary: "Q3 25", "Q1 FY26", "FY24", "CY23", etc.

- **Metrics**: Always mapped row-wise in Excel sheets (vertical progression)  
  - Examples: Revenue→EBITDA→Net Income moving top to bottom down rows
  - Can appear in any column, agent must scan entire sheet
  - May include embedded information of company name, entity or metric_type like "Deposit Growth % (Standalone)" or "HDFC Bank Retail Banking Market Share"

#### **Context Mapping Strategy**
- **Global Items**: Values that are 100% consistent across ALL cells in a specific table range
  - `company_name`: Company identifier (e.g., "HDFC Bank", "ICICI Bank")
  - `entity`: Subsidiary/Business entity/segment of the company (e.g., "HDFC Bank", "Retail Banking", "ICICI Life Insurance")  
  - `metric_type`: Consolidation level ("Consolidated", "Standalone", or empty string)
  - **Critical**: If ANY cell in the table range has different values of company name, entity or metric_type, the item is NOT global for that table range.

- **Cell-Specific Items**: Values that vary per individual cell location
  - `metric`: Financial metric name extracted from row context and headers
  - `quarter`: Time period extracted from column context and headers

#### **Decision Logic for Global vs Cell-Specific**
The agent must apply this logic for each table range independently:
- **100% consistency rule**: For an item to be global, it must be identical across every single cell in the table range
- **Multi-table independence**: Different table ranges can have completely different global items
- **Fallback strategy**: When in doubt, treat as cell-specific (empty string for global items)

---

## LangGraph Agent Workflow

### Agent State Structure

The LangGraph agent maintains comprehensive state across all workflow nodes. This state structure enables full traceability and supports human-in-the-loop interventions at any step.

```python
Initial_State = {
    # === INPUT DATA ===
    "excel_data": str,           # Complete Excel file parsed to markdown format
                                # Format: Column letters (A,B,C,D...) and row numbers (1,2,3,4...)
                                # Example: "| A | B | C |\n|---|---|---|\n| Company | Q1 | Q2 |\n| HDFC | 100 | 200 |"
    
    
    "user_question": str,        # User's natural language request
                                # Examples: "fill data for Q2 FY26", "update Q1 25 values", "add ROE metrics"
    
    # === ANALYSIS RESULTS ===
    "identified_tables": [       # List of ALL detected table ranges in the entire sheet
        {
            "range": "A1:F15",          # Excel cell range notation
            "description": "HDFC Bank Financial Metrics",  # Human-readable description
            "contains_data": True,      # Whether table has actual data (not just headers)
            "relevance_score": 0.85     # Relevance to user's question
        }
    ],
    
    # === CURRENT PROCESSING CONTEXT ===
    "current_table": {
        "current_table_range": str,              # Cell range being processed (e.g., "A4:F15")
        "current_table_dataframe": pd.DataFrame, # Extracted table as pandas DataFrame
                                                 # Maintains original Excel coordinates as index/columns
        
        "global_items": {                        # Items consistent across ALL cells in this table
            "company_name": str,     # e.g., "HDFC Bank" or "" if multiple companies
            "entity": str,           # e.g., "HDFC Bank" or "Retail Banking" or ""
            "metric_type": str,      # "Consolidated", "Standalone", or ""
            "table_description": str # e.g., "Q1-Q3 FY25 Financial Results"
        },
        
        "cell_mappings": {          # Explicit mapping for each cell to be filled
            "F5": {                 # Excel cell reference
                "company_name": "HDFC Bank",      # Final resolved value (global or cell-specific)
                "entity": "HDFC Bank",            # Final resolved value (global or cell-specific)
                "metric_type": "",                # Final resolved value (global or cell-specific)
                "metric": "Loan Growth %",        # Extracted from row context
                "quarter": "Q2 FY26",            # Normalized format for database
                "source_info": {                  # Traceability information
                    "metric_source": "A5",        # Where metric was extracted from
                    "quarter_source": "F1",       # Where quarter was extracted from
                    "global_source":              # mapping the source of global items if any           
                    {
                        "company_name": "A1"
                        "entity" : "B1"
                        "metric_type" : ""
                    }"B2"         # Where global items were extracted from
                }
            }
        }
    },
    
    # === PROCESSING METADATA ===
    "operation_type": str,       # "add_column", "update_existing", "add_metrics"
    "target_period": str,        # e.g., "Q2 FY26" (user's requested period)
    "processing_status": str,    # "analyzing", "extracting", "calling_api", "updating", "complete"
    "errors": [],               # List of any errors encountered during processing
    "warnings": []              # List of warnings or ambiguities detected
}
```

### State Evolution Across Nodes

**Node 1 (Sheet Analysis):** Populates `excel_data`
**Node 2 (LLM Analysis):** Populates `operation_type`, `target_period`, `identified_tables`, `current_table.global_items`  
**Node 3 (Cell Mapping):** Populates `current_table.dataframe`, `cell_mappings`  
**Node 4 (API Orchestration):** Uses `cell_mappings`, updates `processing_status`  
**Node 5 (Excel Update):** Produces final Excel file, updates `processing_status` to "complete"

### Complete Agent Workflow

#### **Node 1: Sheet Analysis**
**Input:** Excel file + User question  
**Output:** Comprehensive sheet understanding with structured data

**Detailed Actions:**

1. **Excel to Markdown Conversion:**
   ```python
   # Call excel_to_markdown tool to convert Excel to structured markdown format
   # This tool preserves exact cell coordinates and content
   excel_data = excel_to_markdown(excel_file)
   
   # Example output format:
   markdown_output = """
   |   A   |    B     |   C   |   D   |   E   |
   |-------|----------|-------|-------|-------|
   |   1   |          | Q3 25 | Q4 25 | Q1 26 |
   |   2   | HDFC Bank|       |       |       |
   |   3   |          |       |       |       |
   |   4   | Metrics  |       |       |       |
   |   5   | Revenue  | 1000  | 1100  | 1200  |
   """
   ```
   - **Preserve all content**: Including empty cells, formatting hints, merged cell indicators
   - **Maintain coordinates**: Exact mapping to Excel cell references (A1, B2, etc.)
   - **Include metadata**: Cell types (text, number, formula), basic formatting information

2. **Initial Assessment:**
   - **Sheet Complexity**: Single table vs multiple tables vs complex layouts
   - **Data Density**: Sparse vs dense data patterns
   - **Structural Patterns**: Professional template vs ad-hoc layout
   - **Content Analysis**: Use of spacing, empty rows/columns to indicate relationships

3. **State Population:**
   ```python
   state["excel_data"] = excel_to_markdown_output
   state["processing_status"] = "analyzing"
   ```

#### **Node 2: LLM Analysis**
**Input:** User question + Complete Excel markdown data  
**Output:** Comprehensive analysis including operation classification, table identification, and global context extraction

**Overview:**

This consolidated node leverages a Large Language Model to perform intelligent analysis of both the user's intent and the Excel structure simultaneously. The LLM receives the complete markdown representation of the Excel file and the user's question, then returns a structured analysis covering all aspects needed for precise cell mapping.

**Core Responsibilities:**

1. **Intent Classification & Period Normalization**
2. **Table Structure Identification & Boundary Detection**  
3. **Global Context Items Extraction**
4. **Feasibility Validation & Error Detection**

**Detailed LLM Analysis Process:**

1. **Input Preparation:**
   ```python
   llm_input = {
       "excel_markdown": state["excel_data"],        # Complete sheet structure
       "user_question": state["user_question"],     # User's natural language request
       "sheet_metadata": state["excel_metadata"]    # Dimensions, merged cells, etc.
   }
   ```

2. **LLM Prompt Structure:**
   ```
   You are an expert financial Excel analyst. Analyze the provided Excel data and user request to:

   1. OPERATION CLASSIFICATION:
      - Classify as: "add_column", "update_existing", or "add_metrics"
      - Extract and normalize target period (Q3 25 → Q3 FY25, 2024 → CY2024, etc.)
      - Validate feasibility based on existing data structure

   2. TABLE IDENTIFICATION:
      - Identify all logical table ranges (e.g., A4:F15)
      - Detect table boundaries using empty rows/columns and data patterns
      - Determine which table(s) are relevant to the user's request
      - Analyze table structure (header rows, data rows, period columns)

   3. GLOBAL CONTEXT EXTRACTION:
      - Apply 100% consistency rule for global items
      - company_name: Consistent across ALL cells in table range?
      - entity: Consistent business unit/segment across ALL cells?
      - metric_type: "Consolidated"/"Standalone" consistent across ALL cells?

   EXCEL DATA:
   {excel_markdown}

   USER REQUEST: {user_question}

   Return structured JSON following this exact format:
   ```

3. **Expected LLM Output Structure:**
   ```python
   llm_output = {
       "operation_analysis": {
           "operation_type": "add_column",           # or "update_existing", "add_metrics"
           "target_period": "Q2 FY26",              # Normalized format
           "original_period": "Q2 26",              # User's original format
           "confidence": 0.95,                      # Classification confidence
           "reasoning": "User requests filling Q2 26 data, which normalizes to Q2 FY26"
       },
       
       "table_analysis": {
           "identified_tables": [
               {
                   "range": "A1:D15",
                   "description": "HDFC Bank Q1-Q3 FY25 Financial Metrics",
                   "relevance_score": 0.9,
                   "table_type": "financial_metrics",
                   "structure": {
                       "header_rows": [1, 4],
                       "data_rows": [5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15],
                       "period_columns": [2, 3, 4],      # B, C, D (Q3 25, Q4 25, Q1 26)
                       "metric_columns": [1]             # A (metric names)
                   }
               }
           ],
           "primary_table": "A1:D15",               # Most relevant table for user request
           "confidence": 0.92
       },
       
       "table_global_contexts": {
           "A5:D15": {
               "company_name": {
                   "value": "HDFC Bank",
                   "is_global": true,
                   "confidence": 0.95,
                   "source": "A1",
                   "reasoning": "Single company mentioned in header, consistent across all data rows in this table"
               },
               "entity": {
                   "value": "",
                   "is_global": false,
                   "confidence": 0.90,
                   "source": "",
                   "reasoning": "No specific entity mentioned for this financial metrics table"
               },
               "metric_type": {
                   "value": "",
                   "is_global": false,
                   "confidence": 1.0,
                   "source": "",
                   "reasoning": "No consolidated/standalone indicators found in this table"
               }
           },
           "A30:D34": {
               "company_name": {
                   "value": "HDFC Bank",
                   "is_global": true,
                   "confidence": 0.95,
                   "source": "A1",
                   "reasoning": "Consistent across all loan book data rows"
               },
               "entity": {
                   "value": "Loan Book",
                   "is_global": true,
                   "confidence": 0.85,
                   "source": "A29",
                   "reasoning": "Loan Book segment consistent across all rows in this table"
               },
               "metric_type": {
                   "value": "",
                   "is_global": false,
                   "confidence": 0.0,
                   "source": "",
                   "reasoning": "Mixed metric types within loan book data"
               }
           }
       },
       
       "validation": {
           "feasible": true,
           "warnings": [],
           "errors": [],
           "suggestions": ["Target period Q2 FY26 will be added as column E"]
       }
   }
   ```

4. **Period Normalization Logic:**
   The LLM is instructed to normalize periods according to these patterns:
   - "Q3 25" → "Q3 FY25"
   - "Q2FY26" → "Q2 FY26"  
   - "Q1 FY 25" → "Q1 FY25"
   - "2024" → "CY2024"
   - "FY 25" → "FY25"

5. **Global Items Decision Framework:**
   The LLM applies the **100% consistency rule**:
   - **Global**: Item value is identical across ALL cells in table range
   - **Not Global**: Item varies across cells or is unclear
   - **Sources tracked**: Where each global item was detected

6. **Error Handling & Validation:**
   ```python
   validation_scenarios = {
       "ambiguous_request": "Multiple possible interpretations",
       "period_not_found": "No valid period extracted from user request",
       "no_relevant_tables": "No tables match user's intent",
       "conflicting_global_items": "Inconsistent global context detected"
   }
   ```

7. **State Population:**
   ```python
   # Populate multiple state fields from single LLM response
   state["operation_type"] = llm_output["operation_analysis"]["operation_type"]
   state["target_period"] = llm_output["operation_analysis"]["target_period"]
   state["identified_tables"] = llm_output["table_analysis"]["identified_tables"]
   
   # Store table-specific global contexts
   state["table_global_contexts"] = llm_output["table_global_contexts"]
   
   # For backward compatibility, store primary table's global context
   primary_table = llm_output["table_analysis"]["primary_table"]
   if primary_table in llm_output["table_global_contexts"]:
       state["current_table"]["global_items"] = llm_output["table_global_contexts"][primary_table]
   
   state["processing_status"] = "llm_analysis_complete"
   
   # Store full LLM analysis for downstream nodes
   state["llm_analysis"] = llm_output
   ```

**Key Advantages of LLM-Based Analysis:**

- **Holistic Understanding**: Processes user intent and Excel structure simultaneously
- **Context-Aware Decisions**: Makes intelligent inferences about table relationships
- **Flexible Pattern Recognition**: Handles various Excel layouts and naming conventions
- **Natural Language Processing**: Understands nuanced user requests
- **Consistent Output**: Structured JSON ensures reliable downstream processing

#### **Node 3: Cell Mapping (Deterministic)**
**Input:** LLM analysis output + Sheet structure  
**Output:** Modified DataFrame + Precise cell coordinates + Complete cell mappings

**Overview:**

This node takes the structured output from Node 2 (LLM Analysis) and performs deterministic cell coordinate mapping and DataFrame modifications. Unlike the previous node which relies on AI interpretation, this node uses algorithmic approaches to generate precise Excel cell references and mappings.

**Critical Tool Integration: DataFrame Modification Tool**

This node involves calling a specialized tool that performs deterministic cell-wise, row-wise, and column-wise modifications to pandas DataFrames based on the LLM analysis.

**Detailed Processing Steps:**

1. **Table Extraction and Preparation:**
   
   **DataFrame Creation:**
   ```python
   def extract_table_as_dataframe(table_range, excel_data):
       """
       Convert Excel range to pandas DataFrame maintaining exact coordinates
       """
       # Parse table range (e.g., "A4:F15")
       start_col, start_row, end_col, end_row = parse_excel_range(table_range)
       
       # Extract relevant section from markdown data
       extracted_data = []
       for row in range(start_row, end_row + 1):
           row_data = []
           for col in range(start_col, end_col + 1):
               cell_ref = f"{excel_col_name(col)}{row}"
               cell_value = get_cell_value(excel_data, cell_ref)
               row_data.append(cell_value)
           extracted_data.append(row_data)
       
       # Create DataFrame with Excel-style column names and row indices
       column_names = [excel_col_name(col) for col in range(start_col, end_col + 1)]
       row_indices = list(range(start_row, end_row + 1))
       
       df = pd.DataFrame(extracted_data, columns=column_names, index=row_indices)
       return df
   ```

2. **Tool Invocation with Complete Context:**
   
   **Tool Input Structure:**
   ```python
   tool_input = {
       "operation_type": "add_column",           # or "update_existing", "add_metrics"
       "target_period": "Q2 FY26",              # Normalized period format
       "target_metric": None,                    # For add_metrics operations
       "current_dataframe": pandas_dataframe,   # Extracted table as DataFrame
       "global_items": {                        # Resolved global items
           "company_name": "HDFC Bank",
           "entity": "HDFC Bank", 
           "metric_type": ""
       },
       "table_metadata": {
           "original_range": "A4:F15",
           "header_rows": [1, 4],
           "data_rows": [5, 6, 7, 8, 9],
           "period_columns": [3, 4, 5],  # C, D, E columns
           "metric_columns": [1, 2]      # A, B columns
       },
       "user_preferences": {
           "period_format": "Q3 25",      # How periods appear in original sheet
           "insert_position": "next_available"  # Where to add new column
       }
   }
   ```

3. **Tool Processing Logic:**
   
   **For "add_column" Operations:**
   ```python
   def process_add_column(tool_input):
       """
       Tool logic for adding new quarter column
       """
       df = tool_input["current_dataframe"]
       target_period = tool_input["target_period"]
       global_items = tool_input["global_items"]
       
       # Determine insertion position
       period_columns = tool_input["table_metadata"]["period_columns"]
       next_col_idx = max(period_columns) + 1
       next_col_name = excel_col_name(next_col_idx)
       
       # Add new column with appropriate header
       display_period = format_period_for_display(target_period, tool_input["user_preferences"]["period_format"])
       
       # Insert new column
       df[next_col_name] = None  # Initialize with empty values
       
       # Set header for new column
       header_rows = tool_input["table_metadata"]["header_rows"]
       for header_row in header_rows:
           if header_row in df.index and is_period_header_row(df, header_row):
               df.loc[header_row, next_col_name] = display_period
       
       # Generate cell mappings for data cells that need filling
       data_rows = tool_input["table_metadata"]["data_rows"]
       cell_mappings = {}
       
       for data_row in data_rows:
           if data_row in df.index:
               cell_ref = f"{next_col_name}{data_row}"
               
               # Extract metric from row context
               metric_columns = tool_input["table_metadata"]["metric_columns"]
               row_metric = extract_metric_from_row(df, data_row, metric_columns)
               
               # Create cell mapping
               cell_mappings[cell_ref] = {
                   "company_name": global_items.get("company_name", ""),
                   "entity": global_items.get("entity", ""), 
                   "metric_type": global_items.get("metric_type", ""),
                   "metric": row_metric,
                   "quarter": target_period,  # Normalized format for API
                   "source_info": {
                       "metric_source": f"A{data_row}",  # Where metric was extracted
                       "quarter_source": f"{next_col_name}{header_rows[0]}",  # Header cell
                       "global_source": "table_analysis"
                   }
               }
       
       return {
           "modified_dataframe": df,
           "cell_coordinates": list(cell_mappings.keys()),
           "cell_mappings": cell_mappings,
           "operation_summary": f"Added column {next_col_name} for {display_period}"
       }
   ```

   **For "update_existing" Operations:**
   ```python
   def process_update_existing(tool_input):
       """
       Tool logic for updating existing quarter data
       """
       df = tool_input["current_dataframe"]
       target_period = tool_input["target_period"]
       
       # Find existing column with target period
       existing_col = find_period_column(df, target_period)
       if not existing_col:
           raise ValueError(f"Period {target_period} not found in existing columns")
       
       # Generate cell mappings for existing data cells
       data_rows = tool_input["table_metadata"]["data_rows"]
       cell_mappings = {}
       
       for data_row in data_rows:
           cell_ref = f"{existing_col}{data_row}"
           row_metric = extract_metric_from_row(df, data_row, tool_input["table_metadata"]["metric_columns"])
           
           cell_mappings[cell_ref] = {
               "company_name": tool_input["global_items"].get("company_name", ""),
               "entity": tool_input["global_items"].get("entity", ""),
               "metric_type": tool_input["global_items"].get("metric_type", ""),
               "metric": row_metric,
               "quarter": target_period
           }
       
       return {
           "modified_dataframe": df,  # No structural changes for updates
           "cell_coordinates": list(cell_mappings.keys()),
           "cell_mappings": cell_mappings,
           "operation_summary": f"Updated existing column {existing_col} for {target_period}"
       }
   ```

4. **Cell-Specific Extraction for Non-Global Items:**
   
   **When Global Items are Empty (Multi-Company Tables):**
   ```python
   def extract_cell_specific_context(df, cell_ref, global_items):
       """
       Extract company/entity/metric_type from cell context when not global
       """
       col, row = parse_cell_reference(cell_ref)
       
       cell_context = {
           "company_name": global_items.get("company_name", ""),
           "entity": global_items.get("entity", ""),
           "metric_type": global_items.get("metric_type", "")
       }
       
       # If any global item is empty, extract from cell context
       if not cell_context["company_name"]:
           # Look for company in row context (typically first column)
           cell_context["company_name"] = extract_company_from_row(df, row)
       
       if not cell_context["entity"]:
           # Entity might be same as company or extracted separately
           cell_context["entity"] = extract_entity_from_row(df, row) or cell_context["company_name"]
       
       if not cell_context["metric_type"]:
           # Check for embedded metric type in metric name or section headers
           cell_context["metric_type"] = extract_metric_type_from_context(df, row, col)
       
       return cell_context
   ```

5. **Tool Response Processing:**
   
   **Complete Tool Output:**
   ```python
   tool_output = {
       "modified_dataframe": updated_pandas_df,
       "cell_coordinates": ["F5", "F6", "F7", "F8", "F9"],  # Cells to fill
       "cell_mappings": {
           "F5": {
               "company_name": "HDFC Bank",
               "entity": "HDFC Bank",
               "metric_type": "",
               "metric": "Loan Growth %",
               "quarter": "Q2 FY26",
               "source_info": {...}
           },
           # ... mappings for F6, F7, F8, F9
       },
       "operation_summary": "Added column F for Q2 FY26 with 5 cells to fill",
       "excel_coordinates": {
           "new_column": "F",
           "header_cell": "F1", 
           "data_range": "F5:F9"
       }
   }
   ```

6. **State Population:**
   ```python
   state["current_table"]["current_table_dataframe"] = tool_output["modified_dataframe"]
   state["current_table"]["cell_mappings"] = tool_output["cell_mappings"]
   state["operation_summary"] = tool_output["operation_summary"]
   state["processing_status"] = "dataframe_processed"
   ```

**Error Handling:**
- **Period Format Mismatch**: Tool normalizes between display format and database format
- **Missing Context**: Tool requests additional context from agent
- **Invalid Operations**: Tool validates operations before execution
- **Coordinate Conflicts**: Tool ensures no overwrites of existing data

#### **Node 4: API Orchestration**
**Input:** Complete cell mappings from DataFrame tool + Validation requirements  
**Output:** Database values + Metadata + Error handling

**Comprehensive API Integration Process:**

1. **Pre-API Validation:**
   
   **Data Quality Checks:**
   ```python
   def validate_cell_mappings(cell_mappings):
       """
       Validate all cell mappings before making API calls
       """
       validation_results = {}
       
       for cell_ref, mapping in cell_mappings.items():
           cell_validation = {
               "valid": True,
               "warnings": [],
               "errors": []
           }
           
           # Check required fields
           if not mapping.get("metric"):
               cell_validation["errors"].append("Missing metric name")
               cell_validation["valid"] = False
           
           if not mapping.get("quarter"):
               cell_validation["errors"].append("Missing quarter")
               cell_validation["valid"] = False
           
           # Check period format normalization
           normalized_quarter = normalize_period_format(mapping["quarter"])
           if normalized_quarter != mapping["quarter"]:
               cell_validation["warnings"].append(
                   f"Period normalized from {mapping['quarter']} to {normalized_quarter}"
               )
               mapping["quarter"] = normalized_quarter
           
           # Validate company/entity combinations
           if mapping.get("company_name") and not mapping.get("entity"):
               mapping["entity"] = mapping["company_name"]  # Default entity to company
           
           validation_results[cell_ref] = cell_validation
       
       return validation_results
   ```

2. **Period Format Normalization:**
   
   **Critical Conversion Logic:**
   ```python
   def normalize_period_format(user_period):
       """
       Convert various period formats to database-compatible format
       CRITICAL: Database expects exact format matching
       """
       period_patterns = {
           # Quarter patterns
           r"Q(\d)\s*(\d{2})": r"Q\1 FY\2",           # "Q3 25" -> "Q3 FY25"
           r"Q(\d)FY(\d{2})": r"Q\1 FY\2",           # "Q3FY25" -> "Q3 FY25"  
           r"Q(\d)\s*FY\s*(\d{2})": r"Q\1 FY\2",     # "Q3 FY 25" -> "Q3 FY25"
           
           # Year patterns  
           r"FY\s*(\d{2})": r"FY\1",                 # "FY 25" -> "FY25"
           r"(\d{4})": r"CY\1",                      # "2024" -> "CY2024"
           r"CY\s*(\d{2})": r"CY20\1"                # "CY24" -> "CY2024"
       }
       
       for pattern, replacement in period_patterns.items():
           if re.match(pattern, user_period):
               return re.sub(pattern, replacement, user_period)
       
       # If no pattern matches, return as-is and log warning
       return user_period
   ```

3. **Parallel API Call Strategy:**
   
   **Batch Processing with Concurrency Control:**
   ```python
   async def orchestrate_api_calls(cell_mappings, max_concurrent=5):
       """
       Process multiple cells in parallel with controlled concurrency
       """
       # Group cells by similar context for potential optimization
       call_groups = group_similar_mappings(cell_mappings)
       
       all_results = {}
       semaphore = asyncio.Semaphore(max_concurrent)
       
       async def process_single_cell(cell_ref, mapping):
           async with semaphore:
               try:
                   # Call existing xl_fill_plugin API
                   response = await call_xl_fill_api(mapping)
                   return cell_ref, "success", response
               except Exception as e:
                   return cell_ref, "error", str(e)
       
       # Create tasks for all cells
       tasks = [
           process_single_cell(cell_ref, mapping)
           for cell_ref, mapping in cell_mappings.items()
       ]
       
       # Execute in parallel
       results = await asyncio.gather(*tasks, return_exceptions=True)
       
       # Process results
       for result in results:
           if isinstance(result, Exception):
               continue
           cell_ref, status, data = result
           all_results[cell_ref] = {"status": status, "data": data}
       
       return all_results
   ```

4. **API Response Processing:**
   
   **Data Extraction and Validation:**
   ```python
   def process_api_responses(api_results):
       """
       Extract values and metadata from API responses
       """
       processed_results = {}
       
       for cell_ref, result in api_results.items():
           if result["status"] == "success":
               api_data = result["data"]
               
               if api_data.get("matched_values"):
                   # Extract the actual value and metadata
                   matched_data = list(api_data["matched_values"].values())[0]
                   
                   processed_results[cell_ref] = {
                       "value": matched_data.get("value", ""),
                       "confidence_score": extract_confidence_score(api_data),
                       "source_url": matched_data.get("source_url", ""),
                       "value_in": matched_data.get("value_in", ""),
                       "units": matched_data.get("units", ""),
                       "document_year": matched_data.get("document_year", ""),
                       "database_company": matched_data.get("company_name", ""),
                       "database_entity": matched_data.get("entity", ""),
                       "database_metric": matched_data.get("metric", ""),
                       "database_metric_type": matched_data.get("metric_type", ""),
                       "status": "filled"
                   }
               else:
                   # No data found in database
                   processed_results[cell_ref] = {
                       "value": "",
                       "status": "no_data",
                       "reason": "No matching data found in database"
                   }
           else:
               # API call failed
               processed_results[cell_ref] = {
                   "value": "",
                   "status": "error", 
                   "reason": result["data"]
               }
       
       return processed_results
   ```

5. **Error Handling and Retry Logic:**
   
   **Comprehensive Error Management:**
   ```python
   def handle_api_errors(failed_cells, original_mappings):
       """
       Implement retry and fallback strategies for failed API calls
       """
       retry_strategies = {
           "timeout": "retry_with_longer_timeout",
           "rate_limit": "retry_with_backoff", 
           "invalid_metric": "suggest_similar_metrics",
           "no_data": "expand_search_criteria"
       }
       
       recovery_results = {}
       for cell_ref, error_info in failed_cells.items():
           error_type = classify_error(error_info)
           strategy = retry_strategies.get(error_type, "manual_review")
           
           if strategy == "suggest_similar_metrics":
               # Use fuzzy matching to suggest alternatives
               similar_metrics = find_similar_metrics(original_mappings[cell_ref]["metric"])
               recovery_results[cell_ref] = {
                   "status": "needs_clarification",
                   "suggestions": similar_metrics
               }
       
       return recovery_results
   ```

6. **State Population:**
   ```python
   state["api_results"] = processed_api_results
   state["failed_cells"] = failed_api_calls  
   state["success_rate"] = calculate_success_rate(processed_api_results)
   state["processing_status"] = "api_complete"
   ```

#### **Node 5: Excel Update**
**Input:** API results + Original Excel structure + Formatting preservation requirements  
**Output:** Complete updated Excel file with metadata

**Detailed Excel Integration Process:**

1. **Value Application Strategy:**
   
   **Precise Cell Updates:**
   ```python
   def apply_values_to_excel(original_excel, api_results, table_structure):
       """
       Apply API results to original Excel while preserving all formatting
       """
       # Load original Excel with formatting preservation
       workbook = load_excel_with_formatting(original_excel)
       worksheet = workbook.active
       
       update_summary = {
           "cells_updated": 0,
           "cells_failed": 0,
           "new_columns_added": 0,
           "formatting_preserved": True
       }
       
       for cell_ref, result_data in api_results.items():
           if result_data["status"] == "filled":
               # Apply the actual value
               cell = worksheet[cell_ref]
               cell.value = float(result_data["value"]) if result_data["value"].replace('.','').isdigit() else result_data["value"]
               
               # Preserve or apply appropriate formatting
               if cell_ref in table_structure["data_cells"]:
                   # Apply number formatting for data cells
                   cell.number_format = detect_number_format(result_data)
               
               update_summary["cells_updated"] += 1
               
               # Add metadata as cell comment (optional)
               if result_data.get("source_url"):
                   add_metadata_comment(cell, result_data)
           
           elif result_data["status"] == "no_data":
               # Mark cells with no data appropriately
               cell = worksheet[cell_ref]
               cell.value = "N/A"
               cell.font = Font(color="FF9999")  # Light red for missing data
               update_summary["cells_failed"] += 1
       
       return workbook, update_summary
   ```

2. **Formatting Preservation:**
   
   **Advanced Formatting Handling:**
   ```python
   def preserve_excel_formatting(original_file, updated_data):
       """
       Ensure all original formatting is maintained
       """
       preservation_strategies = {
           "number_formats": preserve_number_formatting,
           "cell_styles": preserve_cell_styles,
           "conditional_formatting": preserve_conditional_formatting,
           "merged_cells": preserve_merged_cells,
           "column_widths": preserve_column_widths,
           "row_heights": preserve_row_heights,
           "borders": preserve_border_styles,
           "colors": preserve_color_schemes
       }
       
       for strategy_name, strategy_func in preservation_strategies.items():
           try:
               strategy_func(original_file, updated_data)
           except Exception as e:
               log_formatting_warning(strategy_name, e)
   ```

3. **Metadata Integration:**
   
   **Optional Metadata Tracking:**
   ```python
   def add_metadata_tracking(workbook, api_results, update_summary):
       """
       Add metadata sheet for audit trail and data source tracking
       """
       # Create metadata worksheet
       metadata_sheet = workbook.create_sheet("Data_Sources")
       
       # Add headers
       metadata_sheet["A1"] = "Cell Reference"
       metadata_sheet["B1"] = "Value"
       metadata_sheet["C1"] = "Source URL"
       metadata_sheet["D1"] = "Document Year"
       metadata_sheet["E1"] = "Last Updated"
       metadata_sheet["F1"] = "Confidence Score"
       
       # Add data for each filled cell
       row = 2
       for cell_ref, result_data in api_results.items():
           if result_data["status"] == "filled":
               metadata_sheet[f"A{row}"] = cell_ref
               metadata_sheet[f"B{row}"] = result_data["value"]
               metadata_sheet[f"C{row}"] = result_data.get("source_url", "")
               metadata_sheet[f"D{row}"] = result_data.get("document_year", "")
               metadata_sheet[f"E{row}"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
               metadata_sheet[f"F{row}"] = result_data.get("confidence_score", "")
               row += 1
   ```

4. **Quality Assurance:**
   
   **Final Validation:**
   ```python
   def validate_final_excel(updated_workbook, original_structure):
       """
       Ensure the updated Excel meets quality standards
       """
       validation_checks = {
           "all_cells_accessible": verify_cell_accessibility,
           "formatting_intact": verify_formatting_preservation,
           "data_types_correct": verify_data_type_consistency,
           "no_formula_errors": verify_formula_integrity,
           "metadata_complete": verify_metadata_completeness
       }
       
       validation_results = {}
       for check_name, check_func in validation_checks.items():
           validation_results[check_name] = check_func(updated_workbook, original_structure)
       
       return validation_results
   ```

5. **Final State Population:**
   ```python
   state["final_excel_file"] = updated_workbook
   state["update_summary"] = update_summary
   state["validation_results"] = validation_results
   state["processing_status"] = "complete"
   
   # Generate completion report
   state["completion_report"] = {
       "total_cells_processed": len(api_results),
       "successful_fills": update_summary["cells_updated"],
       "failed_fills": update_summary["cells_failed"],
       "success_rate": update_summary["cells_updated"] / len(api_results),
       "new_columns_added": update_summary["new_columns_added"],
       "processing_time": calculate_total_processing_time(state),
       "data_sources": extract_unique_sources(api_results)
   }
   ```

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

#### **Normalization Algorithm**

```python
def normalize_period_for_database(display_period):
    """
    Convert any period format to database-compatible format
    """
    # Remove common separators and spaces
    cleaned = re.sub(r'[^\w]', '', display_period.upper())
    
    # Quarter patterns (highest priority)
    if re.match(r'Q\d+\d{2}', cleaned):
        # Q325 -> Q3 FY25
        quarter = cleaned[0:2]
        year = cleaned[2:]
        return f"{quarter} FY{year}"
    
    # Financial year patterns
    elif re.match(r'FY\d{2}', cleaned):
        return f"FY{cleaned[2:]}"
    
    # Calendar year patterns  
    elif re.match(r'\d{4}', cleaned):
        return f"CY{cleaned}"
    
    # Default: return as-is with warning
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