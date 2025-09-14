from langgraph.graph import StateGraph
from typing import TypedDict, List, Dict, Any
import sys
import os
import json
from openai import AzureOpenAI
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Add parent directory to path to import from scripts
sys.path.append(os.path.join(os.path.dirname(__file__), '..'))
from scripts.excel_to_markdown import parse_sheet_xlsx_with_mapping
from scripts.llm_prompts import create_llm_analysis_system_prompt, create_llm_analysis_user_prompt

class AgentState(TypedDict):
    excel_file_path: str  # Input: Path to Excel file
    excel_data: str       # Output from Node 1: Markdown representation
    user_question: str
    identified_tables: List[Dict]
    current_table: Dict
    operation_type: str
    target_period: str
    processing_status: str
    errors: List[str]
    warnings: List[str]
    excel_metadata: Dict[str, Any]        # Additional metadata from parsing
    llm_analysis: Dict[str, Any]          # Full LLM analysis output
    table_global_contexts: Dict[str, Any] # Table-specific global contexts

# Azure OpenAI Configuration
def get_azure_openai_client():
    """Initialize Azure OpenAI client with environment variables"""
    try:
        # Handle both full URL and deployment name formats
        azure_deployment = os.getenv('AZURE_DEPLOYMENT')
        if azure_deployment.startswith('https://'):
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

# Define the 7 nodes
def sheet_analysis_node(state: AgentState) -> AgentState:
    """
    Node 1: Sheet Analysis
    
    Processes Excel file using parse_sheet_xlsx_with_mapping to convert it to markdown format.
    Updates state with structured Excel data for subsequent analysis.
    
    Args:
        state: Current agent state containing excel_file_path
        
    Returns:
        Updated state with excel_data populated from parsing results
    """
    try:
        # Extract the Excel file path from state
        excel_file_path = state.get("excel_file_path", "")
        
        if not excel_file_path:
            raise ValueError("No Excel file path provided in state")
        
        # Process the Excel file using parse_sheet_xlsx_with_mapping
        parsed_result = parse_sheet_xlsx_with_mapping(excel_file_path)
        
        # Update state with the parsed markdown data
        state["excel_data"] = parsed_result["markdown"]
        state["processing_status"] = "analyzing"
        
        # Store additional metadata for potential use in later nodes
        state["excel_metadata"] = parsed_result["metadata"]
        
        print(f"✅ Node 1 Complete: Successfully parsed Excel file with {parsed_result['metadata']['sheet_info']['rows']} rows and {parsed_result['metadata']['sheet_info']['cols']} columns")
        
    except Exception as e:
        # Handle errors gracefully
        error_msg = f"Error in sheet_analysis_node: {str(e)}"
        state["errors"].append(error_msg)
        state["processing_status"] = "error"
        print(f"❌ Node 1 Error: {error_msg}")
    
    return state

def llm_analysis_node(state: AgentState) -> AgentState:
    """
    Node 2: LLM Analysis
    
    Comprehensive analysis using Azure OpenAI GPT-4o to:
    - Classify user operation intent  
    - Identify table ranges and boundaries
    - Extract global context items
    - Validate feasibility and normalize periods
    
    Args:
        state: Current agent state with excel_data and user_question
        
    Returns:
        Updated state with operation_type, target_period, identified_tables, and global_items
    """
    try:
        # Get Azure OpenAI client
        client = get_azure_openai_client()
        if not client:
            raise Exception("Failed to initialize Azure OpenAI client")
        
        # Prepare input data for LLM
        excel_markdown = state.get("excel_data", "")
        user_question = state.get("user_question", "")
        excel_metadata = state.get("excel_metadata", {})
        
        if not excel_markdown or not user_question:
            raise ValueError("Missing required input: excel_data or user_question")
        
        # Construct the perfect prompt
        system_prompt = create_llm_analysis_system_prompt()
        user_prompt = create_llm_analysis_user_prompt(excel_markdown, user_question, excel_metadata)
        
        print(f"🤖 Node 2: Starting LLM analysis with GPT-4o...")
        print(f"📊 Excel data: {len(excel_markdown)} characters")
        print(f"❓ User question: {user_question}")
        
        # Make Azure OpenAI API call
        response = client.chat.completions.create(
            model=os.getenv('DEPLOYMENT_NAME'),  # Your GPT-4o deployment name
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt}
            ],
            temperature=0,  # Low temperature for consistent, structured output
            max_tokens=4000,
            response_format={"type": "json_object"}  # Enforce JSON output
        )
        
        # Parse LLM response
        llm_response_text = response.choices[0].message.content
        llm_analysis = json.loads(llm_response_text)
        
        print(f"✅ LLM Analysis Complete!")
        print(f"🎯 Operation: {llm_analysis.get('operation_analysis', {}).get('operation_type', 'unknown')}")
        print(f"📅 Target Period: {llm_analysis.get('operation_analysis', {}).get('target_period', 'unknown')}")
        print(f"📋 Tables Found: {len(llm_analysis.get('table_analysis', {}).get('identified_tables', []))}")
        
        # Update state with LLM analysis results
        state["llm_analysis"] = llm_analysis
        
        # Extract key fields for downstream nodes
        operation_analysis = llm_analysis.get("operation_analysis", {})
        state["operation_type"] = operation_analysis.get("operation_type", "")
        state["target_period"] = operation_analysis.get("target_period", "")
        
        table_analysis = llm_analysis.get("table_analysis", {})
        state["identified_tables"] = table_analysis.get("identified_tables", [])
        
        # Store table-specific global contexts
        state["table_global_contexts"] = llm_analysis.get("table_global_contexts", {})
        
        # For backward compatibility, store primary table's global context
        table_analysis = llm_analysis.get("table_analysis", {})
        primary_table = table_analysis.get("primary_table", "")
        if "current_table" not in state:
            state["current_table"] = {}
        
        if primary_table and primary_table in state["table_global_contexts"]:
            state["current_table"]["global_items"] = state["table_global_contexts"][primary_table]
        else:
            state["current_table"]["global_items"] = {}
        
        state["processing_status"] = "llm_analysis_complete"
        
        # Check for validation issues
        validation = llm_analysis.get("validation", {})
        if not validation.get("feasible", True):
            state["errors"].extend(validation.get("errors", []))
        
        state["warnings"].extend(validation.get("warnings", []))
        
    except json.JSONDecodeError as e:
        error_msg = f"Failed to parse LLM response as JSON: {str(e)}"
        state["errors"].append(error_msg)
        state["processing_status"] = "error"
        print(f"❌ Node 2 JSON Error: {error_msg}")
        
    except Exception as e:
        error_msg = f"Error in llm_analysis_node: {str(e)}"
        state["errors"].append(error_msg)
        state["processing_status"] = "error"
        print(f"❌ Node 2 Error: {error_msg}")
    
    return state



def table_identification_node(state: AgentState) -> AgentState:
    # Node 3: Find table boundaries
    pass

# ... etc for all 7 nodes

# Additional nodes will be implemented here...
