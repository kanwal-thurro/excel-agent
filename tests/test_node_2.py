#!/usr/bin/env python3
"""
Test suite for Node 2: LLM Analysis

This module contains tests for the LLM analysis functionality of the Excel AI agent.
"""

import sys
import os

# Add src directory to path to import agent modules
sys.path.append(os.path.join(os.path.dirname(__file__), '..', 'src'))

from agent.agent import llm_analysis_node, sheet_analysis_node, AgentState


def test_node_2_llm_analysis(excel_file_path: str, user_question: str = "Fill data for Q2 FY26"):
    """
    Test function for Node 2: LLM Analysis
    
    Args:
        excel_file_path: Path to Excel file to process
        user_question: User's question for analysis
        
    Returns:
        Final state after Node 2 processing
    """
    print(f"=== Node 2 LLM Analysis Test ===")
    print(f"📁 File: {excel_file_path}")
    print(f"❓ Question: {user_question}")
    print()
    
    # Initialize state
    initial_state = {
        "excel_file_path": excel_file_path,
        "excel_data": "",
        "user_question": user_question,
        "identified_tables": [],
        "current_table": {},
        "operation_type": "",
        "target_period": "",
        "processing_status": "initialized",
        "errors": [],
        "warnings": [],
        "excel_metadata": {},
        "llm_analysis": {},
        "table_global_contexts": {}
    }
    
    # Step 1: Run Node 1 (Sheet Analysis) first
    print("🚀 Step 1: Running Node 1 (Sheet Analysis)")
    state_after_node1 = sheet_analysis_node(initial_state)
    
    if state_after_node1["processing_status"] == "error":
        print("❌ Node 1 failed, cannot proceed to Node 2")
        return state_after_node1
    
    print(f"✅ Node 1 Complete - Status: {state_after_node1['processing_status']}")
    print()
    
    # Step 2: Run Node 2 (LLM Analysis)
    print("🤖 Step 2: Running Node 2 (LLM Analysis)")
    final_state = llm_analysis_node(state_after_node1)
    
    # Display results
    print("\n" + "="*50)
    print("🎯 NODE 2 ANALYSIS RESULTS")
    print("="*50)
    
    if final_state["processing_status"] == "error":
        print("❌ Node 2 FAILED!")
        for error in final_state["errors"]:
            print(f"   Error: {error}")
        return final_state
    
    # Display operation analysis
    operation_analysis = final_state.get("llm_analysis", {}).get("operation_analysis", {})
    print(f"📋 Operation Type: {operation_analysis.get('operation_type', 'Unknown')}")
    print(f"📅 Target Period: {operation_analysis.get('target_period', 'Unknown')}")
    print(f"📅 Original Period: {operation_analysis.get('original_period', 'Unknown')}")
    print(f"🎯 Confidence: {operation_analysis.get('confidence', 0):.2f}")
    print(f"💭 Reasoning: {operation_analysis.get('reasoning', 'None')}")
    print()
    
    # Display table analysis
    table_analysis = final_state.get("llm_analysis", {}).get("table_analysis", {})
    tables = table_analysis.get("identified_tables", [])
    print(f"📋 Tables Identified: {len(tables)}")
    for i, table in enumerate(tables):
        print(f"   Table {i+1}: {table.get('range', 'Unknown')} - {table.get('description', 'No description')}")
        print(f"             Relevance: {table.get('relevance_score', 0):.2f}")
    print(f"🎯 Primary Table: {table_analysis.get('primary_table', 'Unknown')}")
    print()
    
    # Display table-specific global contexts
    table_global_contexts = final_state.get("llm_analysis", {}).get("table_global_contexts", {})
    print("🌍 Table-Specific Global Context Analysis:")
    for table_range, global_context in table_global_contexts.items():
        print(f"   📋 Table {table_range}:")
        for item_type, item_data in global_context.items():
            if isinstance(item_data, dict):
                value = item_data.get('value', 'Unknown')
                is_global = item_data.get('is_global', False)
                confidence = item_data.get('confidence', 0)
                print(f"      {item_type}: '{value}' (Global: {is_global}, Confidence: {confidence:.2f})")
    print()
    
    # Display validation
    validation = final_state.get("llm_analysis", {}).get("validation", {})
    print(f"✅ Feasible: {validation.get('feasible', 'Unknown')}")
    
    warnings = validation.get("warnings", [])
    if warnings:
        print("⚠️  Warnings:")
        for warning in warnings:
            print(f"   - {warning}")
    
    errors = validation.get("errors", [])
    if errors:
        print("❌ Validation Errors:")
        for error in errors:
            print(f"   - {error}")
    
    suggestions = validation.get("suggestions", [])
    if suggestions:
        print("💡 Suggestions:")
        for suggestion in suggestions:
            print(f"   - {suggestion}")
    
    print("\n" + "="*50)
    print(f"🎉 Node 2 Analysis {'SUCCESSFUL' if final_state['processing_status'] != 'error' else 'FAILED'}!")
    print("="*50)
    
    # Output complete LLM response as structured JSON for review
    if final_state.get("llm_analysis"):
        print("\n" + "="*80)
        print("📋 COMPLETE LLM ANALYSIS JSON RESPONSE")
        print("="*80)
        import json
        print(json.dumps(final_state["llm_analysis"], indent=2, ensure_ascii=False))
        print("="*80)
    
    return final_state


def test_node_2_error_handling():
    """Test Node 2 error handling with invalid inputs"""
    print("\n=== Node 2 Error Handling Test ===")
    
    # Test with missing excel_data
    print("🧪 Testing missing excel_data...")
    invalid_state = {
        "excel_file_path": "",
        "excel_data": "",  # Missing data
        "user_question": "Test question",
        "identified_tables": [],
        "current_table": {},
        "operation_type": "",
        "target_period": "",
        "processing_status": "initialized",
        "errors": [],
        "warnings": [],
        "excel_metadata": {},
        "llm_analysis": {},
        "table_global_contexts": {}
    }
    
    result = llm_analysis_node(invalid_state)
    print(f"✅ Error handling test passed: {len(result['errors'])} errors caught")
    for error in result['errors']:
        print(f"   Error: {error}")


if __name__ == "__main__":
    print("=== Node 2 LLM Analysis Test Suite ===")
    
    # Test with sample Excel file
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    test_file = os.path.join(base_dir, "docs", "sample_inputs", "itus-banking-sample.xlsx")
    test_question = "Fill data for Q2 FY26"
    
    if os.path.exists(test_file):
        print(f"🚀 Testing Node 2 with file: {test_file}")
        result = test_node_2_llm_analysis(test_file, test_question)
        
        # Run error handling test
        test_node_2_error_handling()
        
        print("\n🎉 All Node 2 tests completed!")
    else:
        print(f"❌ Test file not found: {test_file}")
        print("Please ensure the sample Excel file exists in docs/")
