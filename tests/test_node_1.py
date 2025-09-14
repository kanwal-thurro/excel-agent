#!/usr/bin/env python3
"""
Test suite for Node 1: Sheet Analysis

This module contains tests for the sheet analysis functionality of the Excel AI agent.
"""

import sys
import os

# Add src directory to path to import agent modules
sys.path.append(os.path.join(os.path.dirname(__file__), '..', 'src'))

from agent.agent import sheet_analysis_node, AgentState


def test_node_1(excel_file_path: str, user_question: str = "Test question"):
    """
    Test function for Node 1: Sheet Analysis
    
    Args:
        excel_file_path: Path to Excel file to process
        user_question: User's question (for testing purposes)
        
    Returns:
        Final state after Node 1 processing
    """
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
        "excel_metadata": {}
    }
    
    print(f"🚀 Testing Node 1 with file: {excel_file_path}")
    
    # Process through Node 1
    result_state = sheet_analysis_node(initial_state)
    
    print(f"📊 Processing Status: {result_state['processing_status']}")
    if result_state.get("errors"):
        print(f"❌ Errors: {result_state['errors']}")
    
    return result_state


def test_node_1_with_sample_file():
    """
    Test Node 1 with the sample HDFC Bank Excel file
    """
    # Get the correct path to the sample file
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    test_file = os.path.join(base_dir, "docs", "itus-banking-sample.xlsx")
    test_question = "Fill data for Q2 FY26"
    
    # Run the test
    result = test_node_1(test_file, test_question)
    
    # Validate results
    if result["processing_status"] == "analyzing":
        print("\n✅ Node 1 Test Successful!")
        print(f"Excel data preview (first 500 chars):")
        excel_data_preview = result["excel_data"][:500] + "..." if len(result["excel_data"]) > 500 else result["excel_data"]
        print(excel_data_preview)
        
        # Additional validations
        assert result["excel_data"], "Excel data should not be empty"
        assert result["excel_metadata"], "Excel metadata should not be empty"
        assert len(result["errors"]) == 0, f"Should have no errors, but got: {result['errors']}"
        
        print(f"\n📊 Sheet Info:")
        sheet_info = result["excel_metadata"]["sheet_info"]
        print(f"   - Rows: {sheet_info['rows']}")
        print(f"   - Columns: {sheet_info['cols']}")
        print(f"   - Sheet Name: {sheet_info['name']}")
        
        return True
    else:
        print(f"\n❌ Node 1 Test Failed!")
        print(f"Status: {result['processing_status']}")
        print(f"Errors: {result.get('errors', [])}")
        return False


def test_node_1_error_handling():
    """
    Test Node 1 error handling with invalid file path
    """
    print("\n🧪 Testing Node 1 Error Handling...")
    
    invalid_file = "nonexistent/file.xlsx"
    result = test_node_1(invalid_file, "Test error handling")
    
    # Should have errors and error status
    if result["processing_status"] == "error" and len(result["errors"]) > 0:
        print("✅ Error handling test passed!")
        return True
    else:
        print("❌ Error handling test failed!")
        return False


if __name__ == "__main__":
    print("=== Node 1 Test Suite ===\n")
    
    # Test 1: Normal operation with sample file
    test1_passed = test_node_1_with_sample_file()
    
    # Test 2: Error handling
    test2_passed = test_node_1_error_handling()
    
    # Summary
    print(f"\n=== Test Results ===")
    print(f"✅ Sample File Test: {'PASSED' if test1_passed else 'FAILED'}")
    print(f"✅ Error Handling Test: {'PASSED' if test2_passed else 'FAILED'}")
    
    if test1_passed and test2_passed:
        print(f"\n🎉 All Node 1 tests PASSED!")
        sys.exit(0)
    else:
        print(f"\n❌ Some tests FAILED!")
        sys.exit(1)
