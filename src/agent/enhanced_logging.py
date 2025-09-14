"""
Enhanced Logging System for Excel Agent Orchestrator

This module provides comprehensive logging capabilities for debugging the orchestrator workflow,
including state tracking, LLM response logging, and tool execution monitoring.
"""

import json
import os
from datetime import datetime
from typing import Dict, Any, List


class OrchestratorLogger:
    """
    Comprehensive logging system for orchestrator debugging
    """
    
    def __init__(self, log_dir: str = "logs"):
        """Initialize the logger with a log directory"""
        self.log_dir = log_dir
        os.makedirs(log_dir, exist_ok=True)
        
        # Create session-specific log file
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.session_log_file = os.path.join(log_dir, f"orchestrator_session_{timestamp}.json")
        
        # Initialize session log
        self.session_log = {
            "session_start": datetime.now().isoformat(),
            "iterations": []
        }
        
        print(f"📝 Logging session to: {self.session_log_file}")
    
    def log_iteration_start(self, iteration: int, state: Dict[str, Any]):
        """Log the start of an orchestrator iteration"""
        iteration_data = {
            "iteration": iteration,
            "timestamp": datetime.now().isoformat(),
            "state_snapshot": self._extract_state_snapshot(state),
            "steps": []
        }
        
        self.session_log["iterations"].append(iteration_data)
        
        print(f"\n📊 === ITERATION {iteration} STATE SNAPSHOT ===")
        print(f"📊 Status: {state.get('processing_status', 'unknown')}")
        print(f"📋 Tables identified: {len(state.get('identified_tables', []))}")
        print(f"📋 Tables processed: {len(state.get('processed_tables', []))}")
        print(f"📊 Current table index: {state.get('current_table_index', 0)}")
        print(f"🎯 Operation type: {state.get('operation_type', 'unknown')}")
        print(f"📅 Target period: {state.get('target_period', 'unknown')}")
        print(f"📝 Total cells filled: {state.get('total_cells_filled', 0)}")
        print(f"❌ Errors: {len(state.get('errors', []))}")
        print(f"⚠️  Warnings: {len(state.get('warnings', []))}")
        
        # Log current table details
        current_table = state.get('current_table', {})
        if current_table:
            print(f"📋 Current table: {current_table.get('range', 'N/A')} - {current_table.get('description', 'N/A')}")
            print(f"🌐 Global items: {current_table.get('global_items', {})}")
        
        # Log identified tables summary
        identified_tables = state.get('identified_tables', [])
        if identified_tables:
            print(f"\n📋 IDENTIFIED TABLES SUMMARY:")
            for i, table in enumerate(identified_tables):
                status = "✅ PROCESSED" if table.get('range') in state.get('processed_tables', []) else "⏳ PENDING"
                print(f"   {i}: {table.get('range', 'N/A')} - {table.get('description', 'N/A')} [{status}]")
                print(f"      Global: {table.get('global_items', {})}")
                print(f"      Needs modification: {table.get('needs_new_column', False)}")
    
    def log_excel_parsing(self, iteration: int, excel_data_length: int, excel_preview: str):
        """Log Excel parsing results"""
        current_iteration = self.session_log["iterations"][-1]
        
        step_data = {
            "step": "excel_parsing",
            "timestamp": datetime.now().isoformat(),
            "excel_data_length": excel_data_length,
            "excel_preview": excel_preview[:500] + "..." if len(excel_preview) > 500 else excel_preview
        }
        
        current_iteration["steps"].append(step_data)
        
        print(f"📖 Excel parsing: {excel_data_length} characters")
        print(f"📄 Excel preview (first 200 chars): {excel_preview[:200]}...")
    
    def log_llm_decision(self, iteration: int, llm_input: Dict[str, Any], llm_output: Dict[str, Any]):
        """Log LLM reasoning and decision"""
        current_iteration = self.session_log["iterations"][-1]
        
        step_data = {
            "step": "llm_decision",
            "timestamp": datetime.now().isoformat(),
            "input": llm_input,
            "output": llm_output
        }
        
        current_iteration["steps"].append(step_data)
        
        print(f"\n🧠 === LLM DECISION ANALYSIS ===")
        print(f"🎯 Tool selected: {llm_output.get('tool_name', 'unknown')}")
        print(f"💭 Reasoning: {llm_output.get('reasoning', 'N/A')}")
        print(f"🔧 Parameters: {json.dumps(llm_output.get('parameters', {}), indent=2)}")
        print(f"📊 Confidence: {llm_output.get('confidence', 'N/A')}")
        
        # Log input context for debugging
        print(f"\n📥 LLM INPUT CONTEXT:")
        print(f"   Status: {llm_input.get('processing_status', 'N/A')}")
        print(f"   Processed tables: {len(llm_input.get('processed_tables', []))}")
        print(f"   Total tables: {len(llm_input.get('identified_tables', []))}")
        print(f"   Excel data length: {len(llm_input.get('excel_data', ''))}")
    
    def log_human_intervention(self, iteration: int, reasoning_result: Dict[str, Any], human_decision: Dict[str, Any]):
        """Log human intervention details"""
        current_iteration = self.session_log["iterations"][-1]
        
        step_data = {
            "step": "human_intervention",
            "timestamp": datetime.now().isoformat(),
            "proposed_action": reasoning_result,
            "human_decision": human_decision
        }
        
        current_iteration["steps"].append(step_data)
        
        print(f"\n👤 === HUMAN INTERVENTION ===")
        print(f"✅ Approved: {human_decision.get('approved', False)}")
        if human_decision.get('modifications'):
            print(f"🔧 Modifications: {human_decision['modifications']}")
        if human_decision.get('reason'):
            print(f"❌ Rejection reason: {human_decision['reason']}")
    
    def log_tool_execution(self, iteration: int, tool_name: str, tool_input: Dict[str, Any], tool_output: Dict[str, Any], success: bool):
        """Log tool execution details"""
        current_iteration = self.session_log["iterations"][-1]
        
        step_data = {
            "step": "tool_execution",
            "timestamp": datetime.now().isoformat(),
            "tool_name": tool_name,
            "input": tool_input,
            "output": tool_output,
            "success": success
        }
        
        current_iteration["steps"].append(step_data)
        
        print(f"\n🔧 === TOOL EXECUTION: {tool_name} ===")
        print(f"✅ Success: {success}")
        print(f"📥 Input: {json.dumps(tool_input, indent=2)}")
        print(f"📤 Output: {json.dumps(tool_output, indent=2)}")
    
    def log_state_changes(self, iteration: int, before_state: Dict[str, Any], after_state: Dict[str, Any]):
        """Log state changes after tool execution"""
        current_iteration = self.session_log["iterations"][-1]
        
        # Extract key state changes
        changes = {}
        key_fields = ['processing_status', 'current_table_index', 'processed_tables', 'total_cells_filled', 'errors', 'warnings']
        
        for field in key_fields:
            before_val = before_state.get(field)
            after_val = after_state.get(field)
            if before_val != after_val:
                changes[field] = {"before": before_val, "after": after_val}
        
        # Check for table range changes
        before_tables = before_state.get('identified_tables', [])
        after_tables = after_state.get('identified_tables', [])
        
        if len(before_tables) == len(after_tables):
            for i, (before_table, after_table) in enumerate(zip(before_tables, after_tables)):
                if before_table.get('range') != after_table.get('range'):
                    changes[f'table_{i}_range'] = {
                        "before": before_table.get('range'),
                        "after": after_table.get('range')
                    }
        
        step_data = {
            "step": "state_changes",
            "timestamp": datetime.now().isoformat(),
            "changes": changes
        }
        
        current_iteration["steps"].append(step_data)
        
        if changes:
            print(f"\n📊 === STATE CHANGES ===")
            for field, change in changes.items():
                print(f"   {field}: {change['before']} → {change['after']}")
        else:
            print(f"\n📊 === NO STATE CHANGES ===")
    
    def log_error(self, iteration: int, error_message: str, error_details: Dict[str, Any] = None):
        """Log error information"""
        current_iteration = self.session_log["iterations"][-1]
        
        step_data = {
            "step": "error",
            "timestamp": datetime.now().isoformat(),
            "error_message": error_message,
            "error_details": error_details or {}
        }
        
        current_iteration["steps"].append(step_data)
        
        print(f"\n❌ === ERROR ===")
        print(f"🚨 Message: {error_message}")
        if error_details:
            print(f"📋 Details: {json.dumps(error_details, indent=2)}")
    
    def save_session_log(self):
        """Save the complete session log to file"""
        try:
            with open(self.session_log_file, 'w') as f:
                json.dump(self.session_log, f, indent=2, default=str)
            print(f"💾 Session log saved to: {self.session_log_file}")
        except Exception as e:
            print(f"❌ Failed to save session log: {e}")
    
    def _extract_state_snapshot(self, state: Dict[str, Any]) -> Dict[str, Any]:
        """Extract a clean snapshot of key state variables for logging"""
        return {
            "processing_status": state.get("processing_status"),
            "current_iteration": state.get("current_iteration"),
            "operation_type": state.get("operation_type"),
            "target_period": state.get("target_period"),
            "current_table_index": state.get("current_table_index"),
            "identified_tables_count": len(state.get("identified_tables", [])),
            "processed_tables_count": len(state.get("processed_tables", [])),
            "processed_tables": state.get("processed_tables", []),
            "total_cells_filled": state.get("total_cells_filled"),
            "errors_count": len(state.get("errors", [])),
            "warnings_count": len(state.get("warnings", [])),
            "excel_data_length": len(state.get("excel_data", "")),
            "human_intervention_enabled": state.get("human_intervention_enabled"),
            "current_table": state.get("current_table", {}),
            "identified_tables_summary": [
                {
                    "range": table.get("range"),
                    "description": table.get("description"),
                    "needs_new_column": table.get("needs_new_column"),
                    "global_items": table.get("global_items", {})
                }
                for table in state.get("identified_tables", [])
            ]
        }


def create_logger() -> OrchestratorLogger:
    """Create and return a new orchestrator logger instance"""
    return OrchestratorLogger()


if __name__ == "__main__":
    """Test the logging system"""
    logger = create_logger()
    
    # Test state
    test_state = {
        "processing_status": "start",
        "current_iteration": 1,
        "operation_type": "add_column",
        "target_period": "Q2 FY26",
        "identified_tables": [
            {"range": "A5:D15", "description": "Test Table", "needs_new_column": True}
        ],
        "processed_tables": [],
        "total_cells_filled": 0,
        "errors": [],
        "warnings": []
    }
    
    logger.log_iteration_start(1, test_state)
    logger.log_excel_parsing(1, 1000, "Sample Excel data...")
    logger.save_session_log()
    
    print("✅ Logging system test completed!")
