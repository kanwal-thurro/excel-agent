# Excel Refresh Functionality with xlwings

This guide explains the new Excel refresh functionality added to the Excel Agent using xlwings for macOS compatibility.

## Overview

The agent now automatically:
- **Opens Excel files** for visual inspection at startup
- **Refreshes Excel workbooks** at each orchestration iteration
- **Keeps Excel visible** throughout the agent execution
- **Preserves Excel workbooks** open for final inspection

This functionality enables you to visually observe changes as the agent modifies your Excel files.

## Prerequisites

### 1. Install xlwings
```bash
pip install xlwings
```

### 2. Excel for Mac Setup
- Ensure you have Microsoft Excel for Mac installed
- Grant necessary permissions when prompted (accessibility permissions may be required)

### 3. Install Dependencies
```bash
pip install -r requirements.txt
```

## How It Works

### At Agent Startup
1. **File Preparation**: Agent creates a working copy of your Excel file
2. **Excel Launch**: Excel application opens with your file visible
3. **Visual Inspection**: You can see the Excel workbook throughout execution

### During Agent Execution
1. **Each Iteration**: Agent refreshes Excel workbook before processing
2. **Data Connections**: All data connections are refreshed
3. **Formulas**: All formulas are recalculated
4. **Visual Updates**: Changes are immediately visible in Excel

### At Agent Completion
1. **Final Save**: Workbook is saved with all changes
2. **Inspection**: Excel remains open for final review
3. **Manual Close**: You can close Excel when ready

## Features

### ExcelManager Class
The `ExcelManager` class handles all xlwings operations:

- `open_excel_file()`: Opens Excel file with visibility control
- `refresh_excel()`: Refreshes data connections and calculations
- `ensure_visible()`: Brings Excel to foreground for inspection
- `cleanup()`: Handles resource management while keeping Excel open

### Visual Inspection Benefits
- **Real-time Changes**: See modifications as they happen
- **Debug Assistance**: Spot issues immediately
- **Data Validation**: Verify results visually
- **Process Transparency**: Understand what the agent is doing

## Usage

### Basic Usage
```python
from agent.agent import run_excel_agent

# Run agent with Excel refresh enabled (default behavior)
result = run_excel_agent(
    excel_file_path="your_file.xlsx",
    user_question="add Q2 FY26 data",
    enable_human_intervention=False
)
```

### Testing the Functionality
```bash
# Run the test script to verify xlwings integration
python test_excel_refresh.py
```

## Configuration

### Environment Variables
The Excel refresh functionality works with existing environment variables:

```bash
# .env file
ENABLE_HUMAN_INTERVENTION=false  # Set to true for manual approval mode
USE_OLLAMA=false                 # LLM service selection
# ... other existing variables
```

### Visibility Control
Excel visibility is automatically managed, but you can modify the behavior in the `ExcelManager` class:

- `display=True`: Excel is visible (default for inspection)
- `display=False`: Excel runs in background (not recommended for inspection)

## Troubleshooting

### Common Issues

#### 1. xlwings Not Found
```bash
# Install xlwings
pip install xlwings
```

#### 2. Excel Permission Issues on macOS
- Go to System Preferences → Security & Privacy → Privacy → Accessibility
- Add your terminal application or IDE to the allowed list

#### 3. Excel Doesn't Open
- Verify Excel for Mac is installed
- Check file path is correct
- Ensure file is not already open in another Excel instance

#### 4. Refresh Fails
- Check if Excel file has data connections that need credentials
- Verify Excel formulas are valid
- Ensure file is not protected or read-only

### Debug Mode
For troubleshooting, you can run the test script:

```bash
python test_excel_refresh.py
```

This will:
1. Test ExcelManager functionality independently
2. Optionally test full agent integration
3. Provide detailed feedback on each step

## Technical Details

### Integration Points

1. **Agent Initialization**: ExcelManager is created and file is opened
2. **Orchestrator Loop**: Excel is refreshed at each iteration
3. **State Management**: ExcelManager instance is tracked in agent state
4. **Resource Cleanup**: Excel is kept open for inspection, with cleanup on errors

### Performance Considerations

- **Refresh Time**: 2-second pause after refresh to allow completion
- **Visual Display**: Excel visibility may briefly interrupt focus
- **Memory Usage**: Excel application remains in memory during execution

### Compatibility

- **macOS**: Primary target platform (Excel for Mac)
- **Windows**: Should work with Excel for Windows (untested)
- **Linux**: Not supported (requires Excel application)

## Example Output

When running the agent with Excel refresh:

```
🚀 Starting Excel Agent
📁 Original File: sample.xlsx
📋 Created Excel copy: sample_copy_20240914_143022.xlsx
📊 Opening Excel file for visual inspection...
✅ Excel file is now open for visual inspection during agent execution

🔄 === Orchestrator Iteration 1 ===
📖 Step 1: Re-parsing Excel file...
🔄 Refreshing Excel workbook for visual inspection...
✅ Excel workbook refreshed and saved
✅ Parsed Excel: 1234 characters
...

🏁 === PROCESSING COMPLETE ===
📊 Excel workbook remains open for final inspection
```

## Best Practices

1. **File Backup**: Agent creates working copies automatically
2. **Visual Monitoring**: Watch Excel during execution to spot issues
3. **Manual Review**: Inspect final results before closing Excel
4. **Permission Setup**: Configure macOS accessibility permissions beforehand
5. **Testing**: Use the test script before running important operations

## Support

If you encounter issues:

1. Run the test script to isolate problems
2. Check macOS permissions for terminal/IDE
3. Verify Excel for Mac installation
4. Review error messages in agent output
5. Check that xlwings is properly installed

## See Also

- [Agent Logic Documentation](docs/logic/agent-logic.md)
- [xlwings Documentation](https://docs.xlwings.org/)
- [Excel for Mac Support](https://support.microsoft.com/en-us/office)
