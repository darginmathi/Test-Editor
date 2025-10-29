Build Commands: uv run main.py

# Test Case Editor Project To-Do List

## Core Application & UI
- Basic Window Structure (`MainWindow`) ✅
- Multi-Tab Interface (`QTabWidget`, `TabWidget`) ✅
- Spreadsheet View (`QTableView`, `Table`) ✅
- Menu Bar (Basic file ops) ✅
- Custom File Explorer (`FileUI`) ✅
- Styling (Basic Excel theme) ✅
- Welcome Screen ✅
- Status Bar (`QStatusBar`) ✅ - need to implement status for functions
- Zoom Bar (`QStatusBar`) ✅
## Editing Features
- Basic Cell Editing ✅
- Context Menu (Copy, Paste, Cut, Insert/Delete Row, Clear) ✅
- Shortcuts (for context menu actions) ✅
- Auto-Save on Pause (`QStyledItemDelegate`, `QTimer`) ✅
- Selection Text Color Fix (Black text on blue highlight) ✅
- **Undo/Redo** (`QUndoStack`, `QUndoCommand`) ✅ -> individual tabs ⏳
- **Find & Replace** ⏳
## Test Case Specific Features
- ComboBox ✅ - fix bug
- **Cross-file Command Retrieval** (`command_manager` logic) ✅
- Objects Retrival ✅
- **Auto-Updating Row IDs** (Logic in `ScenarioTableModel`) ✅
- **Navigate to Object Definition** (Scenario -> Object Repo jump) ⏳
- **Test Step Generation Logic** (Command-based assistance) ⏳
- **Execute Test Case** (External Java command) ⏳
- **Visual Grouping (Workaround)** (Styling/Spanning for Description based on `StartScenario`/`EndScenario`) ⏳
