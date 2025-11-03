# Test Case Editor

A user-friendly desktop application to edit test cases for keyword-driven testing frameworks like Selenium. No programming knowledge required!

![Test Case Editor Application](test_editor-sample-workspace.png)

![License](https://img.shields.io/badge/license-GPL-blue.svg)

## Features

* **Basic Spreadsheet Suite:** All the basic spreadsheet editing features you are used to.
* **MultiTab Functionality:** MultiTab function to actively open and edit multiple tests.
* **Custom File Explorer:** Custom built file explorer to move through project files.
* **Command Manager:** Cross-file Command(functions)/Objects retrieval with dropdown in table for ease of use.
* **Auto Update ID:** Auto Update ID (when inserting or removing rows).
* **Smart Cell Merging:** Automatically merges the 'Description' column for test case blocks between StartScenario and EndScenario commands.
* **Clean UI:** Clean and professional Dark mode UI/Light mode in the works.
* **Excel File Support:** Edit standard (`.xlsx`) test case files.

## Upcoming Features

* **Run:** Run TestCases directly from editor.
* **Light Mode:** Light theme option for different preferences.
* **Command Argument validation:** Syntax Highlighting on commands to indicate required arguments met.
* **Navigate to Object Definition** Scenario -> Object Repo jump for quick editing of objects.

## Known Issues

* **Excel Coppatability:** Some files take longer than expected to load(suspect formating further testing required).

## Installation

### Method 1: Download Ready-to-Run Executable (Recommended)

1. **Go to the Releases Page:**
   * Visit: [https://github.com/darginmathi/Test-Editor/releases](https://github.com/darginmathi/Test-Editor/releases)

2. **Download the Latest Version:**
   * Look for the latest release (e.g., "v1.0.0").
   * Under "Assets," click to download `TestEditor.exe`

3. **Run the Application:**
   * Double-click the downloaded `.exe` file
   * If Windows shows a security warning, click "More info" and then "Run anyway"

### Method 2: For Advanced Users (Building from Source)

If a pre-made .exe is not available, you can build it yourself using the instructions below.

1. **Prerequisites:**
    * [Python](https://www.python.org/)

    ```bash
    pip install uv
    uv pip install pyinstaller
    ```

2. **Build from Source:**

    ```bash
    git clone https://github.com/darginmathi/Test-Editor
    cd Test-Editor
    uv sync
    uv run pyinstaller --onefile --noconsole --name "TestEditor" main.py
    ```

    **Find Your Application:**
    After the process finishes, a new folder called `dist` will be created inside your `Test-Editor` folder. Inside `dist`, you will find `TestEditor.exe`.

## Dependencies

This project uses the following major libraries:

* [PyQt6](https://www.riverbankcomputing.com/static/Docs/PyQt6/)
* [pandas](https://pandas.pydata.org/)
* [openpyxl](https://openpyxl.readthedocs.io/en/stable/)

## Getting Started

### Setting Up

* **Files -> Open Files**
    > Select the data directory.

* **Use the built-in file explorer** to navigate to your project folder.
    > Expected Project structure:

    ```file structure
    data/
    ├── testSuits/
    │   └── <yourproject>
    │       └── Automation_Module_<your_module>.xlsx
    └── ObjectRepositories/
        └── <yourproject>
            └── ObjRep_Module_<your_module>_Test.xlsx
    ```

* **Files -> Open File**
    > Open an existing TestSuit/ObjectRepository (`.xlsx`) combo.

* **Files -> New File**
    > Create New TestSuit/ObjectRepository (`.xlsx`) combo with preset.

* **Edit -> Generate Test Cases**
    > Generate steps performed and expected result based on the command, object and value.

* Start editing with the intuitive spreadsheet interface

* Use the multi-tab feature to work on multiple test files simultaneously.

## Contributing

We welcome contributions! Please feel free to submit pull requests, report bugs, or suggest new features.

## License

This project is licensed under the GPL License - see the LICENSE file for details.
