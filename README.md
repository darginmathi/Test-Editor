# Test Case Editor

A custom TestCase Editor made to edit test cases for a "Keyword-driven development Testing framework".

## Features

*   **MultiTab:** MultiTab function to actively open and edit multiple files.
*   **Custom FileExplorer:** Custom builf file exlorer to move through project files.
*   **Edit Test Cases:** Edit excel files (`.xlsx`).

## Upcomming Features
*   **Command Manager:** Cross file command retrieval.
*   **UpdateID:** Custom Update ID on the go.
*   **Run:** Run TestCases directly from editor.

## Installation

**Clone the repository:**
```bash
git clone https://github.com/darginmathi/Test-Editor
cd TestEditor
uv sync
```

**Install `uv` (if uv not installed "Linux Cmd"):**
```bash
pip install uv
```

## Usage

To run the application, execute the following command from the project's root directory:

```bash
uv run main.py
```

## Dependencies

This project uses the following major libraries:

*   [PyQt6](https://www.riverbankcomputing.com/static/Docs/PyQt6/)
*   [pandas](https://pandas.pydata.org/)
*   [openpyxl](https://openpyxl.readthedocs.io/en/stable/)

## License

This project is licensed under the GPL License.
