# Report Generator

A simple Python Tkinter desktop application for generating reports from search terms.

## Features

- Clean, user-friendly GUI
- Validates search terms (2-12 words)
- Generates formatted .docx reports
- Clear success/error messages
- Safe for PyInstaller packaging

## Requirements

- Python 3.7+
- tkinter (usually included with Python)

## Usage

1. Run the application:
```bash
python app.py
```

2. Enter a search term (2-12 words)
3. Click "Generate Report"
4. The application will show the path to the generated file

## Implementation

The frontend (`app.py`) handles only UI logic:
- Input validation
- Button state management
- User feedback (success/error dialogs)

The business logic (`report_logic.py`) contains:
- `generate_report(search_term: str) -> str` function
- Actual report generation implementation

## Packaging with PyInstaller

To create a standalone executable:

```bash
pip install pyinstaller
pyinstaller --onefile --windowed app.py
```

The `--windowed` flag prevents the console window from appearing.

The executable will be created in the `dist/` folder.

## File Structure

```
aplicatie-cautare/
├── app.py              # Main GUI application
├── report_logic.py     # Report generation logic
└── README.md          # This file
```

## Notes

- Reports are saved to a "Generated Reports" folder in the same directory as the application
- Filenames are timestamped automatically
- The frontend does not contain any business logic - it only calls `generate_report()`

# aplicatie-cautare
