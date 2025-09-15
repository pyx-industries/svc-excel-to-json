# svc-excel-to-json

A tool to convert an SVC workbook (Excel) into an SVC JSON Instance based on the [UNECE Sustainability Vocabulary](https://jargon.sh/user/unece/SustainabilityVocabulary).

**Current supported schema version:** `0.5.0`

---

## Getting Started

### Build and Run Locally

1. Create and activate a Python virtual environment:
   ```bash
   python -m venv .venv
   source .venv/bin/activate   # Linux / macOS
   .venv\Scripts\activate      # Windows
   ```

2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

3. Build the standalone executable with PyInstaller:
   ```bash
   pyinstaller --onefile --name excel2jsonsvc ./code/xlsx_to_json_svc.py
   ```
   - The executable will be created in the `./dist` folder.
   - Build artifacts will be placed in the `./build` folder.

4. Run the generated executable:
   - On Linux / macOS:
     ```bash
     ./dist/excel2jsonsvc
     ```
   - On Windows:
     ```powershell
     dist\excel2jsonsvc.exe
     ```

---

## Notes
- The input SVC workbook **must not** contain any tag filters.
- Input and output paths can be relative or absolute.
- Example input workbooks are available in the `files/` directory.

---

## Development
- Main code: `./code/xlsx_to_json_svc.py`
- Archived versions: `./archive/`
- To run the script directly (without building an executable):
  ```bash
  python ./code/xlsx_to_json_svc.py
  ```

---

## Repository Structure
```
.
├── archive/               # Older/archived scripts
├── build/                 # PyInstaller build artifacts (ignored in git)
├── code/                  # Main source code
│   └── xlsx_to_json_svc.py
├── dist/                  # Built executables (ignored in git)
├── files/                 # Example input/output workbooks & JSONs
└── README.md              # Project documentation
```

---

## License
[MIT](LICENSE) (or whichever license you choose)