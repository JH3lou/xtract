# TaxOpt Converter

A Python application to extract data from ZIP files into Excel and convert Excel templates into ZIP files with structured text files.

## Features
✅ Extract pipe-delimited TXT files from a ZIP and convert them to Excel  
✅ Convert Excel sheets into structured TXT files and package them into a ZIP  
✅ Custom file naming with timestamps and unique IDs  
✅ User-friendly UI with two-tab navigation  

## Installation
```sh
''' git clone https://github.com/your-repo/TaxOptConverter.git '''
cd TaxOptConverter
python -m venv venv
source venv/bin/activate  # macOS/Linux
venv\Scripts\activate     # Windows
pip install -r requirements.txt
```
---
# *xtract* - A Dynamic Template-Driven ZIP to Excel Converter

###### SPEC-001 - 04/10/2025

## Table of Contents

1. [Background](#background)
2. [Requirements](#requirements)
   - [Must Have](#must-have)
   - [Should Have](#should-have)
   - [Could Have](#could-have)
   - [Won’t Have (For MVP)](#wont-have-for-mvp)
3. [Method](#method)
   - [Architecture Overview](#architecture-overview)
   - [Template JSON Schema](#template-json-schema)
   - [Token Engine](#token-engine)
   - [User Interface Architecture](#user-interface-architecture)
   - [File Conversion Logic](#file-conversion-logic)
   - [Project Structure](#project-structure)
4. [Implementation](#implementation)
5. [Milestones](#milestones)
6. [Gathering Results](#gathering-results)

---

## Background

The **Xtract File Converter** began as a Python-based tool for converting financial data between ZIP-packed TXT files and structured Excel formats. Initially designed for a fixed format, evolving demands required a more adaptable approach.
This enhanced version supports **template-driven transformations**, empowering non-technical users to define file parsing rules, headers, delimiters, and more — all without code changes. Additional features include:

- Real-time feedback and logging
- Loading screens for long processes
- Two-tab modern UI
- Reversible conversions between TXT and Excel

---

## Requirements

### Must Have

- Template-based configuration for both ZIP ➝ Excel and Excel ➝ ZIP
- GUI for creating/editing/deleting templates
- Store template metadata as:
- `/settings_configs/{template_name}.json`
- `/templates/{template_name}_template.xlsx`
- Configurable:
- Header/trailer formats with tokens
- Delimiters
- Tokenized file naming
- Header mappings
- Skip rules (regex)
- Template selection during conversion
- Error handling:
- Skip invalid files, log to error tab
- JSON validation
- Threaded operations to prevent UI freeze
- Loading screen during conversions
- ZIP packaging/unpacking with structured Excel conversion

### Should Have

- Token resolution system for built-in and custom tokens
- File previews based on template
- File type filtering (.txt, .csv, .dat)
- Persistent logs
- Header/trailer row logic
- Excel template upload UI

### Could Have

- Centralized `config.json` index of templates
- Custom transformer plugin system
- GUI-based regex rule builder
- Conditional trailer logic (e.g., column sums)
- GitHub Actions CI for pull request testing

### Won’t Have (For MVP)

- Cloud storage/database integration
- Real-time collaboration or template sharing

---

## Method
xtract_uml_diagrams/Architecture Overview UML.svg
### Architecture Overview
<img src="./xtract_uml_diagrams/Architecture Overview UML.svg">

**Modules:**

- `App UI`: Two-tab interface
- `TemplateManager`: Create/edit/validate templates
- `TemplateStore`: Reads/writes templates
- `FileConverter`: Runs conversions
- `ExcelToZipConverter`: Converts Excel → TXT + ZIP
- `ZipToExcelExtractor`: Converts ZIP → Excel
- `TokenEngine`: Resolves tokens
- `LoadingScreen`: Shows spinner overlay

---

### Template JSON Schema

Stored under `settings_configs/{template_name}.json`

```json
{
 "header_format": "HDR|LPB|{timestamp}|{file_name}|{random_6_digits}",
 "trailer_format": "TLR|{row_count}|End",
 "file_naming_pattern": "TAXOPT.{timestamp}.{format_name}.{file_name}.txt",
 "delimiter": "|",
 "header_mappings": {
   "MODALLOC": "ModelAllocation",
   "POSITION": "Position"
 },
 "excel_to_zip": {
   "include_header_row": true,
   "include_trailer_row": true
 },
 "zip_to_excel": {
   "delimiter": "|",
   "has_header": true,
   "has_trailer": true,
   "sheet_map": {
     "ModelAllocation": "MODALLOC",
     "Position": "POSITION"
   },
   "skip_patterns": ["^HDR\|", "^TLR\|"]
 }
}
```

---

### Token Engine

Handles token replacement using context and template configuration.

#### Built-in Tokens

| Token                       | Description                          | Example        |
| --------------------------- | ------------------------------------ | -------------- |
| `{timestamp}`             | Current datetime                     | 20250410121800 |
| `{file_name}`             | Logical file name                    | AccountMaster  |
| `{sheet_name}`            | Excel sheet name                     | MODALLOC       |
| `{row_count}`             | Row count (excluding header/trailer) | 134            |
| `{random_6_digits}`       | Random 6-digit number                | 382104         |
| `{random_5_alphanumeric}` | Mixed alphanumeric                   | A9T3Z          |
| `{random_L5N}`            | 1 letter + 5 digits                  | G39284         |

#### Custom Token Support

```json
"custom_tokens": {
 "{rand_alpha_4}": {
   "length": 4,
   "type": "alpha"
 },
 "{rand_num_3}": {
   "length": 3,
   "type": "numeric"
 },
 "{rand_alphanum_8}": {
   "length": 8,
   "type": "alphanumeric"
 }
}
```

#### Replacement API

```python
class TokenEngine:
   def resolve_tokens(self, text: str, context: dict, template_config: dict) -> str
```

---
### User Interface Architecture
<img src="./xtract_uml_diagrams/User Interface Architecture.svg">


#### Tabs

- **Extract Tab**: ZIP → Excel
- **Convert Tab**: Excel → ZIP
- **Settings Panel**: Manage templates
- **File Preview**: Simulated outputs
- **Loading Overlay**: Spinner UI
- **Error Log Tab**: Persistent logging

---

### File Conversion Logic

#### Excel → ZIP

##### Basic Flow
<img src="./xtract_uml_diagrams/Excel 2 ZIP Basic Flow.svg">


##### Expanded User Journey Diagram
<img src="./xtract_uml_diagrams/Excel 2 ZIP Expanded User Journey Diagram.svg">


#### ZIP → Excel

##### Basic Flow
<img src="./xtract_uml_diagrams/ZIP 2 Excel Basic Flow.svg">


##### Expanded User Journey Diagram
<img src="./xtract_uml_diagrams/ZIP 2 Excel Expanded User Journey Diagram.svg">



#### Reversible Round-Trip

- One template supports both directions
- Logical name mappings ensure consistency

---

### Project Structure

```
xtract_file_converter/
├── app/
│   ├── ui/
│   ├── core/
│   ├── app.py
│   └── __init__.py
├── templates/
├── settings_configs/
├── assets/
├── tests/
├── requirements.txt
├── setup.py
├── .gitignore
└── README.md
```

---

## Implementation

### Phase 1: Core Refactor & Template Engine

- Refactor existing TXT ↔ Excel logic into:
- `excel_to_zip.py`
- `zip_to_excel.py`
- Implement `TokenEngine` with built-in token resolution
- Add support for `custom_tokens` from JSON
- Define schema loader for reading and validating template JSONs
- Create starter templates under `settings_configs/` and `templates/`

### Phase 2: UI Shell & File Operations

- Build two-tab UI layout (`ExtractTab`, `ConvertTab`)
- Implement file selection + template dropdown
- Wire up ZIP ➝ Excel and Excel ➝ ZIP logic from UI
- Integrate threaded execution with `QThread` or equivalent
- Add loading overlay with spinner (non-blocking)

### Phase 3: Settings Panel & Template Management

- Implement `SettingsPanel` UI
- Dropdown for existing templates
- Form for editing header/trailer, delimiter, token pattern
- File browser to upload `.xlsx` template
- Build `TemplateManager` service for CRUD actions
- Validate template JSON structure on save

### Phase 4: UX Enhancements

- Implement preview panel for TXT and Excel output
- Add persistent error log panel with scroll/save/export options
- Support regex-based skip rules for TXT parsing

### Phase 5: Token Extensions & Unit Testing

- Add UI helper for defining `custom_tokens`
- Extend `TokenEngine` to support all token types
- Write unit tests for:
- `token_engine.py`
- `excel_to_zip.py`
- `zip_to_excel.py`
- `template_manager.py`
- Add test templates and test data files

### Phase 6: Packaging & Deployment

- Organize `setup.py` for distribution
- Add `requirements.txt` for dependency installation
- Create `.gitignore` and GitHub Actions CI for lint + test

---

## Milestones

| Milestone    | Deliverables                                    |
| ------------ | ----------------------------------------------- |
| **M1** | Refactored logic, TokenEngine, static templates |
| **M2** | UI tabs, threading, loading screen              |
| **M3** | Full Settings UI, template CRUD                 |
| **M4** | Preview system, error tab, filtering            |
| **M5** | Token builder UI, testing coverage              |
| **M6** | Packaging, sample data, CI integration          |

---

## Gathering Results

### Functional Validation

- Templates enable accurate, bidirectional transformation between Excel and ZIP formats.
- Header/trailer formats, delimiters, filenames, and sheet mappings follow the configured logic.
- Custom tokens produce expected values with appropriate formats.
  **Validation Method:**
- Use multiple templates with varying formats
- Perform round-trip tests (ZIP ➝ Excel ➝ ZIP) and compare output integrity
- Use preview tool to pre-validate formatting and filenames

### UX & UI Feedback

- Application remains responsive during long operations (no freezing).
- Users can create and manage templates without editing raw JSON files.
- Error messages appear in a dedicated tab, not via blocking popups.
  **Validation Method:**
- Internal user testing (ops team, QA)
- Track time taken for common flows (conversion, template edits)
- Collect UI feedback via user surveys or review sessions

### Robustness & Edge Handling

- System handles:
- Missing or malformed files
- Invalid templates
- Unmapped sheet names or token errors
- Errors are captured in a central log, and application does not crash
  **Validation Method:**
- Fuzz test with corrupted and edge-case inputs
- Review generated error logs for clarity and completeness

### Extensibility & Maintainability

- New templates can be added without code changes.
- Logic is modular and testable.
- Developers can add new token types or transformation plugins easily.
  **Validation Method:**
- Add a net-new format (e.g., CSV-based extract, new firm) via config only
- Assess code coverage and modularity
- Confirm CI pipeline passes on push