# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.0.0-beta] - YYYY-MM-DD

### Added
- **CSV/Excel Import (機能A)**:
    - Implemented functionality to import person data from CSV and Excel (.xlsx) files.
    - Data is imported into the application's person list, replacing existing data after user confirmation.
    - Includes validation for file format and required fields.
    - User-friendly error messages are displayed for import issues.
- **Direct PDF/PNG Export (機能B)**:
    - Implemented functionality to export the relationship diagram directly to PNG image files.
    - Implemented functionality to export the relationship diagram directly to PDF documents.
    - New actions added to "File" menu and main toolbar with icons for these export options.
- **Sample Import Files**:
    - Added `sample_people.csv` and `sample_people.xlsx` to the `examples/` directory to demonstrate import format and provide test data.
- **Unit Tests for Importer**:
    - Developed unit tests for the CSV and Excel import logic (`tests/test_importer.py`) using pytest, covering success and error cases.

### Changed
- Updated `README.md` to include descriptions of new import and export features.

### Fixed
- (No specific bug fixes noted for this version as it focuses on new features)

---

## [1.0.0] - YYYY-MM-DD
(Placeholder for initial stable release - features developed prior to beta)

### Added
- Initial release of 相続関係図くん.
- Person information input form with validation.
- Family templates for quick data entry.
- Interactive relationship diagram generation (drag-and-drop nodes).
- CSV export of person data.
- Excel export of the diagram as an image.
- Save/Load project functionality to/from `.srk` (JSON) files (persists person data, diagram layout, checklist state).
- Dockable checklist for task management.
- Informational Inheritance Patterns Guide.
- Basic UI enhancements (icons, tooltips).
- `requirements.txt` and initial `README.md`.
- `LICENSE` file (MIT License).
