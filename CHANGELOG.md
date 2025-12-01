# Changelog

All notable changes to PDF Manager will be documented in this file.

## [1.1.0] - 2025-12-01

### Added
- **Multi-range PDF slicing:** Extract multiple non-contiguous page ranges and merge them into a single file
- Text area for entering multiple page ranges (one per line, e.g., "1-5", "10-15")
- "Add Current Range" button for quick range entry
- "Clear Ranges" button to reset all ranges
- Support for single page extraction (e.g., just "5" instead of "5-5")
- Auto-increment feature after adding ranges for convenience
- Enhanced success message showing all extracted ranges and total pages

### Changed
- Slice PDF interface redesigned with multi-range support
- Page range validation now handles multiple ranges
- Status messages now show comprehensive extraction details

### Example Use Cases
- Extract pages 1-5, 10-15, and 20-25 in one operation
- Remove unwanted sections by extracting everything else
- Combine introduction and conclusion sections (e.g., pages 1-10, 90-100)
- Extract specific pages only (e.g., 1, 5, 10, 15)

---

## [1.0.1] - 2025-11-22

### Fixed
- File browser dialogs now open in intelligent default locations
- Enhanced action button visibility with larger size and emoji icons
- Improved user workflow by remembering file locations

### Changed
- Window size increased from 800x600 to 1600x1200
- All action buttons now use larger padding and prominent styling
- Button labels updated with emojis

---

## [1.0.0] - 2025-11-21

### Added
- Initial release of PDF Manager
- PDF Slicing feature: Extract specific page ranges from PDF files
- PDF Merging feature: Combine multiple PDF files into one
- PPTX to PDF conversion: Convert PowerPoint presentations to PDF
- Custom output path selection for all operations
- Cross-platform support (Windows, Linux, macOS)
- Tabbed interface for easy navigation
