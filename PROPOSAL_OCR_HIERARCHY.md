# Proposal: Structural Hierarchy & OCR Digitization for XLSX

## The Problem: Data vs. Document Structure
Current `markitdown` handles Excel files primarily as tabular data. This causes issues with many administrative and tax forms (like the French Cerfa templates) that use complex visual structures:

1. **Hierarchy Destruction**: `pd.read_excel()` treats merged cells as a single value followed by `NaN`s. The vertical parent-child relationship (e.g., a "Section Header" merging across 10 rows) is lost immediately at Stage 1.
2. **Hidden Content**: Many legacy templates implement complex headers or sections as **images**. `pandas` skips these entirely, resulting in "empty" sections in the Markdown output.

## The Solution: Option A (Pre-process Enrichment)
I am proposing an enrichment of the `XlsxConverter` that uses `openpyxl` (already a dependency) to pre-scan the workbook for structural metadata before the tabular extraction.

### Key Features
- **`detect_hierarchy` (Opt-in)**:
    - Scans `ws.merged_cells.ranges` for vertical merges in designated header columns.
    - Emits these as Markdown headers (`### Section Name`) before their respective rows.
    - This transforms a flat, broken table into a structured, hierarchical document.
- **Integrated OCR**:
    - If `paddleocr` is installed, the converter identifies embedded images.
    - It digitizes these images into text and injects them back into the proper coordinates in the DataFrame.
    - This allows "image-only" sections to be fully searchable and readable in the final Markdown.

### Implementation Details
The logic remains modular and backward-compatible:
- **Modified File**: `_xlsx_converter.py`
- **New Logic**: A pre-render scan using `openpyxl` to build a row-to-header map.
- **Rendering**: The loop now fragmentizes the sheet, alternating between headers and smaller sub-tables.

## Why this approach?
- **Minimal Impact**: Uses existing dependencies (`openpyxl`). 
- **Non-breaking**: Default behavior is unchanged.
- **Surgical Integration**: Plugs directly into the existing pandas -> HTML -> Markdown pipeline.

---
*Proposed for improved support of complex forms and templates.*
