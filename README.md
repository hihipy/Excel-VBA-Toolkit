# Excel-VBA-Toolkit (.bas Modules)

A collection of Excel VBA macros, organized as .bas code modules for easy reuse and import into any Excel workbook.

These macros cover:
- Data cleaning and transformation
- Smart exports (CSV, Markdown) with advanced formatting
- Documentation of PivotTables and Excel Tables
- Utility functions for hyperlink extraction and data manipulation

All files are stored as plain-text .bas modules, which can be imported directly into Excel's VBA editor.

## What Are .bas Files?

.bas files are VBA code modules exported from Excel. They contain reusable procedures (macros) that can be imported into other Excel workbooks. Each .bas file in this repo corresponds to a specific macro or functional category (e.g., exports, utilities).

## How to Download Macros from GitHub

You can download and use the macros in this repository in two ways:

### Option 1: Download the Entire Toolkit
1. Click the green **Code** button at the top of this page.
2. Select "Download ZIP".
3. Once downloaded, unzip the file on your computer.
4. Inside the unzipped folder, you'll find organized folders such as DataCleaning, Exports, and Utilities, each containing .bas files.

### Option 2: Download an Individual Macro
1. Click on the .bas file you want.
2. Click the **Raw** button to view the plain text.
3. Right-click the page and select "Save As...".
4. Ensure the file is saved with a .bas extension (e.g., `ExportPivotToMarkdown.bas`).

## How to Enable the Developer Tab in Excel

Before you can use or import macros, the Developer tab must be visible in the Excel ribbon.

### Enabling Developer Tab (Windows & Mac)

**On Windows:**
1. Open Excel.
2. Go to **File > Options**.
3. In the left pane, click **Customize Ribbon**.
4. On the right side, check the box labeled **Developer**.
5. Click **OK**.

**On Mac:**
1. Open Excel.
2. Go to **Excel > Preferences**.
3. Select **Ribbon & Toolbar**.
4. In the Main Tabs section, check **Developer**.
5. Click **Save**.

You will now see the Developer tab in the ribbon, which gives access to the VBA editor, macro tools, and import options.

## How to Use These in Excel

### Step-by-Step: Importing a .bas File

1. Open your Excel workbook.
2. Click the **Developer** tab, then click **Visual Basic** (or press `Alt + F11`) to open the VBA Editor.
3. In the left pane (Project Explorer), click your workbook name.
4. From the top menu, go to **File > Import File...**.
5. Navigate to the .bas file you downloaded and select it.
6. The module will now appear under "Modules".
7. Press `Alt + Q` to return to Excel.
8. Press `Alt + F8`, select the macro, and click **Run**.

**Tip:** You can import and use multiple macros in the same workbook. For frequently used macros, consider adding them to your Personal Macro Workbook (PERSONAL.XLSB) to make them available in all Excel files.

## Requirements

- Microsoft Excel (Windows or Mac)
- Macro-enabled workbook format (.xlsm)
- Macros must be enabled in Excel (click "Enable Content" if prompted)
- For optimal performance: Excel 2016 or later recommended

## Repository Structure

Each folder groups macros by category. Below is the current structure:

```
Excel-VBA-Toolkit/
├── Data Cleaning/
│   ├── DeleteHiddenRowsOptimized.bas
│   ├── FillBlanksDown.bas
│   └── WhitespaceTools.bas
├── Exports/
│   ├── ExportPivotToMarkdown.bas
│   ├── ExportRangeToCSV.bas
│   ├── GenerateAdvancedPivotReport.bas
│   └── GenerateUniversalAITableDoc.bas
├── Utilities/
│   └── GetHyperlinkURL.bas
└── README.md
```

## Macro Descriptions

Summaries of what each .bas file in the toolkit does:

### Data Cleaning

#### DeleteHiddenRowsOptimized.bas
Deletes all hidden rows in the active worksheet using a bottom-up approach. Features real-time progress tracking, execution time reporting, and optimized performance for large datasets (50,000+ rows). Useful for cleaning filtered data before analysis or export.

#### FillBlanksDown.bas
Fills blank cells in a selected range with values from the cell directly above. Handles merged cells gracefully and provides detailed feedback on cells modified. Useful for cleaning pivot table exports or grouped data where labels are omitted in repeated rows.

#### WhitespaceTools.bas
High-performance toolkit for detecting and fixing whitespace issues across entire workbooks. Uses single-pass processing that detects and highlights leading, trailing, and multiple internal spaces in one operation. Automatically processes all sheets with optimized performance through memory optimization and bulk operations. Includes one-click workbook-wide cleaning with reporting. No selection required.

### Exports & Documentation

#### ExportPivotToMarkdown.bas
Exports the first PivotTable on the active worksheet to GitHub-compatible Markdown format. Preserves table structure with proper pipe delimiters and escapes special characters. Useful for documentation, reports, or sharing pivot analysis in markdown-friendly platforms.

#### ExportRangeToCSV.bas
CSV export tool with intelligent data type detection, configurable text quoting, and buffered writing for performance. Handles 50,000+ rows efficiently while preserving data integrity. Supports custom delimiters and provides detailed export statistics.

#### GenerateAdvancedPivotReport.bas
Creates documentation for all PivotTables (both OLAP and regular) in a workbook. Includes field configurations, data sources, OLAP connection details, MDX references, calculated fields, and slicer information. Outputs Markdown reports with complete metadata analysis.

#### GenerateUniversalAITableDoc.bas
Creates documentation of all Excel Tables across every worksheet in the workbook. Features data profiling with intelligent data type detection, sample values, formula transparency with exact syntax and dependency mapping, data quality assessment (CLEAN/WARNING/ERROR flags), and performance optimization guidance. Generates Markdown output designed for feeding to AI tools for Excel analysis, formula generation, and automated data manipulation. Works with datasets from 100 to 100,000+ rows.

### Utilities

#### GetHyperlinkURL.bas
Custom Excel function that extracts the actual URL from hyperlinked cells. Use with `=GetHyperlinkURL(A1)` to retrieve hyperlink addresses for link inventories, validation, or export lists. Includes error handling for non-hyperlinked cells and multiple cell selections.

## License

This project is licensed under [CC BY-NC-SA 4.0](https://creativecommons.org/licenses/by-nc-sa/4.0/).

You are free to:
- Use, share, and adapt this work
- Use it at your job

Under these terms:
- **Attribution** — Credit the original author
- **NonCommercial** — No selling or commercial products
- **ShareAlike** — Derivatives must use the same license
