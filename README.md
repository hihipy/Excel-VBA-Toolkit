# 📘 Excel-VBA-Toolkit (.bas Modules)

This repository contains a curated set of Excel VBA macros, organized as .bas code modules for easy reuse and import into any Excel workbook.

These macros are designed for:
- ✅ Data cleaning and transformation
- ✅ Smart exports (CSV, Markdown) with advanced formatting
- ✅ Comprehensive documentation of PivotTables and Excel Tables
- ✅ Utility functions for hyperlink extraction and data manipulation

All files are stored as plain-text .bas modules, which can be imported directly into Excel's VBA editor.

## 📂 What Are .bas Files?

.bas files are VBA code modules exported from Excel. They contain reusable procedures (macros) that can be imported into other Excel workbooks. Each .bas file in this repo corresponds to a specific macro or functional category (e.g., exports, utilities).

## 💾 How to Download Macros from GitHub

You can download and use the macros in this repository in two ways:

### Option 1: Download the Entire Toolkit
1. Click the green **Code** button at the top of this page.
2. Select "Download ZIP".
3. Once downloaded, unzip the file on your computer.
4. Inside the unzipped folder, you'll find organized folders such as DataCleaning, Exports, and Utilities, each containing .bas files.

### Option 2: Download an Individual Macro
1. Click on the .bas file you want.
2. Click the **Raw** button to view the plain text.
3. Right-click the page and select "Save As…".
4. Ensure the file is saved with a .bas extension (e.g., `ExportPivotToMarkdown.bas`).

## 🧰 How to Enable the Developer Tab in Excel

Before you can use or import macros, the Developer tab must be visible in the Excel ribbon.

### 🔓 Enabling Developer Tab (Windows & Mac)

**On Windows:**
1. Open Excel.
2. Go to **File → Options**.
3. In the left pane, click **Customize Ribbon**.
4. On the right side, check the box labeled **Developer**.
5. Click **OK**.

**On Mac:**
1. Open Excel.
2. Go to **Excel → Preferences**.
3. Select **Ribbon & Toolbar**.
4. In the Main Tabs section, check **Developer**.
5. Click **Save**.

You will now see the Developer tab in the ribbon, which gives access to the VBA editor, macro tools, and import options.

## 🛠️ How to Use These in Excel

### Step-by-Step: Importing a .bas File

1. Open your Excel workbook.
2. Click the **Developer** tab, then click **Visual Basic** (or press `Alt + F11`) to open the VBA Editor.
3. In the left pane (Project Explorer), click your workbook name.
4. From the top menu, go to **File → Import File...**.
5. Navigate to the .bas file you downloaded and select it.
6. The module will now appear under "Modules".
7. Press `Alt + Q` to return to Excel.
8. Press `Alt + F8`, select the macro, and click **Run**.

💡 **Pro Tip:** You can import and use multiple macros in the same workbook. For frequently used macros, consider adding them to your Personal Macro Workbook (PERSONAL.XLSB) to make them available in all Excel files.

## 🚀 Requirements

- Microsoft Excel (Windows or Mac)
- Macro-enabled workbook format (.xlsm)
- Macros must be enabled in Excel (click "Enable Content" if prompted)
- For optimal performance: Excel 2016 or later recommended

## 📁 Repository Structure

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

## 📌 Macro Descriptions

Detailed summaries of what each .bas file in the toolkit does:

### 🧹 Data Cleaning

#### **DeleteHiddenRowsOptimized.bas**
Efficiently deletes all hidden rows in the active worksheet using a bottom-up approach. Features real-time progress tracking, execution time reporting, and optimized performance for large datasets (50,000+ rows). Perfect for cleaning filtered data before analysis or export.

#### **FillBlanksDown.bas**
Intelligently fills blank cells in a selected range with values from the cell directly above. Handles merged cells gracefully and provides detailed feedback on cells modified. Essential for cleaning pivot table exports or grouped data where labels are omitted in repeated rows.

#### **WhitespaceTools.bas**
Ultra-high-performance toolkit for detecting and fixing whitespace issues across entire workbooks. Features advanced single-pass processing that detects AND highlights leading, trailing, and multiple internal spaces in one operation. Automatically processes all sheets with 50-80% faster performance than traditional methods through memory optimization and bulk operations. Includes one-click workbook-wide cleaning with comprehensive reporting. Essential for ensuring data integrity across large datasets before performing lookups, analysis, or exports. No selection required - simply run and let it optimize your entire workbook.

### 📊 Exports & Documentation

#### **ExportPivotToMarkdown.bas**
Exports the first PivotTable on the active worksheet to GitHub-compatible Markdown format. Preserves table structure with proper pipe delimiters and escapes special characters. Perfect for documentation, reports, or sharing pivot analysis in markdown-friendly platforms.

#### **ExportRangeToCSV.bas**
Advanced CSV export tool with intelligent data type detection, configurable text quoting, and buffered writing for optimal performance. Handles 50,000+ rows efficiently while preserving data integrity. Supports custom delimiters and provides detailed export statistics.

#### **GenerateAdvancedPivotReport.bas**
Creates comprehensive documentation for all PivotTables (both OLAP and regular) in a workbook. Includes field configurations, data sources, OLAP connection details, MDX references, calculated fields, and slicer information. Outputs detailed Markdown reports with complete metadata analysis.

#### **GenerateUniversalAITableDoc.bas**
Creates enterprise-grade, AI-optimized documentation of all Excel Tables across every worksheet in the workbook. Features robust data profiling with intelligent data type detection, sample values, complete formula transparency with exact syntax and dependency mapping, comprehensive data quality assessment (CLEAN/WARNING/ERROR flags), and performance optimization guidance. Generates clean, text-only Markdown output specifically designed for feeding to AI tools for advanced Excel analysis, formula generation, and automated data manipulation. Universal compatibility across all industries (financial, healthcare, manufacturing, retail) with scalable performance for datasets from 100 to 100,000+ rows. Zero business logic assumptions - provides pure structural analysis suitable for any Excel table structure. Essential for creating professional data dictionaries and enabling sophisticated AI-assisted coding workflows.

### 🔧 Utilities

#### **GetHyperlinkURL.bas**
Custom Excel function that extracts the actual URL from hyperlinked cells. Use with `=GetHyperlinkURL(A1)` to retrieve hyperlink addresses for link inventories, validation, or export lists. Includes error handling for non-hyperlinked cells and multiple cell selections.

## 📄 Licensing

This project is licensed under the **Creative Commons Attribution-NonCommercial-ShareAlike 4.0 International (CC BY-NC-SA 4.0)** license.

[![Creative Commons License](https://i.creativecommons.org/l/by-nc-sa/4.0/88x31.png)](http://creativecommons.org/licenses/by-nc-sa/4.0/)

**You are free to:**
- **Share** — copy and redistribute the material in any medium or format
- **Adapt** — remix, transform, and build upon the material

The licensor cannot revoke these freedoms as long as you follow the license terms.

**Under the following terms:**
- **Attribution** — You must give appropriate credit, provide a link to the license, and indicate if changes were made. You may do so in any reasonable manner, but not in any way that suggests the licensor endorses you or your use.
- **NonCommercial** — You may not use the material for commercial purposes.
- **ShareAlike** — If you remix, transform, or build upon the material, you must distribute your contributions under the same license as the original.
- **No additional restrictions** — You may not apply legal terms or technological measures that legally restrict others from doing anything the license permits.

**Notices:**
- You do not have to comply with the license for elements of the material in the public domain or where your use is permitted by an applicable exception or limitation.
- No warranties are given. The license may not give you all of the permissions necessary for your intended use. For example, other rights such as publicity, privacy, or moral rights may limit how you use the material.

For more details, see the [LICENSE](LICENSE) file.