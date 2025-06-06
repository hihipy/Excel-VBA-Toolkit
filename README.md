# 📘 Excel VBA Macro Toolkit (`.bas` Modules)

This repository contains a curated set of Excel VBA macros, organized as `.bas` code modules for easy reuse and import into any Excel workbook.

These macros are designed for:
- ✅ Data cleaning
- ✅ Table exports
- ✅ PivotTable documentation
- ✅ Utility functions (e.g. hyperlink extraction, autofill)

All files are stored as **plain-text `.bas` modules**, which can be imported directly into Excel's VBA editor.

---

## 📂 What Are `.bas` Files?

`.bas` files are **VBA code modules** exported from Excel. They contain reusable procedures (macros) that can be imported into other Excel workbooks. Each `.bas` file in this repo corresponds to a specific macro or functional category (e.g., exports, utilities).

---

## 💾 How to Download Macros from GitHub

You can download and use the macros in this repository in two ways:

### Option 1: Download the Entire Toolkit
1. Click the green **`Code`** button at the top of this page.
2. Select **“Download ZIP”**.
3. Once downloaded, unzip the file on your computer.
4. Inside the unzipped folder, you’ll find organized folders such as `DataCleaning`, `Exports`, and `Utilities`, each containing `.bas` files.

### Option 2: Download an Individual Macro
1. Click on the `.bas` file you want.
2. Click the **`Raw`** button to view the plain text.
3. Right-click the page and select **“Save As…”**.
4. Ensure the file is saved with a `.bas` extension (e.g., `ExportPivotToMarkdown.bas`).

---

## 🧰 How to Enable the Developer Tab in Excel

Before you can use or import macros, the **Developer tab** must be visible in the Excel ribbon.

### 🔓 Enabling Developer Tab (Windows & Mac)

#### On Windows:
1. Open Excel.
2. Go to `File` → `Options`.
3. In the left pane, click `Customize Ribbon`.
4. On the right side, check the box labeled **Developer**.
5. Click **OK**.

#### On Mac:
1. Open Excel.
2. Go to `Excel` → `Preferences`.
3. Select `Ribbon & Toolbar`.
4. In the **Main Tabs** section, check **Developer**.
5. Click **Save**.

You will now see the **Developer** tab in the ribbon, which gives access to the VBA editor, macro tools, and import options.

---

## 🛠️ How to Use These in Excel

### Step-by-Step: Importing a `.bas` File

1. Open your Excel workbook.
2. Click the **Developer** tab, then click **Visual Basic**  
   *(or press `Alt + F11`)* to open the **VBA Editor**.
3. In the left pane (Project Explorer), click your workbook name.
4. From the top menu, go to `File` → `Import File...`.
5. Navigate to the `.bas` file you downloaded and select it.
6. The module will now appear under "Modules".
7. Press `Alt + Q` to return to Excel.
8. Press `Alt + F8`, select the macro, and click **Run**.

> 💡 You can import and use multiple macros in the same workbook.

---

## 🚀 Requirements

- Microsoft Excel (Windows or Mac)
- Macro-enabled workbook format (`.xlsm`)
- Macros must be enabled in Excel (click “Enable Content” if prompted)

---

## 📁 Repository Structure

Each folder groups macros by category:

Excel-VBA-Toolkit/
├── Data Cleaning
│   ├── Delete Hidden Rows.bas
│   └── Fill Cell from Above.bas
├── Exports
│   ├── All Excel Table Formula.bas
│   ├── Document All PivotTables (OLAP + Regular) to Markdown with Full MDX and Layout Metadata.bas
│   ├── Export Pivot Table to Markdown (AI-Friendly).bas
│   ├── Export Selected Range to Clean CSV.bas
│   └── Export Table Metadata.bas
├── Utilities
│   └── Extract URL from Hyperlink.bas
└── README.md

---

## 📌 Macro Descriptions

Brief summaries of what each `.bas` file in the toolkit does:

### 🔄 Data Cleaning

- **`Delete Hidden Rows.bas`**  
  Deletes all rows hidden by filters across all visible sheets. Ideal for prepping data before export or transformation.

- **`Fill Cell from Above.bas`**  
  Fills empty cells in a selected range with the value from the cell above. Useful for restoring grouped labels in flattened reports or pivot exports.

---

### 📤 Exports

- **`All Excel Table Formula.bas`**  
  Extracts all formulas used in Excel **ListObject tables**, generating an AI-friendly Markdown document for auditing or automation.

- **`Document All PivotTables (OLAP + Regular) to Markdown with Full MDX and Layout Metadata.bas`**  
  Scans every PivotTable (OLAP and regular) and creates a rich Markdown report including slicers, layout, data source, and MDX queries where applicable.

- **`Export Pivot Table to Markdown (AI-Friendly).bas`**  
  Exports the active sheet’s first PivotTable as a clean Markdown table — useful for GitHub, documentation, or LLM tools.

- **`Export Selected Range to Clean CSV.bas`**  
  Buffered, fast export of a selected Excel range to CSV. Handles large row counts and intelligently sanitizes delimiters and line breaks.

- **`Export Table Metadata.bas`**  
  Maps every Excel **ListObject table** in the workbook, detailing structure, column types, formulas, sample data, and inferred relationships — written in structured Markdown.

---

### 🧠 Utilities

- **`Extract URL from Hyperlink.bas`**  
  Custom function (`=URL(A1)`) that extracts the hyperlink address from a cell, useful for audits, exports, or cleaning reports.

---

## 📄 Licensing

This project is licensed under the [Creative Commons Attribution-NonCommercial-ShareAlike 4.0 International (CC BY-NC-SA 4.0)](https://creativecommons.org/licenses/by-nc-sa/4.0/) license.

### Creative Commons Attribution-NonCommercial-ShareAlike 4.0 International (CC BY-NC-SA 4.0)

#### You are free to:
- **Share** — copy and redistribute the material in any medium or format
- **Adapt** — remix, transform, and build upon the material

The licensor cannot revoke these freedoms as long as you follow the license terms.

#### Under the following terms:
- **Attribution** — You must give appropriate credit, provide a link to the license, and indicate if changes were made. You may do so in any reasonable manner, but not in any way that suggests the licensor endorses you or your use.
- **NonCommercial** — You may not use the material for commercial purposes.
- **ShareAlike** — If you remix, transform, or build upon the material, you must distribute your contributions under the same license as the original.

#### No additional restrictions — You may not apply legal terms or technological measures that legally restrict others from doing anything the license permits.

#### Notices:
You do not have to comply with the license for elements of the material in the public domain or where your use is permitted by an applicable exception or limitation.

No warranties are given. The license may not give you all of the permissions necessary for your intended use. For example, other rights such as publicity, privacy, or moral rights may limit how you use the material.

For more details, see the [LICENSE](https://creativecommons.org/licenses/by-nc-sa/4.0/legalcode) file.