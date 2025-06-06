' ==========================================================================================
' 📌 Macro: DocumentAllTables
' 📁 Module Purpose:
'     Creates a comprehensive Markdown-style documentation of **all Excel Tables (ListObjects)**
'     in the current workbook — ideal for auditing, AI analysis, or developer metadata mapping.
'
' ------------------------------------------------------------------------------------------
' ✅ Key Features:
'     - Scans every worksheet and every table (ListObject)
'     - Outputs detailed metadata including:
'         • Table name, range, header/body, anchor cell
'         • Field names, datatypes, sample values
'         • Formula fields with Excel syntax and human-readable descriptions
'         • Potential primary keys and relationships
'         • Suggested formula types for each column
'     - Outputs a GitHub- and AI-compatible `.txt` file in Markdown style
'
' ------------------------------------------------------------------------------------------
' 🔍 Core Behaviors:
'     - Skips hidden sheets unless explicitly included
'     - Automatically abbreviates long field values
'     - Uses helper functions to infer types, formulas, and relationships
'     - Prompts the user to save the output `.txt` file
'     - Generates structured Markdown output ready for LLM parsing or human review
'
' ------------------------------------------------------------------------------------------
' 📁 Sample Output Format (Table Excerpt):
'     # TABLE_DEFINITION: Grants_Table
'     Worksheet: Financials
'     TableRange: A1:H500
'     ColumnCount: 8
'     HasHeaders: Yes
'     HasTotals: No
'
'     ## COLUMN_DEFINITIONS
'     | ColumnIndex | ColumnName | DataType | PotentialKey | FormulaSuggestion | RangeAddress |
'     |--------------|-------------|-----------|--------------|------------------|---------------|
'     | 1 | PI Name | Text | Yes | VLOOKUP | A2:A500 |
'     | 2 | Grant_ID | Text | Yes | VLOOKUP | B2:B500 |
'
' ------------------------------------------------------------------------------------------
' 🧠 Use Cases:
'     - Documenting tables across large, multi-sheet workbooks
'     - Generating data dictionaries for analytics or Power BI projects
'     - Allowing GPT-like tools to build formulas or interpret data programmatically
'     - Quality-checking Excel table structure before data pipeline ingestion
'
' ------------------------------------------------------------------------------------------
' 🛠️ Output Location:
'     Defaults to: `C:\Users\<YourUsername>\Downloads\Table_FieldMap_Combined.txt`
'
' ==========================================================================================
Sub DocumentAllTables()
    ' Purpose: Creates a comprehensive mapping of all tables (ListObjects) in a workbook
    ' Output: Markdown-formatted text file for AI and documentation purposes

    Dim ws As Worksheet, tbl As ListObject, col As ListColumn
    Dim strFilePath As String, fileNum As Integer
    Dim anchorCell As String, startTime As Double
    Dim totalTablesCount As Long, tableCount As Long, totalTables As Long
    Dim skipTable As Boolean

    On Error GoTo ErrorHandler
    startTime = Timer

    strFilePath = ChooseOutputLocation("Table_FieldMap_Combined.txt")
    If strFilePath = "" Then MsgBox "Operation cancelled by user.", vbInformation: Exit Sub

    Application.StatusBar = "Initializing table mapping..."
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    fileNum = FreeFile
    On Error Resume Next
    Open strFilePath For Output As #fileNum
    If Err.Number <> 0 Then MsgBox "Could not open file: " & Err.Description, vbCritical: GoTo CleanExit
    Err.Clear
    On Error GoTo ErrorHandler

    ' Write header
    Print #fileNum, "# Excel Table Field Mapping"
    Print #fileNum, "Generated: " & Format(Now(), "yyyy-mm-dd hh:mm:ss")
    Print #fileNum, "Workbook: " & ThisWorkbook.Name & vbNewLine

    totalTables = 0
    For Each ws In ThisWorkbook.Worksheets
        totalTables = totalTables + ws.ListObjects.Count
    Next ws

    For Each ws In ThisWorkbook.Worksheets
        tableCount = 0
        For Each tbl In ws.ListObjects
            skipTable = False
            tableCount = tableCount + 1
            totalTablesCount = totalTablesCount + 1

            If ws.Visible <> xlSheetVisible Then skipTable = True

            Application.StatusBar = "Processing table " & totalTablesCount & " of " & totalTables & ": " & tbl.Name

            anchorCell = tbl.Range.Cells(1, 1).Address

            Print #fileNum, "# TABLE_DEFINITION: " & tbl.Name
            Print #fileNum, "Worksheet: " & ws.Name & IIf(skipTable, " (Hidden)", "")
            Print #fileNum, "SourceType: " & GetSourceTypeName(tbl.SourceType)
            Print #fileNum, "AnchorCell: " & anchorCell
            Print #fileNum, "TableRange: " & tbl.Range.Address(False, False)
            Print #fileNum, "HeadersRange: " & tbl.HeaderRowRange.Address(False, False)
            Print #fileNum, "DataBodyRange: " & IIf(tbl.DataBodyRange Is Nothing, "N/A", tbl.DataBodyRange.Address(False, False))
            Print #fileNum, "RowCount: " & GetRowCount(tbl)
            Print #fileNum, "ColumnCount: " & tbl.ListColumns.Count
            Print #fileNum, "HasHeaders: " & IIf(tbl.ShowHeaders, "Yes", "No")
            Print #fileNum, "HasTotals: " & IIf(tbl.ShowTotals, "Yes", "No")
            Print #fileNum, ""

            Print #fileNum, "## COLUMN_DEFINITIONS"
            Print #fileNum, "| ColumnIndex | ColumnName | DataType | PotentialKey | FormulaSuggestion | RangeAddress |"
            Print #fileNum, "|--------------|-------------|-----------|--------------|------------------|---------------|"

            For Each col In tbl.ListColumns
                Print #fileNum, "| " & col.Index & " | " & SafeText(col.Name) & " | " & GetColumnDataType(col) & _
                    " | " & IsPotentialKeyField(col.Name) & " | " & SuggestFormula(col.Name, GetColumnDataType(col)) & _
                    " | " & IIf(col.DataBodyRange Is Nothing, "N/A", col.DataBodyRange.Address(False, False)) & " |"
            Next col

            Print #fileNum, vbNewLine & "## SAMPLE_DATA" & vbNewLine & "### First Row"
            Print #fileNum, "| ColumnName | Value | ExcelAddress |"
            Print #fileNum, "|-------------|-------|--------------|"

            If Not tbl.DataBodyRange Is Nothing Then
                For Each col In tbl.ListColumns
                    Print #fileNum, "| " & SafeText(col.Name) & " | " & SafeText(GetCellValue(tbl.DataBodyRange.Cells(1, col.Index))) & _
                        " | " & tbl.DataBodyRange.Cells(1, col.Index).Address & " |"
                Next col
            Else
                Print #fileNum, "| *No data rows* | - | - |"
            End If

            Print #fileNum, vbNewLine & "## CALCULATED_COLUMNS"
            Print #fileNum, "| ColumnName | Formula | Description |"
            Print #fileNum, "|------------|---------|-------------|"
            Dim hasFormula As Boolean: hasFormula = False
            For Each col In tbl.ListColumns
                If Not col.DataBodyRange Is Nothing Then
                    If col.DataBodyRange.Cells(1).HasFormula Then
                        hasFormula = True
                        Print #fileNum, "| " & SafeText(col.Name) & " | " & SafeText(col.DataBodyRange.Cells(1).Formula) & _
                            " | " & GetFormulaDescription(col.DataBodyRange.Cells(1).Formula) & " |"
                    End If
                End If
            Next col
            If Not hasFormula Then Print #fileNum, "| None | N/A | No calculated columns in this table |"

            Print #fileNum, vbNewLine & "## POTENTIAL_RELATIONSHIPS"
            Dim hasRel As Boolean: hasRel = False
            For Each col In tbl.ListColumns
                If IsPotentialKeyField(col.Name) = "Yes" Then
                    hasRel = True
                    Print #fileNum, "- " & col.Name & " could be used to relate to other tables"
                End If
            Next col
            If Not hasRel Then Print #fileNum, "- No obvious relationship keys detected"

            Print #fileNum, vbNewLine & "## FORMULA_HINTS"
            Print #fileNum, "- " & tbl.Name & "[[Column1]:[Column2]]"
            Print #fileNum, "- Use SUMIFS, AVERAGEIFS, or INDEX/MATCH for lookups" & vbNewLine & "---" & vbNewLine
        Next tbl
    Next ws

    ' Final summary
    Print #fileNum, "# SUMMARY"
    Print #fileNum, "TotalTables: " & totalTablesCount
    Print #fileNum, "ProcessingTime: " & Format(Timer - startTime, "0.00") & " seconds"

    Close #fileNum

CleanExit:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

    If totalTablesCount > 0 Then
        MsgBox "Saved to: " & strFilePath & vbNewLine & _
               "✓ Tables: " & totalTablesCount & vbNewLine & _
               "✓ Time: " & Format(Timer - startTime, "0.00") & " sec", vbInformation
    Else
        MsgBox "No tables found in this workbook.", vbExclamation
    End If
    Exit Sub

ErrorHandler:
    On Error Resume Next
    Close #fileNum
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Macro Error"
    Resume CleanExit
End Sub

' Prompt user for where to save the output file
Function ChooseOutputLocation(defaultFileName As String) As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogSaveAs)
    With fd
        .Title = "Save Table Field Map"
        .InitialFileName = Environ("USERPROFILE") & "\Downloads\" & defaultFileName
        .Filters.Clear
        .Filters.Add "Text Files", "*.txt"
        If .Show = -1 Then ChooseOutputLocation = .SelectedItems(1)
    End With
End Function

' Safely sanitize text for markdown
Function SafeText(inputText As Variant) As String
    If IsError(inputText) Or IsEmpty(inputText) Then
        SafeText = "(empty)"
    Else
        SafeText = Replace(Replace(Replace(Replace(CStr(inputText), "|", "\|"), vbCr, " "), vbLf, " "), vbTab, " ")
        If Len(SafeText) > 100 Then SafeText = Left(SafeText, 97) & "..."
    End If
End Function

' Infer data type based on first cell
Function GetColumnDataType(col As ListColumn) As String
    On Error Resume Next
    If col.DataBodyRange Is Nothing Then GetColumnDataType = "Unknown": Exit Function
    Dim v As Variant: v = col.DataBodyRange.Cells(1, 1).Value
    Select Case True
        Case IsNumeric(v): GetColumnDataType = "Numeric"
        Case IsDate(v): GetColumnDataType = "Date"
        Case v = "": GetColumnDataType = "Empty"
        Case Else: GetColumnDataType = "Text"
    End Select
End Function

Function GetRowCount(tbl As ListObject) As Long
    If tbl.DataBodyRange Is Nothing Then GetRowCount = 0 Else GetRowCount = tbl.DataBodyRange.Rows.Count
End Function

Function GetSourceTypeName(sourceType As XlListObjectSourceType) As String
    Select Case sourceType
        Case xlSrcRange: GetSourceTypeName = "Range"
        Case xlSrcExternal: GetSourceTypeName = "External"
        Case xlSrcQuery: GetSourceTypeName = "Query"
        Case Else: GetSourceTypeName = "Other"
    End Select
End Function

Function IsPotentialKeyField(colName As String) As String
    colName = LCase(colName)
    If colName Like "*id*" Or colName Like "*key*" Or colName Like "*name*" Then
        IsPotentialKeyField = "Yes"
    Else
        IsPotentialKeyField = "No"
    End If
End Function

Function SuggestFormula(colName As String, dataType As String) As String
    If dataType = "Numeric" Then
        SuggestFormula = "SUMIFS or AVERAGEIFS"
    ElseIf colName Like "*date*" Then
        SuggestFormula = "DATEDIF, TODAY(), or YEAR()"
    Else
        SuggestFormula = "VLOOKUP or MATCH"
    End If
End Function

Function GetCellValue(cell As Range) As String
    If IsEmpty(cell.Value) Then
        GetCellValue = "(empty)"
    Else
        GetCellValue = CStr(cell.Value)
    End If
End Function

Function GetFormulaDescription(fx As String) As String
    If InStr(1, fx, "SUM") > 0 Then
        GetFormulaDescription = "Summation calculation"
    ElseIf InStr(1, fx, "IF") > 0 Then
        GetFormulaDescription = "Conditional logic"
    Else
        GetFormulaDescription = "General formula"
    End If
End Function
