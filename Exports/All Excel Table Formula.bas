' ==========================================================================================
' 📌 Macro: ListTableFormulas_Clean
' 📁 Module Purpose:
'     Scans all Excel tables in the current workbook and creates a comprehensive Markdown
'     documentation file in your Downloads folder. Maps structure, formulas, and relationships.
'
' ------------------------------------------------------------------------------------------
' ✅ Output Summary:
'     • Table-by-table metadata including worksheet, range, and size
'     • Column locations with cell references
'     • Column-level formula documentation and categorization
'     • Cross-references between tables (detected via formula text)
'     • Markdown-friendly output for AI-assisted analysis or documentation
'
' ------------------------------------------------------------------------------------------
' 🔍 Code Behavior Overview:
'     - Loops through all ListObjects (tables) in all worksheets
'     - Extracts metadata: name, location, size
'     - For each column:
'         - Records range and header cell address
'         - Checks if a formula exists and extracts the formula
'         - Categorizes formula type (e.g. SUMIFS, VLOOKUP, IFERROR, etc.)
'         - Detects cross-table references
'     - Exports results in Markdown format to: Downloads\Table_Formulas_AI.txt
'
' ------------------------------------------------------------------------------------------
' 🛠️ Usage Notes:
'     - Does not modify any data — read-only operation
'     - Designed to help audit formulas, document data models, or assist AI code generation
'     - Handles empty tables, merged cells, or complex ranges gracefully
'
' ------------------------------------------------------------------------------------------
' 🧠 Use Cases:
'     - Generate documentation for technical handoff or training
'     - Feed structure into AI models for smart formula generation
'     - Audit complex Excel workbooks and find table relationships
'     - Reverse-engineer workbook logic with clean output
'
' ==========================================================================================

Sub ListTableFormulas_Clean()
    Dim ws As Worksheet, tbl As ListObject, col As ListColumn
    Dim strFilePath As String, fileNum As Integer
    Dim tableCount As Long
    Dim columnIndex As Long, columnName As String
    Dim formulaText As String, formulaCategory As String, hasFormula As String
    Dim columnRange As String, headerCell As String

    strFilePath = Environ("USERPROFILE") & "\Downloads\Table_Formulas_AI.txt"
    fileNum = FreeFile
    Open strFilePath For Output As #fileNum

    Print #fileNum, "# Excel Table Formulas"
    Print #fileNum, "Generated on: " & Now()
    Print #fileNum, ""

    tableCount = 0

    For Each ws In ThisWorkbook.Worksheets
        If ws.ListObjects.Count = 0 Then GoTo NextSheet

        For Each tbl In ws.ListObjects
            tableCount = tableCount + 1

            Print #fileNum, "# TABLE: " & tbl.Name
            Print #fileNum, "Worksheet: " & ws.Name
            Print #fileNum, "Table Range: " & tbl.Range.Address
            Print #fileNum, "Header Row: " & tbl.HeaderRowRange.Address

            If Not tbl.DataBodyRange Is Nothing Then
                Print #fileNum, "Data Range: " & tbl.DataBodyRange.Address
                Print #fileNum, "Total Rows: " & tbl.DataBodyRange.Rows.Count
            Else
                Print #fileNum, "Data Range: N/A"
                Print #fileNum, "Total Rows: 0"
            End If

            Print #fileNum, "Total Columns: " & tbl.ListColumns.Count
            Print #fileNum, ""

            ' --- Column Locations ---
            Print #fileNum, "## COLUMN LOCATIONS"
            Print #fileNum, "| Column Index | Column Name | Column Range | Header Cell |"
            Print #fileNum, "|--------------|-------------|----------------|--------------|"

            For Each col In tbl.ListColumns
                columnIndex = col.Index
                columnName = col.Name
                columnRange = IIf(Not col.Range Is Nothing, col.Range.Address, "N/A")
                headerCell = IIf(Not col.Range Is Nothing, col.Range.Cells(1, 1).Address, "N/A")
                Print #fileNum, "| " & columnIndex & " | " & SafeText(columnName) & " | " & columnRange & " | " & headerCell & " |"
            Next col

            Print #fileNum, ""
            ' --- Column Formulas ---
            Print #fileNum, "## COLUMN FORMULAS"
            Print #fileNum, "| Column Index | Column Name | Has Formula | Formula | Formula Category |"
            Print #fileNum, "|--------------|-------------|-------------|---------|------------------|"

            For Each col In tbl.ListColumns
                columnIndex = col.Index
                columnName = col.Name
                hasFormula = "No"
                formulaText = "No formula used"
                formulaCategory = "N/A"

                If Not col.DataBodyRange Is Nothing And col.DataBodyRange.Rows.Count > 0 Then
                    If col.DataBodyRange.Cells(1).HasFormula Then
                        hasFormula = "Yes"
                        formulaText = col.DataBodyRange.Cells(1).Formula

                        Select Case True
                            Case InStr(formulaText, "SUMIFS") > 0: formulaCategory = "Aggregation (SUMIFS)"
                            Case InStr(formulaText, "SUM") > 0: formulaCategory = "Aggregation (SUM)"
                            Case InStr(formulaText, "AVERAGE") > 0: formulaCategory = "Aggregation (AVERAGE)"
                            Case InStr(formulaText, "COUNT") > 0: formulaCategory = "Aggregation (COUNT)"
                            Case InStr(formulaText, "INDEX") > 0 And InStr(formulaText, "MATCH") > 0: formulaCategory = "Lookup (INDEX/MATCH)"
                            Case InStr(formulaText, "VLOOKUP") > 0: formulaCategory = "Lookup (VLOOKUP)"
                            Case InStr(formulaText, "IF(") > 0: formulaCategory = "Conditional (IF)"
                            Case InStr(formulaText, "IFERROR") > 0: formulaCategory = "Error Handling (IFERROR)"
                            Case Else: formulaCategory = "Other"
                        End Select

                        ' Detect cross-table references
                        Dim checkWs As Worksheet, otherTbl As ListObject
                        For Each checkWs In ThisWorkbook.Worksheets
                            For Each otherTbl In checkWs.ListObjects
                                If otherTbl.Name <> tbl.Name Then
                                    If InStr(formulaText, otherTbl.Name & "[") > 0 Then
                                        formulaCategory = formulaCategory & " (References " & otherTbl.Name & ")"
                                    End If
                                End If
                            Next otherTbl
                        Next checkWs
                    End If
                End If

                Print #fileNum, "| " & columnIndex & " | " & SafeText(columnName) & " | " & hasFormula & " | " & SafeText(formulaText) & " | " & formulaCategory & " |"
            Next col

            Print #fileNum, vbNewLine & "---" & vbNewLine
        Next tbl
NextSheet:
    Next ws

    Print #fileNum, "# SUMMARY"
    Print #fileNum, "- Total Tables: " & tableCount
    Close #fileNum

    If tableCount > 0 Then
        MsgBox "✅ Table formulas report saved to:" & vbNewLine & strFilePath & vbNewLine & _
               "Processed " & tableCount & " table(s).", vbInformation
    Else
        MsgBox "⚠️ No tables found in this workbook.", vbExclamation
    End If
End Sub

Function SafeText(ByVal inputText As Variant) As String
    If IsError(inputText) Or IsEmpty(inputText) Then
        SafeText = "(empty)"
    Else
        SafeText = Replace(Replace(Replace(Replace(CStr(inputText), "|", "\|"), vbCrLf, " "), vbCr, " "), vbLf, " ")
        If Len(SafeText) > 250 Then SafeText = Left(SafeText, 247) & "..."
    End If
End Function