' ==========================================================================================
' 📌 Macro: ExportPivotToMarkdown
' 📁 Module Purpose:
'     Exports the first PivotTable on the active worksheet to a GitHub-friendly Markdown (.md)
'     file. Output preserves header structure and table alignment with pipe (`|`) delimiters.
'
' ------------------------------------------------------------------------------------------
' ✅ Sample Output:
'     | Department | Total ASF | Avg RDD |
'     |------------|-----------|---------|
'     | Surgery    | 98,000    | $312    |
'     | Pediatrics | 82,000    | $275    |
'
' ------------------------------------------------------------------------------------------
' 🔍 Code Behavior Overview:
'     - Detects the first PivotTable on the active worksheet
'     - Extracts the visible table range (including subtotals/grand totals)
'     - Translates it into Markdown format with headers and separator rows
'     - Saves to a user-specified location using `.md` file extension
'
' ------------------------------------------------------------------------------------------
' 🛠️ Notes:
'     - Requires at least one PivotTable on the active worksheet
'     - Uses `TableRange2` to capture the full pivot including filter labels
'     - Escapes special characters like `|` in values to avoid formatting issues
'     - Works best when row labels and column headers are cleanly formatted
'     - Output is compatible with GitHub, Obsidian, Notion, and Markdown preview tools
'
' ==========================================================================================
Sub ExportPivotToMarkdown()
    ' Export the first pivot table on the active sheet to a GitHub-compatible markdown file

    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim rng As Range
    Dim filePath As Variant
    Dim fileNum As Integer
    Dim rowIndex As Long, colIndex As Long
    Dim headerRow As Boolean
    Dim separatorLine As String
    Dim safeVal As String

    Set ws = ActiveSheet

    If ws.PivotTables.Count = 0 Then
        MsgBox "No pivot table found on this sheet!", vbCritical
        Exit Sub
    End If

    Set pt = ws.PivotTables(1)
    Set rng = pt.TableRange2

    If rng Is Nothing Then
        MsgBox "Could not determine the pivot table range.", vbCritical
        Exit Sub
    End If

    filePath = Application.GetSaveAsFilename("PivotTableExport.md", "Markdown Files (*.md), *.md")
    If filePath = "False" Then Exit Sub

    fileNum = FreeFile
    Open filePath For Output As #fileNum

    headerRow = True

    For rowIndex = 1 To rng.Rows.Count
        Print #fileNum, "|";
        For colIndex = 1 To rng.Columns.Count
            safeVal = SafeText(rng.Cells(rowIndex, colIndex).Text)
            Print #fileNum, " " & safeVal & " |";
        Next colIndex
        Print #fileNum, ""

        If headerRow Then
            separatorLine = "|"
            For colIndex = 1 To rng.Columns.Count
                separatorLine = separatorLine & " --- |"
            Next colIndex
            Print #fileNum, separatorLine
            headerRow = False
        End If
    Next rowIndex

    Close #fileNum

    MsgBox "✅ Pivot table exported to Markdown:" & vbNewLine & filePath, vbInformation
End Sub

Function SafeText(val As Variant) As String
    Dim txt As String
    txt = CStr(val)
    txt = Replace(txt, "|", "\|")
    txt = Replace(txt, vbCrLf, " ")
    txt = Replace(txt, vbLf, " ")
    txt = Replace(txt, vbCr, " ")
    SafeText = Trim(txt)
End Function
