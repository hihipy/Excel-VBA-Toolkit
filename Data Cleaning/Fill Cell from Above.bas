' ==========================================================================================
' 📌 Macro: Fill Empty Cells with Value from Above
' 📁 Module Purpose:
'     Fills blank cells in a selected range with the value from the cell directly above it.
'     This is commonly used for cleaning exports or grouped data that omits repeated labels.
'
' ------------------------------------------------------------------------------------------
' ✅ Sample Output:
'     • "✅ Filled 42 empty cell(s) with values from above."
'     • Or: "No empty cells were found to fill."
'
' ------------------------------------------------------------------------------------------
' 🔍 Code Behavior Overview:
'     - Works on user-selected range
'     - Skips merged cells and top row
'     - Detects empty cells (Value2 = "") in each column
'     - Fills them only if the cell above is non-empty
'     - Tracks how many cells were updated
'
' ------------------------------------------------------------------------------------------
' 🛠️ Usage Notes:
'     - Select a column or rectangular range before running
'     - Ideal for structured, top-down lists like department names, IDs, or categories
'     - Will not work on merged cells (alerts the user)
'
' ------------------------------------------------------------------------------------------
' 🧠 When to Use:
'     - Flattening pivot tables or grouped reports that only show the first row label
'     - Cleaning CSV exports that omit repeated values
'     - Preparing tabular data for filters, joins, or analysis
'
' ==========================================================================================

Sub FillCellFromAbove()
    ' Purpose: Fills empty cells in a selection with the value from the cell directly above

    Dim originalSelection As Range
    Dim col As Range, r As Long
    Dim affectedCells As Long

    On Error GoTo ErrorHandler

    Set originalSelection = Selection
    If originalSelection Is Nothing Then
        MsgBox "Please select a range first.", vbExclamation
        Exit Sub
    End If

    If originalSelection.MergeCells Then
        MsgBox "Cannot process merged cells. Please unmerge first.", vbExclamation
        Exit Sub
    End If

    If originalSelection.Row = 1 Then
        MsgBox "Top row selected — nothing above to pull from.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    affectedCells = 0

    For Each col In originalSelection.Columns
        For r = 2 To col.Cells.Count
            If col.Cells(r).Value2 = "" And col.Cells(r - 1).Value2 <> "" Then
                col.Cells(r).Value2 = col.Cells(r - 1).Value2
                affectedCells = affectedCells + 1
            End If
        Next r
    Next col

    Application.ScreenUpdating = True
    originalSelection.Select

    If affectedCells > 0 Then
        MsgBox "✅ Filled " & affectedCells & " empty cell(s) with values from above.", vbInformation
    Else
        MsgBox "No empty cells were found to fill.", vbInformation
    End If
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "❌ Error " & Err.Number & ": " & Err.Description, vbCritical
End Sub