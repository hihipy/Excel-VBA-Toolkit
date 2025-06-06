' ==========================================================================================
' 📌 Macro: Delete Hidden Rows (Bottom-Up)
' 📁 Module Purpose:
'     Deletes all hidden rows in the active worksheet, working from bottom to top.
'     Optimized for large datasets with visual progress feedback and execution time tracking.
'
' ------------------------------------------------------------------------------------------
' ✅ Sample Output:
'     • Status bar: "Processing row 990 of 1000 (1%)"
'     • Final MsgBox: "✅ Process complete! • Deleted: 45 hidden row(s) • Total rows checked: 1000"
'
' ------------------------------------------------------------------------------------------
' 🔍 Code Behavior Overview:
'     - Automatically detects last used row
'     - Skips visible rows, deletes hidden ones
'     - Uses reverse loop to avoid shifting row indexes
'     - Disables screen updating & events for performance
'     - Displays real-time status in Excel’s status bar
'
' ------------------------------------------------------------------------------------------
' 🛠️ Notes:
'     - Targets the active worksheet (Set ws = ActiveSheet)
'     - For specific sheets, replace with Worksheets("YourSheetName")
'     - Safe to run multiple times
'     - If you prefer SpecialCells approach:
'         ws.Range("A1:A" & lastRow).SpecialCells(xlCellTypeVisible).EntireRow.Delete
'
' ==========================================================================================

Sub sbVBS_To_Delete_Hidden_Rows()
    ' Deletes hidden rows in the active sheet
    Dim lastRow As Long, iCntr As Long
    Dim ws As Worksheet
    Dim startTime As Double, deletedRows As Long

    On Error GoTo ErrorHandler
    startTime = Timer
    deletedRows = 0
    Set ws = ActiveSheet

    ' Find last used row in column A
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Optimize performance
    Application.StatusBar = "Preparing to delete hidden rows..."
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ' Loop bottom to top
    For iCntr = lastRow To 1 Step -1
        ' Optional: Show progress every 100 rows
        If iCntr Mod 100 = 0 Then
            Application.StatusBar = "Processing row " & iCntr & " of " & lastRow & _
                                    " (" & Format((lastRow - iCntr) / lastRow, "0%") & ")"
            DoEvents
        End If

        ' If hidden, delete row
        If ws.Rows(iCntr).Hidden = True Then
            ws.Rows(iCntr).EntireRow.Delete
            deletedRows = deletedRows + 1
        End If
    Next iCntr

CleanExit:
    ' Restore settings
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

    ' Completion message
    MsgBox "✅ Process complete!" & vbNewLine & _
           "• Deleted: " & deletedRows & " hidden row(s)" & vbNewLine & _
           "• Total rows checked: " & lastRow & vbNewLine & _
           "• Time elapsed: " & Format(Timer - startTime, "0.00") & " seconds", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "❌ Error " & Err.Number & ": " & Err.Description, vbCritical
    Resume CleanExit
End Sub