' ==========================================================================================
' 📌 Macro: DeleteAllHiddenRows_BottomUp
' 📁 Module Purpose:
'     Deletes all hidden rows in the active worksheet, working from bottom to top.
'     Optimized for large datasets using Union batching, progress feedback, and time tracking.
'
' ------------------------------------------------------------------------------------------
' ✅ Sample Output:
'     • Status bar: "Scanning row 500 of 1000 (50%)"
'     • Final MsgBox: "Process complete! • Deleted: 45 hidden row(s) • Total rows scanned: 1000"
'
' ------------------------------------------------------------------------------------------
' 🔍 Code Behavior Overview:
'     - Uses UsedRange to detect last row across ALL columns (not just column A)
'     - Unhides hidden rows before collecting them (ensures deletion works under filters)
'     - Collects all hidden rows into a single Union range, then deletes in one operation
'     - Reverse loop prevents row index shifting issues
'     - Disables screen updating, events, and auto-calculation for performance
'     - Displays real-time scan progress in Excel's status bar every 500 rows
'
' ------------------------------------------------------------------------------------------
' 🛠️ Notes:
'     - Targets the active worksheet (Set ws = ActiveSheet)
'     - For specific sheets, replace with Worksheets("YourSheetName")
'     - Safe to run multiple times
'     - Union batching is significantly faster than row-by-row deletion on large datasets
'
' ==========================================================================================
Sub DeleteAllHiddenRows_BottomUp()
    Dim lastRow As Long, iCntr As Long
    Dim ws As Worksheet
    Dim startTime As Double, deletedRows As Long
    Dim rngToDelete As Range

    On Error GoTo ErrorHandler
    startTime = Timer
    deletedRows = 0

    Set ws = ActiveSheet

    ' Find true last row across ALL columns (not just A)
    lastRow = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1

    If lastRow < 1 Then
        MsgBox "No data found on the active sheet.", vbInformation
        Exit Sub
    End If

    ' Optimize performance
    Application.StatusBar = "Preparing to delete hidden rows..."
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ' Collect all hidden rows into a single range
    For iCntr = lastRow To 1 Step -1
        If iCntr Mod 500 = 0 Then
            Application.StatusBar = "Scanning row " & iCntr & " of " & lastRow & _
                                    " (" & Format((lastRow - iCntr) / lastRow, "0%") & ")"
            DoEvents
        End If

        If ws.Rows(iCntr).Hidden = True Then
            ' Unhide first so Excel can delete it
            ws.Rows(iCntr).Hidden = False

            If rngToDelete Is Nothing Then
                Set rngToDelete = ws.Rows(iCntr)
            Else
                Set rngToDelete = Union(rngToDelete, ws.Rows(iCntr))
            End If
            deletedRows = deletedRows + 1
        End If
    Next iCntr

    ' Delete all collected rows at once
    If Not rngToDelete Is Nothing Then
        rngToDelete.EntireRow.Delete
    End If

CleanExit:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

    MsgBox "Process complete!" & vbNewLine & _
           "Deleted: " & deletedRows & " hidden row(s)" & vbNewLine & _
           "Total rows scanned: " & lastRow & vbNewLine & _
           "Time elapsed: " & Format(Timer - startTime, "0.00") & " seconds", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
    Resume CleanExit
End Sub