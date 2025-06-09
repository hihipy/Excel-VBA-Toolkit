' ==========================================================================================
' 📌 Macro Suite: WhitespaceTools
' 📁 Module Purpose:
'     An optimized collection of macros to detect, highlight, and FIX various whitespace
'     issues in Excel. It reads data into memory arrays for maximum performance on large
'     datasets, finding leading, trailing, and multiple consecutive spaces.
'
' ------------------------------------------------------------------------------------------
' ✅ Key Features & Enhancements:
'     - DetectAllWhitespaceIssues_Fast: A new, highly efficient macro that processes
'       data in memory for a dramatic speed increase on large ranges.
'     - RemoveAllWhitespaceIssues: A new utility to automatically FIX all detected
'       whitespace issues in a selected range (trims ends and normalizes internal spaces).
'     - Original macros are retained for smaller, quick-and-dirty scans.
'
' ------------------------------------------------------------------------------------------
' 🔍 Code Behavior Overview:
'     - The _Fast macro reads the entire selection into a variant array.
'     - All logic (checking for spaces) is performed on the in-memory array.
'     - Addresses of problematic cells are collected into a string.
'     - Formatting is applied in a single operation at the end using the collected addresses.
'     - The "Fix" macro directly modifies cell values to clean them.
'
' ------------------------------------------------------------------------------------------
' 🛠️ Notes:
'     - The _Fast version is recommended for any selection larger than a few thousand cells.
'     - A confirmation prompt is included before the "Fix" macro modifies any data.
'     - All macros are robust, with clear user feedback and error handling.
'
' ==========================================================================================

Option Explicit

'==================================================================================================
'  HIGH-PERFORMANCE DETECTION & FIXING MACROS
'==================================================================================================

Sub DetectAllWhitespaceIssues_Fast()
    '==========================================================================
    ' PURPOSE: High-speed, comprehensive whitespace detection using memory arrays.
    ' USAGE: Select range/column, then run macro. Ideal for large datasets.
    '==========================================================================
    
    On Error GoTo ErrorHandler
    
    Dim targetRange As Range
    Dim dataArray As Variant
    Dim cell As Range
    Dim r As Long, c As Long
    Dim originalText As String
    Dim problemAddresses As String
    Dim leadingCount As Long, trailingCount As Long, multipleCount As Long
    Dim totalIssues As Long
    Dim startTime As Double
    
    startTime = Timer
    
    ' --- Validate and Set Target Range ---
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a range of cells first.", vbExclamation, "No Selection"
        Exit Sub
    End If
    
    Set targetRange = Intersect(Selection, ActiveSheet.UsedRange)
    
    If targetRange Is Nothing Then
        MsgBox "The selection is empty or outside the used range.", vbInformation, "Empty Selection"
        Exit Sub
    End If
    
    ' --- Performance Optimization ---
    Application.ScreenUpdating = False
    
    ' --- Clear existing formatting ---
    targetRange.Interior.Color = xlNone
    targetRange.Font.ColorIndex = xlAutomatic
    
    ' --- Read range into memory array ---
    If targetRange.count > 1 Then
        dataArray = targetRange.Value2
    Else
        ReDim dataArray(1 To 1, 1 To 1)
        dataArray(1, 1) = targetRange.Value2
    End If
    
    problemAddresses = ""
    
    ' --- Process the array in memory ---
    For r = 1 To UBound(dataArray, 1)
        For c = 1 To UBound(dataArray, 2)
            If VarType(dataArray(r, c)) = vbString Then
                originalText = dataArray(r, c)
                
                If Len(originalText) > 0 Then
                    Dim hasIssue As Boolean: hasIssue = False
                    
                    ' Check for leading spaces
                    If Left(originalText, 1) = " " Then
                        leadingCount = leadingCount + 1
                        hasIssue = True
                    End If
                    
                    ' Check for trailing spaces
                    If Right(originalText, 1) = " " Then
                        trailingCount = trailingCount + 1
                        hasIssue = True
                    End If
                    
                    ' Check for multiple consecutive spaces
                    If InStr(originalText, "  ") > 0 Then
                        multipleCount = multipleCount + 1
                        hasIssue = True
                    End If
                    
                    ' If an issue is found, collect the cell's address
                    If hasIssue Then
                        totalIssues = totalIssues + 1
                        If problemAddresses <> "" Then problemAddresses = problemAddresses & ","
                        problemAddresses = problemAddresses & targetRange.Cells(r, c).Address(False, False)
                    End If
                End If
            End If
        Next c
    Next r
    
    ' --- Apply formatting in a single operation ---
    If totalIssues > 0 Then
        Range(problemAddresses).Interior.Color = RGB(255, 200, 200) ' Light red background
    End If
    
    ' --- Restore Excel settings ---
    Application.ScreenUpdating = True
    
    ' --- Display comprehensive results ---
    Dim message As String
    If totalIssues = 0 Then
        message = "NO WHITESPACE ISSUES FOUND!" & vbNewLine & vbNewLine & _
                  "All " & targetRange.Cells.count & " cells are clean." & vbNewLine & vbNewLine & _
                  "Time: " & Format(Timer - startTime, "0.00") & " seconds"
    Else
        message = "WHITESPACE ISSUES DETECTED!" & vbNewLine & String(35, "=") & vbNewLine & vbNewLine & _
                  "ISSUE BREAKDOWN:" & vbNewLine & _
                  "   • Leading spaces found: " & leadingCount & " instances" & vbNewLine & _
                  "   • Trailing spaces found: " & trailingCount & " instances" & vbNewLine & _
                  "   • Multiple spaces found: " & multipleCount & " instances" & vbNewLine & _
                  "   • Total problem cells: " & totalIssues & vbNewLine & vbNewLine & _
                  "Problem cells highlighted in light red." & vbNewLine & _
                  "Time: " & Format(Timer - startTime, "0.00") & " seconds" & vbNewLine & vbNewLine & _
                  "Use the 'RemoveAllWhitespaceIssues' macro to fix these."
    End If
    
    MsgBox message, IIf(totalIssues = 0, vbInformation, vbExclamation), "Whitespace Analysis"
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "An unexpected error occurred: " & Err.Description, vbCritical, "Error Analyzing Whitespace"
End Sub

Sub RemoveAllWhitespaceIssues()
    '==========================================================================
    ' PURPOSE: Automatically cleans all whitespace issues from the selected cells.
    ' USAGE: Select range, then run macro. It will trim and normalize spaces.
    '==========================================================================
    On Error GoTo ErrorHandler
    
    Dim targetRange As Range
    Dim dataArray As Variant
    Dim r As Long, c As Long
    Dim cleanText As String
    Dim cellsModified As Long
    Dim startTime As Double
    
    startTime = Timer
    
    ' --- Validate and Set Target Range ---
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a range of cells first.", vbExclamation, "No Selection"
        Exit Sub
    End If
    
    Set targetRange = Intersect(Selection, ActiveSheet.UsedRange)
    
    If targetRange Is Nothing Then
        MsgBox "The selection is empty or outside the used range.", vbInformation, "Empty Selection"
        Exit Sub
    End If
    
    ' --- Confirm data modification with user ---
    If MsgBox("This will permanently modify the data in " & targetRange.count & " selected cells." & vbNewLine & vbNewLine & _
        "It will trim leading/trailing spaces and replace multiple internal spaces with a single space." & vbNewLine & vbNewLine & _
        "Do you want to proceed?", vbQuestion + vbYesNo, "Confirm Data Cleaning") = vbNo Then
        Exit Sub
    End If
    
    ' --- Performance Optimization ---
    Application.ScreenUpdating = False
    
    ' --- Read range into memory array ---
    If targetRange.count > 1 Then
        dataArray = targetRange.Value2
    Else
        ReDim dataArray(1 To 1, 1 To 1)
        dataArray(1, 1) = targetRange.Value2
    End If
    
    cellsModified = 0
    
    ' --- Process the array in memory ---
    For r = 1 To UBound(dataArray, 1)
        For c = 1 To UBound(dataArray, 2)
            If VarType(dataArray(r, c)) = vbString Then
                ' Use WorksheetFunction.Trim to handle both trimming and multiple spaces
                cleanText = Application.WorksheetFunction.Trim(dataArray(r, c))
                
                ' Only mark as modified if the text actually changed
                If dataArray(r, c) <> cleanText Then
                    dataArray(r, c) = cleanText
                    cellsModified = cellsModified + 1
                End If
            End If
        Next c
    Next r
    
    ' --- Write the cleaned array back to the range in one operation ---
    targetRange.Value2 = dataArray
    
    ' --- Clear any lingering highlighting ---
    targetRange.Interior.Color = xlNone
    targetRange.Font.ColorIndex = xlAutomatic
    
    Application.ScreenUpdating = True
    
    ' --- Display summary ---
    MsgBox cellsModified & " cells were cleaned successfully." & vbNewLine & vbNewLine & _
           "Total time: " & Format(Timer - startTime, "0.00") & " seconds.", vbInformation, "Cleaning Complete"

    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "An unexpected error occurred: " & Err.Description, vbCritical, "Error Cleaning Whitespace"
End Sub

'==================================================================================================
'  ORIGINAL CELL-BY-CELL MACROS (for smaller tasks)
'==================================================================================================

Sub DetectTrailingSpaces()
    '==========================================================================
    ' PURPOSE: Detect and highlight cells with trailing invisible spaces
    ' USAGE: Select range/column, then run macro
    '==========================================================================
    
    On Error GoTo ErrorHandler_Original
    
    Dim ws As Worksheet
    Dim targetRange As Range
    Dim cell As Range
    Dim trailingSpaceCells As String
    Dim count As Long
    Dim startTime As Double
    
    startTime = Timer
    
    ' Validate selection
    If Selection Is Nothing Then
        MsgBox "Please select a range or column first.", vbExclamation, "No Selection"
        Exit Sub
    End If
    
    ' Handle entire column selections efficiently
    Set ws = ActiveSheet
    If Selection.Rows.count = Rows.count Then
        Set targetRange = Intersect(Selection, ws.UsedRange)
        If targetRange Is Nothing Then
            MsgBox "No data found in selected column.", vbInformation, "Empty Selection"
            Exit Sub
        End If
    Else
        Set targetRange = Selection
    End If
    
    ' Clear any existing highlighting first
    targetRange.Interior.Color = xlNone
    
    ' Initialize results
    count = 0
    trailingSpaceCells = ""
    
    ' Check each cell for trailing spaces
    For Each cell In targetRange
        If Not IsEmpty(cell.Value) And VarType(cell.Value) = vbString Then
            Dim originalText As String
            originalText = CStr(cell.Value)
            
            ' Check if text ends with space(s) but isn't just spaces
            If Len(originalText) > 0 And Right(originalText, 1) = " " And Trim(originalText) <> "" Then
                ' Highlight the cell
                cell.Interior.Color = RGB(255, 255, 0) ' Yellow background
                cell.Font.Color = RGB(255, 0, 0)       ' Red text
                
                count = count + 1
                
                ' Build list of problem cells (limit to first 10 for display)
                If count <= 10 Then
                    If trailingSpaceCells <> "" Then trailingSpaceCells = trailingSpaceCells & vbNewLine
                    trailingSpaceCells = trailingSpaceCells & cell.Address & ": """ & originalText & """"
                End If
            End If
        End If
    Next cell
    
    ' Display results
    Dim message As String
    If count = 0 Then
        message = "NO TRAILING SPACES FOUND!" & vbNewLine & vbNewLine & _
                  "All " & targetRange.Cells.count & " cells are clean."
        MsgBox message, vbInformation, "Clean Data"
    Else
        message = "TRAILING SPACES DETECTED!" & vbNewLine & String(30, "=") & vbNewLine & vbNewLine & _
                  "Found " & count & " cells with trailing spaces:" & vbNewLine & vbNewLine & _
                  trailingSpaceCells
        
        If count > 10 Then
            message = message & vbNewLine & "... and " & (count - 10) & " more cells"
        End If
        
        message = message & vbNewLine & vbNewLine & _
                  "HIGHLIGHTED: Problem cells are now yellow with red text" & vbNewLine & _
                  "TIME: " & Format(Timer - startTime, "0.00") & " seconds" & vbNewLine & vbNewLine & _
                  "Use 'RemoveAllWhitespaceIssues' macro to fix these issues."
        
        MsgBox message, vbExclamation, "Trailing Spaces Found"
    End If
    
    Exit Sub
    
ErrorHandler_Original:
    MsgBox "Error: " & Err.Description, vbCritical, "Error Detecting Trailing Spaces"
End Sub

Sub ClearAllHighlighting()
    '==========================================================================
    ' PURPOSE: Clear highlighting from trailing space detection
    ' USAGE: Select the range to clear, then run macro
    '==========================================================================
    
    On Error GoTo ErrorHandler_Clear
    
    If Selection Is Nothing Then
        MsgBox "Please select the range to clear highlighting from.", vbExclamation, "No Selection"
        Exit Sub
    End If
    
    ' --- Clear formatting ---
    Selection.Interior.Color = xlNone
    Selection.Font.ColorIndex = xlAutomatic
    
    MsgBox "Highlighting cleared from selected range.", vbInformation, "Formatting Cleared"
    
    Exit Sub
    
ErrorHandler_Clear:
    MsgBox "Error: " & Err.Description, vbCritical, "Error Clearing Highlighting"
End Sub
