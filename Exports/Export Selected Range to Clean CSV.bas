' ==========================================================================================
' 📌 Macro: Export Selected Range to CSV (Fast Buffered)
' 📁 Module Purpose:
'     Exports a user-selected Excel range to a clean, analysis-ready `.csv` file.
'     The macro optimizes for speed, quoting consistency, and large datasets.
'
' ------------------------------------------------------------------------------------------
' ✅ Key Features:
'     - Supports 50,000+ rows with fast buffered output
'     - Auto-detects numeric vs. text columns
'     - Handles quote escaping, commas, line breaks, and scientific notation
'     - Lets the user choose delimiter (system default or comma)
'     - Output is compatible with R, Python (Pandas), Stata, and SPSS
'
' ------------------------------------------------------------------------------------------
' 🔍 Core Behaviors:
'     - Prompts user to select a range and output file
'     - Infers column types based on sample rows
'     - Escapes line breaks and double quotes
'     - Allows user to control text quoting behavior:
'         • Quote all
'         • Quote only when needed
'         • No quoting
'     - Processes all lines in-memory and writes in a single batch
'
' ------------------------------------------------------------------------------------------
' 📊 Example Output (CSV):
'     "Name","Age","Income"
'     "John Smith",35,"$75,000"
'     "Jane Doe",29,"$82,500"
'
' ------------------------------------------------------------------------------------------
' 🛠️ Technical Notes:
'     - Uses `Scripting.Dictionary` as an in-memory buffer for speed
'     - Removes line breaks and CR/LF characters inside cells
'     - Replaces empty cells with "0" (numeric) or blank (text)
'     - Uses `InputBox` for flexible range selection
'     - Safely handles overwrite warnings and open-file conflicts
'
' ------------------------------------------------------------------------------------------
' 📁 Use Cases:
'     - Clean export of pivot table output for machine learning
'     - Preprocessing Excel exports for Pandas DataFrames
'     - Rapid CSV generation during audit or QA sessions
'     - Extracting time series, cross-tab, or survey data from Excel
'
' ==========================================================================================
Sub ExportPivotToCSV_FastBuffered()
    ' Export selected range to CSV with fast buffered writing and smart quoting

    Const QUOTE_NONE As Integer = 3
    Dim selectedRange As Range, filePath As Variant
    Dim data As Variant, lineParts() As String
    Dim isNumericColumn() As Boolean
    Dim sb As Object ' Dictionary buffer
    Dim r As Long, c As Long, totalRows As Long
    Dim startTime As Double, fileNum As Integer
    Dim delimiterChar As String, quotingOption As VbMsgBoxResult

    On Error GoTo ErrorHandler

    startTime = Timer

    ' Select delimiter
    If Application.International(xlListSeparator) <> "," Then
        If MsgBox("Use system list separator (" & Application.International(xlListSeparator) & _
                  ") instead of comma?", vbYesNo + vbQuestion, "CSV Delimiter") = vbYes Then
            delimiterChar = Application.International(xlListSeparator)
        Else
            delimiterChar = ","
        End If
    Else
        delimiterChar = ","
    End If

    ' Ask about quoting style
    quotingOption = MsgBox("How should text be quoted?" & vbNewLine & _
        "• Yes = Quote all text" & vbNewLine & _
        "• No = Quote only when necessary" & vbNewLine & _
        "• Cancel = No quoting at all", _
        vbYesNoCancel + vbQuestion, "Text Quoting")
    If quotingOption = vbCancel Then quotingOption = QUOTE_NONE

    MsgBox "Highlight the range to export (including headers), then click OK.", vbInformation

    ' Select range
    On Error Resume Next
    Set selectedRange = Application.InputBox("Select range to export:", Type:=8)
    On Error GoTo ErrorHandlerNoCleanup
    If selectedRange Is Nothing Then
        MsgBox "No range selected. Exiting.", vbExclamation
        Exit Sub
    End If

    ' Choose output file
    filePath = Application.GetSaveAsFilename( _
        InitialFileName:=ActiveSheet.Name & "_Export.csv", _
        FileFilter:="CSV Files (*.csv), *.csv", _
        Title:="Save CSV As")
    If filePath = False Then Exit Sub

    ' Check for open file conflict
    If Dir(filePath) <> "" Then
        On Error Resume Next
        fileNum = FreeFile
        Open filePath For Input As #fileNum
        If Err.Number <> 0 Then
            MsgBox "File is open in another program. Please close it and try again.", vbExclamation
            Exit Sub
        End If
        Close #fileNum
        If MsgBox("File exists. Overwrite?", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    End If

    ' Optimize environment
    Application.StatusBar = "Preparing export..."
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    ' Load data
    data = selectedRange.Value
    totalRows = UBound(data, 1)
    ReDim isNumericColumn(1 To UBound(data, 2))

    ' Infer data types
    Dim inferRow As Long: inferRow = IIf(totalRows >= 2, 2, 1)
    For c = 1 To UBound(data, 2)
        Dim numericCount As Long: numericCount = 0
        Dim samplesChecked As Long: samplesChecked = 0
        For r = inferRow To Application.WorksheetFunction.Min(inferRow + 10, totalRows)
            If Not IsEmpty(data(r, c)) And Trim(CStr(data(r, c))) <> "" Then
                samplesChecked = samplesChecked + 1
                If IsNumeric(data(r, c)) Then numericCount = numericCount + 1
            End If
        Next r
        isNumericColumn(c) = (samplesChecked > 0 And numericCount / samplesChecked >= 0.8)
    Next c

    ' Use Dictionary as line buffer
    Set sb = CreateObject("Scripting.Dictionary")
    Dim progressInterval As Long: progressInterval = Application.WorksheetFunction.Max(1, totalRows \ 20)

    ' Process each row
    For r = 1 To totalRows
        If r Mod progressInterval = 0 Or r = totalRows Then
            Application.StatusBar = "Exporting row " & r & " of " & totalRows & "..."
            DoEvents
        End If

        ReDim lineParts(1 To UBound(data, 2))
        For c = 1 To UBound(data, 2)
            Dim val As String
            If IsEmpty(data(r, c)) Or Trim(CStr(data(r, c))) = "" Then
                val = IIf(isNumericColumn(c), "0", "")
            ElseIf IsNumeric(data(r, c)) Then
                Dim dblVal As Double: dblVal = CDbl(data(r, c))
                If Abs(dblVal) < 0.0000001 Then
                    val = "0"
                ElseIf Abs(dblVal) >= 100000000000000# Then
                    val = Format(dblVal, "0")
                Else
                    val = CStr(dblVal)
                End If
            Else
                val = CStr(data(r, c))
            End If

            val = Replace(val, vbCrLf, " ")
            val = Replace(val, vbCr, " ")
            val = Replace(val, vbLf, " ")
            val = Replace(val, """", """""")

            Select Case quotingOption
                Case vbYes
                    If Not isNumericColumn(c) Then val = """" & val & """"
                Case vbNo
                    If InStr(val, delimiterChar) > 0 Or InStr(val, """") > 0 Or _
                       InStr(val, vbCr) > 0 Or InStr(val, vbLf) > 0 Then
                        val = """" & val & """"
                    End If
                Case QUOTE_NONE
                    ' do nothing
            End Select
            lineParts(c) = val
        Next c
        sb.Add sb.Count + 1, Join(lineParts, delimiterChar)
    Next r

    ' Write all lines in one pass
    Application.StatusBar = "Writing to file..."
    fileNum = FreeFile
    Open filePath For Output As #fileNum
    For r = 1 To sb.Count
        Print #fileNum, sb(r)
    Next r
    Close #fileNum

CleanExit:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "✅ Export complete!" & vbNewLine & _
           "• File: " & filePath & vbNewLine & _
           "• Rows: " & Format(totalRows, "#,##0") & vbNewLine & _
           "• Time: " & Format(Timer - startTime, "0.00") & " seconds", vbInformation
    Exit Sub

ErrorHandler:
    On Error Resume Next
    Close #fileNum
    MsgBox "❌ Error " & Err.Number & ": " & Err.Description, vbCritical
    Resume CleanExit

ErrorHandlerNoCleanup:
    MsgBox "❌ Error selecting range: " & Err.Description, vbCritical
    Application.StatusBar = False
End Sub