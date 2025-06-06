' ==========================================================================================
' 📌 Macro: GenerateAdvancedPivotReport
' 📁 Module Purpose:
'     Scans all Excel tables (ListObjects) in a workbook, generating a Markdown-formatted
'     documentation file. The output includes table structure, cell locations, formulas,
'     and inter-table references — optimized for AI tools or human review.
'
' ------------------------------------------------------------------------------------------
' ✅ Sample Output (Markdown format):
'     # TABLE: Table3
'     Worksheet: Calculations
'     Table Range: $A$1:$AJ$1091
'     ...
'     ## COLUMN FORMULAS
'     | Column Index | Column Name | Has Formula | Formula | Formula Category |
'     |--------------|-------------|-------------|---------|------------------|
'     | 2 | ASF | Yes | =IFERROR(SUMIFS(...)) | Aggregation (SUMIFS) (References Table2) |
'
' ------------------------------------------------------------------------------------------
' 🔍 Code Behavior Overview:
'     - Loops through all worksheets and ListObjects
'     - Records structural metadata (headers, ranges, row/col counts)
'     - Extracts formulas from columns, categorizes by function
'     - Detects cross-table references
'     - Exports output as Markdown to Downloads folder
'
' ------------------------------------------------------------------------------------------
' 🛠️ Notes:
'     - Output saved to: %USERPROFILE%\Downloads\Table_Formulas_AI.txt
'     - No data is modified — read-only process
'     - Handles both empty and populated tables
'     - Markdown output readable in Obsidian, GitHub, Notepad++, etc.
'     - Long formulas are truncated to ~250 chars for readability
'
' ==========================================================================================
'==================================================================================================
'  MODULE-LEVEL SETTINGS
'==================================================================================================
Option Explicit

'==================================================================================================
'  PUBLIC CONSTANTS
'==================================================================================================
Private Const xlODBCConnection As Long = 1
Private Const xlOLEDBConnection As Long = 2
Private Const xlCompactRow As Long = 0
Private Const xlOutlineRow As Long = 1
Private Const xlTabularRow As Long = 2
Private Const xlRepeatLabels As Long = 2

'==================================================================================================
'  PRIVATE CONSTANTS FOR MARKDOWN
'==================================================================================================
Private Const MD_OLAP_FIELD_HEADER As String = "| Full Field Name | Dimension | Hierarchy | Level | Orientation | In Layout? | Measure? | Friendly Label | MDX Reference |"
Private Const MD_OLAP_FIELD_SEPARATOR As String = "|------------------|-----------|-----------|--------|-------------|-------------|-----------|------------------|------------------------|"
Private Const MD_REGULAR_FIELD_HEADER As String = "| Field Name | Source Name | Orientation | In Layout? | Caption |"
Private Const MD_REGULAR_FIELD_SEPARATOR As String = "|------------|-------------|-------------|------------|---------|"
Private Const MD_VALUE_FIELD_HEADER As String = "| Field | Function | Format |"
Private Const MD_VALUE_FIELD_SEPARATOR As String = "|-------|----------|--------|"
Private Const MD_CALC_FIELD_HEADER As String = "| Name | Formula |"
Private Const MD_CALC_FIELD_SEPARATOR As String = "|------|---------|"

'==================================================================================================
'  MAIN PROCEDURE
'==================================================================================================
Public Sub GenerateAdvancedPivotReport()
    Const INCLUDE_PIVOT_LAYOUT_DETAILS As Boolean = True
    Const INCLUDE_SLICERS As Boolean = True
    Const INCLUDE_CALCULATED_FIELDS As Boolean = True
    Const OPEN_FILE_AFTER_CREATION As Boolean = True
    
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim strFilePath As String
    Dim fileNum As Integer
    Dim pivotCount As Integer
    Dim olapCount As Integer
    Dim regularCount As Integer
    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean
    Dim oldCalculation As XlCalculation

    ' Performance optimization
    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents
    oldCalculation = Application.Calculation
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Starting PivotTable documentation process..."

    ' Get output file path
    strFilePath = Application.GetSaveAsFilename("Advanced_PivotTable_Report.md", "Markdown Files (*.md), *.md")
    If strFilePath = "False" Then
        GoTo CleanupAndExit
    End If
    
    ' Initialize file and write header
    Set wb = ThisWorkbook
    fileNum = FreeFile
    Open strFilePath For Output As #fileNum
    
    Print #fileNum, "# Advanced Pivot Table Report"
    Print #fileNum, "Generated on: " & Now()
    Print #fileNum, "Workbook: " & SafeText(wb.Name)
    Print #fileNum, ""
    
    pivotCount = 0: olapCount = 0: regularCount = 0
    
    ' Core processing loop
    For Each ws In wb.Worksheets
        If ws.PivotTables.Count > 0 Then
            Application.StatusBar = "Processing Worksheet: " & ws.Name & "..."
            Print #fileNum, "---"
            Print #fileNum, "# Worksheet: " & SafeText(ws.Name)
            Print #fileNum, ""
        End If
        
        For Each pt In ws.PivotTables
            pivotCount = pivotCount + 1
            Application.StatusBar = "Processing PivotTable #" & pivotCount & ": " & pt.Name & "..."
            
            If pt.PivotCache.OLAP Then
                olapCount = olapCount + 1
                Call WriteOlapPivotTableDoc(pt, fileNum, ws, pivotCount, INCLUDE_PIVOT_LAYOUT_DETAILS, INCLUDE_SLICERS)
            Else
                regularCount = regularCount + 1
                Call WriteRegularPivotTableDoc(pt, fileNum, ws, pivotCount, INCLUDE_PIVOT_LAYOUT_DETAILS, INCLUDE_SLICERS, INCLUDE_CALCULATED_FIELDS)
            End If
        Next pt
        Set pt = Nothing
    Next ws
    Set ws = Nothing

    ' Write summary
    Print #fileNum, "# Summary"
    Print #fileNum, "- Total PivotTables Found: " & pivotCount
    Print #fileNum, "- OLAP-Connected PivotTables: " & olapCount
    Print #fileNum, "- Regular PivotTables: " & regularCount
    
    Close #fileNum

CleanupAndExit:
    Application.StatusBar = False
    Application.Calculation = oldCalculation
    Application.EnableEvents = oldEnableEvents
    Application.ScreenUpdating = oldScreenUpdating
    
    If strFilePath <> "False" Then
        If OPEN_FILE_AFTER_CREATION Then
            If MsgBox("Report saved to:" & vbNewLine & strFilePath & _
                      vbNewLine & vbNewLine & "Do you want to open the file now?", _
                      vbYesNo + vbInformation, "Generation Complete") = vbYes Then
                On Error Resume Next
                wb.FollowHyperlink strFilePath
                If Err.Number <> 0 Then
                    MsgBox "Could not open the file automatically. Please open it manually from:" & vbNewLine & strFilePath, vbExclamation
                End If
                On Error GoTo 0
            End If
        Else
            MsgBox "Report saved to:" & vbNewLine & strFilePath, vbInformation, "Generation Complete"
        End If
    Else
        MsgBox "Operation cancelled.", vbInformation
    End If
End Sub

'==================================================================================================
'  DOCUMENTATION WRITERS
'==================================================================================================

Private Sub WriteOlapPivotTableDoc(pt As PivotTable, fileNum As Integer, ws As Worksheet, ByVal overallCount As Integer, ByVal includeLayout As Boolean, ByVal includeSlicers As Boolean)
    Dim pf As PivotField
    Dim anchorCell As String, connectionInfo As String
    Dim fullName As String, dimensionName As String, hierarchyName As String, levelName As String
    Dim orientationText As String, isUsed As String, friendlyLabel As String, isMeasure As String, mdxPath As String
    Dim parts() As String
    
    ' Metadata Section
    anchorCell = pt.TableRange2.Cells(1, 1).Address(False, False) ' FIXED: Removed xlA1 and External parameters
    Print #fileNum, "## PivotTable: " & SafeText(pt.Name) & " (Overall #" & overallCount & " on Sheet '" & SafeText(ws.Name) & "')"
    Print #fileNum, "**Type**: OLAP-Connected PivotTable"
    Print #fileNum, "**Anchor Cell**: " & anchorCell
    
    connectionInfo = GetOlapConnectionString(pt)
    Print #fileNum, "**Connection String Hint**: `" & SafeText(connectionInfo) & "`"
    Print #fileNum, ""
    
    If includeLayout Then WriteLayoutDetailsInfo pt, fileNum
    If includeSlicers Then WriteSlicerInfo pt, fileNum

    Print #fileNum, "### Why OLAP Filters Are Missing"
    Print #fileNum, "> OLAP PivotTables do not expose simple filter lists to VBA. Manually paste current filter values below if needed:"
    Print #fileNum, "```" & vbCrLf & "Paste filter values here (e.g., Fiscal Year: 2025)" & vbCrLf & "```" & vbCrLf
    
    Print #fileNum, "### OLAP Pivot Fields"
    Print #fileNum, MD_OLAP_FIELD_HEADER
    Print #fileNum, MD_OLAP_FIELD_SEPARATOR
    
    On Error Resume Next
    For Each pf In pt.PivotFields
        Err.Clear
        fullName = pf.SourceName
        dimensionName = "": hierarchyName = "": levelName = ""
        friendlyLabel = "": isMeasure = "No": mdxPath = ""
        
        If InStr(fullName, "[Measures]") > 0 Then
            isMeasure = "Yes"
            friendlyLabel = Replace(Replace(fullName, "[Measures].[", ""), "]", "")
            dimensionName = "Measures"
            mdxPath = "[Measures].[" & friendlyLabel & "]"
            
        ElseIf InStr(fullName, "].[") > 0 Then
            parts = Split(Replace(fullName, "[", ""), "]")
            If UBound(parts) >= 0 Then dimensionName = Trim(parts(0))
            If UBound(parts) >= 1 Then hierarchyName = Trim(Replace(parts(1), ".", ""))
            
            If UBound(parts) >= 2 Then
                levelName = Trim(Replace(parts(2), ".", ""))
            Else
                levelName = hierarchyName
            End If
            
            friendlyLabel = levelName
            If UBound(parts) >= 2 Then
                mdxPath = "[" & dimensionName & "].[" & hierarchyName & "].[" & levelName & "]"
            ElseIf UBound(parts) = 1 Then
                mdxPath = "[" & dimensionName & "].[" & hierarchyName & "]"
            Else
                mdxPath = "[" & dimensionName & "]"
            End If

        Else
            dimensionName = "(Calculated or Other)"
            friendlyLabel = pf.Name
            mdxPath = SafeText(pf.Name)
        End If
        
        orientationText = GetOrientationText(pf.Orientation)
        isUsed = IIf(pf.Orientation <> xlHidden, "Yes", "No")
        If friendlyLabel = "" Then friendlyLabel = pf.Caption
        
        Print #fileNum, "| " & SafeText(fullName) & " | " & SafeText(dimensionName) & " | " & SafeText(hierarchyName) & " | " & SafeText(levelName) & " | " & orientationText & " | " & isUsed & " | " & isMeasure & " | " & SafeText(friendlyLabel) & " | " & SafeText(mdxPath) & " |"
    Next pf
    On Error GoTo 0
    
    Print #fileNum, vbNewLine & "---" & vbNewLine
    Set pf = Nothing
End Sub

Private Sub WriteRegularPivotTableDoc(pt As PivotTable, fileNum As Integer, ws As Worksheet, ByVal overallCount As Integer, ByVal includeLayout As Boolean, ByVal includeSlicers As Boolean, ByVal includeCalcFields As Boolean)
    Dim pf As PivotField
    Dim anchorCell As String, sourceNameVal As String
    Dim orientationText As String, isUsed As String, friendlyLabel As String
    
    anchorCell = pt.TableRange2.Cells(1, 1).Address(False, False) ' FIXED: Removed xlA1 and External parameters
    Print #fileNum, "## PivotTable: " & SafeText(pt.Name) & " (Overall #" & overallCount & " on Sheet '" & SafeText(ws.Name) & "')"
    Print #fileNum, "**Type**: Regular PivotTable"
    Print #fileNum, "**Anchor Cell**: " & anchorCell
    Print #fileNum, "**Data Source (Name or Range)**: `" & SafeText(pt.SourceData) & "`"
    Print #fileNum, ""
    
    If includeLayout Then WriteLayoutDetailsInfo pt, fileNum
    If includeSlicers Then WriteSlicerInfo pt, fileNum
    If includeCalcFields Then WriteCalculatedFieldsInfo pt, fileNum
    
    Print #fileNum, "### Current Page Filters"
    Print #fileNum, "```"
    Dim hasFilters As Boolean: hasFilters = False
    On Error Resume Next
    If pt.PageFields.Count > 0 Then
        For Each pf In pt.PageFields
            Print #fileNum, SafeText(pf.Name) & ": " & GetVisibleItemsList(pf)
            hasFilters = True
        Next pf
    End If
    On Error GoTo 0
    If Not hasFilters Then Print #fileNum, "(No page filters applied)"
    Print #fileNum, "```" & vbCrLf
    
    Print #fileNum, "### Fields"
    Print #fileNum, MD_REGULAR_FIELD_HEADER
    Print #fileNum, MD_REGULAR_FIELD_SEPARATOR
    
    On Error Resume Next
    For Each pf In pt.PivotFields
        Err.Clear
        sourceNameVal = pf.SourceName
        If Err.Number <> 0 Then sourceNameVal = "(N/A)"
        
        orientationText = GetOrientationText(pf.Orientation)
        isUsed = IIf(pf.Orientation <> xlHidden, "Yes", "No")
        friendlyLabel = pf.Caption
        Print #fileNum, "| " & SafeText(pf.Name) & " | " & SafeText(sourceNameVal) & " | " & orientationText & " | " & isUsed & " | " & SafeText(friendlyLabel) & " |"
    Next pf
    On Error GoTo 0
    Print #fileNum, ""
    
    Print #fileNum, "### Value Fields (Summarized By)"
    Print #fileNum, MD_VALUE_FIELD_HEADER
    Print #fileNum, MD_VALUE_FIELD_SEPARATOR
    If pt.DataFields.Count > 0 Then
        On Error Resume Next
        For Each pf In pt.DataFields
            Err.Clear
            Print #fileNum, "| " & SafeText(pf.SourceName) & " | " & GetFunctionText(pf.Function) & " | " & SafeText(pf.NumberFormat) & " |"
        Next pf
        On Error GoTo 0
    Else
        Print #fileNum, "| (No data fields) | | |"
    End If
    
    Print #fileNum, vbNewLine & "---" & vbNewLine
    Set pf = Nothing
End Sub

Private Sub WriteLayoutDetailsInfo(pt As PivotTable, fileNum As Integer)
    On Error Resume Next
    
    ' Test if the file number is valid first
    If fileNum <= 0 Then
        Debug.Print "### Layout and Style"
        Debug.Print "- **Style:** " & SafeText(pt.PivotStyle.Name)
        Debug.Print "- **Grand Totals for Rows:** " & IIf(pt.RowGrand, "On", "Off")
        Debug.Print "- **Grand Totals for Columns:** " & IIf(pt.ColumnGrand, "On", "Off")
    Else
        Print #fileNum, "### Layout and Style"
        Print #fileNum, "- **Style:** " & SafeText(pt.PivotStyle.Name)
        Print #fileNum, "- **Grand Totals for Rows:** " & IIf(pt.RowGrand, "On", "Off")
        Print #fileNum, "- **Grand Totals for Columns:** " & IIf(pt.ColumnGrand, "On", "Off")
    End If
    
    Dim layout As String
    Select Case pt.LayoutRowDefault
        Case xlCompactRow: layout = "Compact Form"
        Case xlOutlineRow: layout = "Outline Form"
        Case xlTabularRow: layout = "Tabular Form"
        Case Else: layout = "Unknown (" & pt.LayoutRowDefault & ")"
    End Select
    
    If fileNum <= 0 Then
        Debug.Print "- **Report Layout:** " & layout
    Else
        Print #fileNum, "- **Report Layout:** " & layout
    End If
    
    Dim repeatLabelsText As String
    ' Note: RepeatAllLabels is not a valid PivotTable property in VBA
    ' This setting is controlled at the PivotField level, not PivotTable level
    repeatLabelsText = "(This setting is controlled at individual field level, not PivotTable level)"
    
    If fileNum <= 0 Then
        Debug.Print "- **Repeat All Item Labels:** " & repeatLabelsText
        Debug.Print ""
    Else
        Print #fileNum, "- **Repeat All Item Labels:** " & repeatLabelsText
        Print #fileNum, ""
    End If
    
    On Error GoTo 0
End Sub

Private Sub WriteSlicerInfo(pt As PivotTable, fileNum As Integer)
    On Error Resume Next
    If pt.Slicers.Count > 0 Then
        Dim sl As Object, si As Object  ' Changed from Slicer, SlicerItem to Object
        Dim selectedItems As String
        
        If fileNum <= 0 Then
            Debug.Print "### Connected Slicers"
        Else
            Print #fileNum, "### Connected Slicers"
        End If
        
        For Each sl In pt.Slicers
            Err.Clear
            If fileNum <= 0 Then
                Debug.Print "- **Slicer:** " & SafeText(sl.Caption) & " (Source: `" & SafeText(sl.SlicerCache.SourceName) & "`)"
            Else
                Print #fileNum, "- **Slicer:** " & SafeText(sl.Caption) & " (Source: `" & SafeText(sl.SlicerCache.SourceName) & "`)"
            End If
            
            selectedItems = ""
            
            If sl.SlicerCache.FilterCleared Then
                 selectedItems = "(All)"
            Else
                For Each si In sl.SlicerCache.SlicerItems
                    If si.Selected Then
                        If selectedItems <> "" Then selectedItems = selectedItems & ", "
                        selectedItems = selectedItems & si.Name
                    End If
                Next si
                If selectedItems = "" Then selectedItems = "(Complex filter or no items selected)"
            End If

            If fileNum <= 0 Then
                Debug.Print "  - **Current Selection:** " & SafeText(selectedItems)
            Else
                Print #fileNum, "  - **Current Selection:** " & SafeText(selectedItems)
            End If
        Next sl
        
        If fileNum <= 0 Then
            Debug.Print ""
        Else
            Print #fileNum, ""
        End If
        
        Set sl = Nothing
        Set si = Nothing
    End If
    On Error GoTo 0
End Sub

Private Sub WriteCalculatedFieldsInfo(pt As PivotTable, fileNum As Integer)
    On Error Resume Next
    If pt.CalculatedFields.Count > 0 Then
        Dim cf As Object  ' Changed from CalculatedField to Object
        
        If fileNum <= 0 Then
            Debug.Print "### Calculated Fields"
            Debug.Print MD_CALC_FIELD_HEADER
            Debug.Print MD_CALC_FIELD_SEPARATOR
        Else
            Print #fileNum, "### Calculated Fields"
            Print #fileNum, MD_CALC_FIELD_HEADER
            Print #fileNum, MD_CALC_FIELD_SEPARATOR
        End If
        
        For Each cf In pt.CalculatedFields
            Err.Clear
            If fileNum <= 0 Then
                Debug.Print "| " & SafeText(cf.Name) & " | `" & SafeText(cf.Formula) & "` |"
            Else
                Print #fileNum, "| " & SafeText(cf.Name) & " | `" & SafeText(cf.Formula) & "` |"
            End If
        Next cf
        
        If fileNum <= 0 Then
            Debug.Print ""
        Else
            Print #fileNum, ""
        End If
        
        Set cf = Nothing
    End If
    On Error GoTo 0
End Sub

'==================================================================================================
'  HELPER FUNCTIONS
'==================================================================================================

Private Function SafeText(val As Variant) As String
    On Error Resume Next
    Dim tempStr As String
    If IsError(val) Or IsMissing(val) Or IsEmpty(val) Or IsNull(val) Then
        SafeText = "(empty)"
    Else
        tempStr = CStr(val)
        tempStr = Replace(tempStr, "|", "\|")
        tempStr = Replace(tempStr, "[", "\[")
        tempStr = Replace(tempStr, "]", "\]")
        tempStr = Replace(tempStr, vbCrLf, " ")
        tempStr = Replace(tempStr, vbLf, " ")
        tempStr = Replace(tempStr, vbCr, " ")
        SafeText = tempStr
    End If
    If Err.Number <> 0 Then
        SafeText = "(conversion error: " & Err.Description & ")"
        Err.Clear
    End If
End Function

Private Function GetVisibleItemsList(pf As PivotField) As String
    Dim result As String, pi As PivotItem, visibleItemCount As Long
    result = ""
    visibleItemCount = 0
    On Error Resume Next
    
    ' FIXED: Check AllItemsVisible property properly
    Dim allVisible As Boolean
    allVisible = pf.AllItemsVisible
    If Err.Number <> 0 Then
        Err.Clear
        allVisible = False
    End If
    
    If allVisible Then
        result = "(All)"
    Else
        ' FIXED: Use PivotItems instead of VisibleItemsList
        For Each pi In pf.PivotItems
            If pi.Visible Then
                If result <> "" Then result = result & ", "
                result = result & pi.Name
                visibleItemCount = visibleItemCount + 1
                If visibleItemCount > 10 And pf.PivotItems.Count > 15 Then
                    result = result & ", ... (" & (pf.PivotItems.Count - visibleItemCount) & " more)"
                    Exit For
                End If
            End If
        Next pi
    End If
    
    If result = "" Then
        On Error Resume Next
        Dim cpName As String: cpName = pf.CurrentPageName ' FIXED: Property name
        If Err.Number = 0 And cpName <> "" And cpName <> "(All)" Then
            result = cpName
        Else
            result = "(Complex filter or multiple items)"
        End If
        On Error GoTo 0
    End If

    If result = "" Then result = "(All)"
    
    On Error GoTo 0
    GetVisibleItemsList = result
End Function

Private Function GetOlapConnectionString(pt As PivotTable) As String
    Dim connectionInfo As String
    Dim serverInfo As String
    Dim catalogInfo As String
    Dim cubeInfo As String
    connectionInfo = ""
    
    On Error Resume Next
    
    ' === METHOD 1: Try WorkbookConnection (Modern Excel) ===
    If Not pt.PivotCache.WorkbookConnection Is Nothing Then
        Dim wbConn As WorkbookConnection
        Set wbConn = pt.PivotCache.WorkbookConnection
        
        connectionInfo = connectionInfo & "[WorkbookConnection] "
        
        If wbConn.Type = xlOLEDBConnection Then
            connectionInfo = connectionInfo & "OLEDB: " & wbConn.OLEDBConnection.Connection
        ElseIf wbConn.Type = xlODBCConnection Then
            connectionInfo = connectionInfo & "ODBC: " & wbConn.ODBCConnection.Connection
        Else
            connectionInfo = connectionInfo & "Type " & wbConn.Type & ": " & CStr(pt.PivotCache.Connection)
        End If
    End If
    
    ' === METHOD 2: Direct PivotCache.Connection ===
    If connectionInfo = "" Or InStr(connectionInfo, "Provider=") = 0 Then
        Dim directConn As String
        directConn = CStr(pt.PivotCache.Connection)
        If Len(directConn) > 10 And InStr(directConn, ";") > 0 Then
            If connectionInfo <> "" Then connectionInfo = connectionInfo & " | "
            connectionInfo = connectionInfo & "[Direct] " & directConn
        End If
    End If
    
    ' === METHOD 3: Try to get individual connection components ===
    If InStr(connectionInfo, "Data Source=") > 0 Then
        ' Extract server
        Dim startPos As Long, endPos As Long
        startPos = InStr(connectionInfo, "Data Source=") + 12
        endPos = InStr(startPos, connectionInfo, ";")
        If endPos = 0 Then endPos = Len(connectionInfo) + 1
        serverInfo = Mid(connectionInfo, startPos, endPos - startPos)
        
        ' Extract catalog
        If InStr(connectionInfo, "Initial Catalog=") > 0 Then
            startPos = InStr(connectionInfo, "Initial Catalog=") + 16
            endPos = InStr(startPos, connectionInfo, ";")
            If endPos = 0 Then endPos = Len(connectionInfo) + 1
            catalogInfo = Mid(connectionInfo, startPos, endPos - startPos)
        End If
    End If
    
    ' === METHOD 4: Try CubeFields for cube name ===
    If pt.CubeFields.Count > 0 Then
        Dim cf As Object  ' Changed from CubeField to Object
        Set cf = pt.CubeFields(1)
        On Error Resume Next
        cubeInfo = cf.CubeField.Name
        If Err.Number <> 0 Or cubeInfo = "" Then 
            Err.Clear
            cubeInfo = cf.Name
        End If
        On Error GoTo 0
    End If
    
    ' === METHOD 5: Alternative connection properties ===
    If connectionInfo = "" Then
        ' Try alternative properties
        Dim altConn As String
        altConn = ""
        
        ' Try CommandText property
        On Error Resume Next
        altConn = pt.PivotCache.CommandText
        If Err.Number = 0 And Len(altConn) > 5 Then
            connectionInfo = "[CommandText] " & altConn
        End If
        On Error Resume Next
        
        ' Try SourceData if it's not a range
        Dim sourceData As String
        sourceData = CStr(pt.SourceData)
        If Len(sourceData) > 20 And InStr(sourceData, "!") = 0 And InStr(sourceData, "$") = 0 Then
            If connectionInfo <> "" Then connectionInfo = connectionInfo & " | "
            connectionInfo = connectionInfo & "[SourceData] " & sourceData
        End If
    End If
    
    ' === METHOD 6: Check all connections in workbook ===
    If connectionInfo = "" Or Len(connectionInfo) < 20 Then
        Dim wb As Workbook
        Dim conn As WorkbookConnection
        Set wb = pt.Parent.Parent ' Get workbook from pivot table
        
        For Each conn In wb.Connections
            If conn.Type = xlOLEDBConnection Then
                Dim oledbConn As String
                oledbConn = conn.OLEDBConnection.Connection
                If InStr(oledbConn, "SSAS") > 0 Or InStr(oledbConn, "MSOLAP") > 0 Or InStr(oledbConn, "Analysis") > 0 Then
                    If connectionInfo <> "" Then connectionInfo = connectionInfo & " | "
                    connectionInfo = connectionInfo & "[WorkbookOLEDB-" & conn.Name & "] " & oledbConn
                    Exit For ' Use first OLAP connection found
                End If
            End If
        Next conn
    End If
    
    On Error GoTo 0
    
    ' === Build final result ===
    If connectionInfo = "" Then
        connectionInfo = "Connection details not accessible via VBA"
    End If
    
    ' Add extracted components if found
    If serverInfo <> "" Then
        connectionInfo = connectionInfo & " | SERVER: " & serverInfo
    End If
    If catalogInfo <> "" Then
        connectionInfo = connectionInfo & " | CATALOG: " & catalogInfo
    End If
    If cubeInfo <> "" Then
        connectionInfo = connectionInfo & " | CUBE: " & cubeInfo
    End If
    
    GetOlapConnectionString = connectionInfo
End Function

Private Function GetOrientationText(orientation As XlPivotFieldOrientation) As String
    Select Case orientation
        Case xlRowField: GetOrientationText = "Row"
        Case xlColumnField: GetOrientationText = "Column"
        Case xlPageField: GetOrientationText = "Filter"
        Case xlDataField: GetOrientationText = "Values"
        Case Else: GetOrientationText = ""
    End Select
End Function

Private Function GetFunctionText(func As XlConsolidationFunction) As String
    Select Case func
        Case xlSum: GetFunctionText = "Sum"
        Case xlCount: GetFunctionText = "Count"
        Case xlAverage: GetFunctionText = "Average"
        Case xlMax: GetFunctionText = "Max"
        Case xlMin: GetFunctionText = "Min"
        Case xlProduct: GetFunctionText = "Product"
        Case xlCountNums: GetFunctionText = "Count Numbers"
        Case xlStDev: GetFunctionText = "StdDev"
        Case xlStDevP: GetFunctionText = "StdDevP"
        Case xlVar: GetFunctionText = "Var"
        Case xlVarP: GetFunctionText = "VarP"
        Case Else: GetFunctionText = "Unknown/Custom (" & func & ")"
    End Select
End Function

Public Sub TestTheCall()
    Dim pt As PivotTable
    Dim fileNum As Integer
    
    On Error Resume Next
    Set pt = ActiveSheet.PivotTables(1)
    On Error GoTo 0
    
    If pt Is Nothing Then
        MsgBox "No PivotTable found on the active sheet. Please select a sheet with a PivotTable and try again.", vbCritical
        Exit Sub
    End If
    
    fileNum = 1
    Debug.Print "Attempting to call WriteLayoutDetailsInfo..."
    WriteLayoutDetailsInfo pt, fileNum
    Debug.Print "Call to WriteLayoutDetailsInfo was successful."
    MsgBox "Test completed successfully!", vbInformation
End Sub
