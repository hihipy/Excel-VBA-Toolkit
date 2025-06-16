' ==========================================================================================
' 📌 Macro: GenerateUniversalAITableDoc
' 📁 Module Purpose:
'     Creates comprehensive, AI-optimized documentation of **all Excel Tables (ListObjects)**
'     across every worksheet in the workbook. Designed for universal compatibility across
'     all industries and data types - focuses on structural and technical properties only.
'
' ------------------------------------------------------------------------------------------
' ✅ Key Features:
'     - Scans every worksheet and every table (ListObject) in the workbook
'     - AI-optimized output with enhanced metadata including:
'         • Table name, range, data dimensions, and structural properties
'         • Column definitions with agnostic data type detection and sample values
'         • Complete formula transparency with exact syntax and dependencies
'         • Data quality flags (CLEAN, WARNING, ERROR) with specific percentages
'         • Performance optimization guidance based on dataset size and complexity
'     - Outputs clean, text-only Markdown file compatible with any AI tool
'     - No business-specific assumptions - works with ANY Excel table structure
'
' ------------------------------------------------------------------------------------------
' 🔍 Core Behaviors:
'     - Processes all sheets (visible, hidden, protected) with comprehensive error handling
'     - Agnostic data quality assessment including null percentages and empty field detection
'     - Universal data type recognition (Text, Number, Date, Currency, Formula, Empty)
'     - Formula detection with complete syntax display and dependency mapping
'     - Performance considerations tailored to dataset size (small/medium/large/very large)
'     - Text-only formatting for universal AI compatibility
'
' ------------------------------------------------------------------------------------------
' 🧠 Use Cases:
'     - Providing complete table context to AI for accurate formula generation
'     - Enabling AI to understand data quality constraints and validation rules
'     - Supporting complex multi-table operations with dependency awareness
'     - Generating enterprise-grade data dictionaries optimized for AI parsing
'     - Quality-checking Excel structures before automated analysis workflows
'     - Universal compatibility for any industry: financial, healthcare, manufacturing, retail, etc.
'
' ==========================================================================================

Sub GenerateUniversalAITableDoc()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim output As String
    Dim outputPath As String
    Dim totalTables As Integer
    Dim totalRows As Long
    Dim startTime As Double
    
    startTime = Timer
    Set wb = ActiveWorkbook
    totalTables = 0
    totalRows = 0
    
    ' Count all tables first
    For Each ws In wb.Worksheets
        totalTables = totalTables + ws.ListObjects.Count
    Next ws
    
    If totalTables = 0 Then
        MsgBox "No Excel Tables (ListObjects) found in this workbook.", vbInformation
        Exit Sub
    End If
    
    ' Build header
    output = "# AI-READY EXCEL TABLE DOCUMENTATION" & vbNewLine
    output = output & "Generated: " & Format(Now, "yyyy-mm-dd hh:mm:ss") & vbNewLine
    output = output & "Workbook: " & wb.Name & vbNewLine
    output = output & "Total Tables: " & totalTables & vbNewLine & vbNewLine
    
    ' Quick reference for AI
    output = output & "## QUICK REFERENCE FOR AI" & vbNewLine
    output = output & "- Use table references: TableName[ColumnName]" & vbNewLine
    output = output & "- XLOOKUP is preferred over VLOOKUP" & vbNewLine
    output = output & "- Check data quality flags before complex analysis" & vbNewLine
    output = output & "- Consider performance notes for large datasets" & vbNewLine & vbNewLine
    output = output & "---" & vbNewLine & vbNewLine
    
    ' Process each table
    For Each ws In wb.Worksheets
        For Each tbl In ws.ListObjects
            output = output & ProcessTable(tbl)
            totalRows = totalRows + GetTotalRows(tbl)
        Next tbl
    Next ws
    
    ' Add summary
    output = output & "# AI CODING SUMMARY" & vbNewLine
    output = output & "- **Tables processed**: " & totalTables & vbNewLine
    output = output & "- **Total data rows**: " & Format(totalRows, "#,##0") & vbNewLine
    output = output & "- **Processing time**: " & Format(Timer - startTime, "0.0") & " seconds" & vbNewLine
    output = output & "- **Recommended approach**: Use structured references and XLOOKUP functions" & vbNewLine
    output = output & "- **Performance**: " & GetOverallPerformanceAdvice(totalRows, totalTables) & vbNewLine
    
    ' Save output
    outputPath = ChooseOutputLocation()
    If outputPath <> "" Then
        SaveOutput output, outputPath
        MsgBox "AI-ready documentation generated successfully!" & vbNewLine & vbNewLine & _
               "File: " & outputPath, vbInformation
    End If
End Sub

Function ProcessTable(tbl As ListObject) As String
    Dim output As String
    Dim col As ListColumn
    Dim rowCount As Long
    Dim colCount As Integer
    Dim sizeCategory As String
    
    rowCount = GetTotalRows(tbl)
    colCount = tbl.ListColumns.Count
    sizeCategory = GetSizeCategory(rowCount)
    
    ' Table header
    output = "# TABLE: " & tbl.Name & vbNewLine & vbNewLine
    
    ' Basic info
    output = output & "## BASIC INFO" & vbNewLine
    output = output & "- **Worksheet**: " & tbl.Parent.Name & vbNewLine
    output = output & "- **Range**: " & tbl.Range.Address & vbNewLine
    output = output & "- **Rows**: " & Format(rowCount, "#,##0") & " data rows" & vbNewLine
    output = output & "- **Columns**: " & colCount & vbNewLine
    output = output & "- **Size**: " & sizeCategory & vbNewLine & vbNewLine
    
    ' Column details table
    output = output & "## COLUMNS FOR AI CODING" & vbNewLine
    output = output & "| # | Column Name | Data Type | Sample Values | Quality | AI Notes |" & vbNewLine
    output = output & "|---|-------------|-----------|---------------|---------|----------|" & vbNewLine
    
    Dim colIndex As Integer
    colIndex = 1
    
    For Each col In tbl.ListColumns
        Dim dataType As String
        Dim samples As String
        Dim qualityFlag As String
        Dim aiNotes As String
        Dim formulaInfo As String
        
        dataType = GetRealDataType(col)
        samples = GetSampleValues(col)
        qualityFlag = GetQualityFlag(col)
        aiNotes = GetUniversalAICodeNotes(dataType)
        
        ' Add formula information if it's a formula column
        If dataType = "Formula" Then
            formulaInfo = " | " & GetFormulaInfo(col)
        Else
            formulaInfo = ""
        End If
        
        output = output & "| " & colIndex & " | `" & col.Name & "` | " & dataType & " | " & _
                 samples & " | " & qualityFlag & " | " & aiNotes & formulaInfo & " |" & vbNewLine
        
        colIndex = colIndex + 1
    Next col
    
    output = output & vbNewLine
    
    ' Data quality section
    output = output & GetDataQualitySection(tbl)
    
    ' Formula dependencies (if any formulas exist)
    output = output & GetFormulaDependenciesSection(tbl)
    
    ' Performance considerations
    output = output & GetPerformanceSection(tbl)
    
    output = output & "---" & vbNewLine & vbNewLine
    
    ProcessTable = output
End Function

Function GetUniversalAICodeNotes(dataType As String) As String
    ' Purely agnostic notes based only on data type
    Select Case dataType
        Case "Currency", "Number"
            GetUniversalAICodeNotes = "Numeric/aggregate candidate"
        Case "Date"
            GetUniversalAICodeNotes = "Date calculations/filtering"
        Case "Text"
            GetUniversalAICodeNotes = "Category/filter or text lookup"
        Case "Formula"
            GetUniversalAICodeNotes = "Calculated field"
        Case "Empty"
            GetUniversalAICodeNotes = "Empty column - consider removing"
        Case Else
            GetUniversalAICodeNotes = "General data field"
    End Select
End Function

Function GetRealDataType(col As ListColumn) As String
    On Error GoTo ErrorHandler
    
    If col.DataBodyRange Is Nothing Then
        GetRealDataType = "Empty"
        Exit Function
    End If
    
    Dim sampleCell As Range
    Dim formulaCount As Integer
    Dim numberCount As Integer
    Dim dateCount As Integer
    Dim currencyCount As Integer
    Dim emptyCount As Integer
    Dim totalCells As Integer
    Dim checkRange As Range
    
    ' Sample up to 100 cells for performance
    Set checkRange = col.DataBodyRange
    If checkRange.Rows.Count > 100 Then
        Set checkRange = checkRange.Resize(100, 1)
    End If
    
    totalCells = checkRange.Rows.Count
    
    For Each sampleCell In checkRange
        If sampleCell.HasFormula Then
            formulaCount = formulaCount + 1
        ElseIf IsEmpty(sampleCell.Value) Or sampleCell.Value = "" Then
            emptyCount = emptyCount + 1
        ElseIf IsNumeric(sampleCell.Value) Then
            If sampleCell.NumberFormat Like "*$*" Or sampleCell.NumberFormat Like "*currency*" Then
                currencyCount = currencyCount + 1
            ElseIf IsDate(sampleCell.Value) Then
                dateCount = dateCount + 1
            Else
                numberCount = numberCount + 1
            End If
        End If
    Next sampleCell
    
    ' Determine dominant type
    If formulaCount > 0 Then
        GetRealDataType = "Formula"
    ElseIf emptyCount = totalCells Then
        GetRealDataType = "Empty"
    ElseIf currencyCount > totalCells * 0.5 Then
        GetRealDataType = "Currency"
    ElseIf dateCount > totalCells * 0.5 Then
        GetRealDataType = "Date"
    ElseIf numberCount > totalCells * 0.5 Then
        GetRealDataType = "Number"
    Else
        GetRealDataType = "Text"
    End If
    
    Exit Function
ErrorHandler:
    GetRealDataType = "Unknown"
End Function

Function GetSampleValues(col As ListColumn) As String
    On Error GoTo ErrorHandler
    
    If col.DataBodyRange Is Nothing Then
        GetSampleValues = "(empty/null)"
        Exit Function
    End If
    
    Dim samples(1) As String
    Dim sampleCount As Integer
    Dim cell As Range
    Dim checkRange As Range
    
    ' Get first few non-empty values
    Set checkRange = col.DataBodyRange
    If checkRange.Rows.Count > 10 Then
        Set checkRange = checkRange.Resize(10, 1)
    End If
    
    sampleCount = 0
    For Each cell In checkRange
        If sampleCount >= 2 Then Exit For
        If Not IsEmpty(cell.Value) And cell.Value <> "" Then
            samples(sampleCount) = CStr(cell.Value)
            If Len(samples(sampleCount)) > 20 Then
                samples(sampleCount) = Left(samples(sampleCount), 17) & "..."
            End If
            sampleCount = sampleCount + 1
        End If
    Next cell
    
    If sampleCount = 0 Then
        GetSampleValues = "(empty/null)"
    ElseIf sampleCount = 1 Then
        GetSampleValues = samples(0)
    Else
        GetSampleValues = samples(0) & ", " & samples(1)
    End If
    
    Exit Function
ErrorHandler:
    GetSampleValues = "(error reading)"
End Function

Function GetQualityFlag(col As ListColumn) As String
    On Error GoTo ErrorHandler
    
    If col.DataBodyRange Is Nothing Then
        GetQualityFlag = "ERROR: 100% empty"
        Exit Function
    End If
    
    Dim totalRows As Long
    Dim emptyCount As Long
    Dim cell As Range
    Dim emptyPercentage As Double
    
    totalRows = col.DataBodyRange.Rows.Count
    
    For Each cell In col.DataBodyRange
        If IsEmpty(cell.Value) Or cell.Value = "" Then
            emptyCount = emptyCount + 1
        End If
    Next cell
    
    emptyPercentage = (emptyCount / totalRows) * 100
    
    If emptyPercentage >= 80 Then
        GetQualityFlag = "ERROR: " & Round(emptyPercentage) & "% empty"
    ElseIf emptyPercentage >= 10 Then
        GetQualityFlag = "WARNING: " & Round(emptyPercentage) & "% empty"
    Else
        GetQualityFlag = "CLEAN"
    End If
    
    Exit Function
ErrorHandler:
    GetQualityFlag = "ERROR: Cannot assess"
End Function

Function GetFormulaInfo(col As ListColumn) As String
    On Error GoTo ErrorHandler
    
    If col.DataBodyRange Is Nothing Then
        GetFormulaInfo = ""
        Exit Function
    End If
    
    Dim firstFormulaCell As Range
    Set firstFormulaCell = col.DataBodyRange.Cells(1, 1)
    
    If firstFormulaCell.HasFormula Then
        Dim formulaText As String
        Dim category As String
        
        formulaText = firstFormulaCell.Formula
        category = GetFormulaCategory(formulaText)
        
        GetFormulaInfo = "Formula field: " & category & " | Formula: " & formulaText
    Else
        GetFormulaInfo = ""
    End If
    
    Exit Function
ErrorHandler:
    GetFormulaInfo = "Formula field: Error reading formula"
End Function

Function GetFormulaCategory(formula As String) As String
    Dim upperFormula As String
    upperFormula = UCase(formula)
    
    If InStr(upperFormula, "IF(") > 0 Or InStr(upperFormula, "IFS(") > 0 Then
        GetFormulaCategory = "Conditional (IF)"
    ElseIf InStr(upperFormula, "VLOOKUP") > 0 Or InStr(upperFormula, "XLOOKUP") > 0 Or InStr(upperFormula, "INDEX") > 0 Then
        GetFormulaCategory = "Lookup"
    ElseIf InStr(upperFormula, "SUM") > 0 Or InStr(upperFormula, "AVERAGE") > 0 Or InStr(upperFormula, "COUNT") > 0 Then
        GetFormulaCategory = "Aggregation"
    ElseIf InStr(upperFormula, "CONCATENATE") > 0 Or InStr(upperFormula, "&") > 0 Then
        GetFormulaCategory = "Text"
    ElseIf InStr(upperFormula, "DATE") > 0 Or InStr(upperFormula, "TODAY") > 0 Or InStr(upperFormula, "NOW") > 0 Then
        GetFormulaCategory = "Date/Time"
    Else
        GetFormulaCategory = "Calculation"
    End If
End Function

Function GetDataQualitySection(tbl As ListObject) As String
    Dim output As String
    Dim col As ListColumn
    Dim hasQualityIssues As Boolean
    
    output = "## DATA QUALITY FOR AI" & vbNewLine
    
    For Each col In tbl.ListColumns
        Dim qualityFlag As String
        qualityFlag = GetQualityFlag(col)
        
        If qualityFlag <> "CLEAN" Then
            output = output & "- WARNING: **" & col.Name & "**: " & qualityFlag & vbNewLine
            hasQualityIssues = True
        End If
    Next col
    
    If Not hasQualityIssues Then
        output = output & "- All columns have good data quality (CLEAN status)" & vbNewLine
    End If
    
    output = output & vbNewLine & vbNewLine
    GetDataQualitySection = output
End Function

Function GetFormulaDependenciesSection(tbl As ListObject) As String
    Dim output As String
    Dim col As ListColumn
    Dim hasDependencies As Boolean
    
    For Each col In tbl.ListColumns
        If GetRealDataType(col) = "Formula" Then
            Dim deps As String
            deps = GetFormulaDependencies(col)
            If deps <> "" Then
                If Not hasDependencies Then
                    output = "## FORMULA DEPENDENCIES" & vbNewLine
                    hasDependencies = True
                End If
                output = output & "- **" & tbl.Name & "[" & col.Name & "]** -> " & deps & vbNewLine
            End If
        End If
    Next col
    
    If hasDependencies Then
        output = output & vbNewLine & vbNewLine
    End If
    
    GetFormulaDependenciesSection = output
End Function

Function GetFormulaDependencies(col As ListColumn) As String
    On Error GoTo ErrorHandler
    
    If col.DataBodyRange Is Nothing Then
        GetFormulaDependencies = ""
        Exit Function
    End If
    
    Dim firstCell As Range
    Set firstCell = col.DataBodyRange.Cells(1, 1)
    
    If firstCell.HasFormula Then
        Dim formula As String
        Dim dependencies As String
        
        formula = UCase(firstCell.Formula)
        
        ' Look for table references
        If InStr(formula, "[@") > 0 Then
            dependencies = dependencies & "Uses same-row columns, "
        End If
        
        ' Look for other table references
        Dim ws As Worksheet
        Set ws = col.Parent.Parent
        Dim otherTbl As ListObject
        
        For Each otherTbl In ws.ListObjects
            If otherTbl.Name <> col.Parent.Name Then
                If InStr(formula, otherTbl.Name) > 0 Then
                    dependencies = dependencies & "Depends on **" & otherTbl.Name & "** table, "
                End If
            End If
        Next otherTbl
        
        ' Clean up trailing comma
        If Right(dependencies, 2) = ", " Then
            dependencies = Left(dependencies, Len(dependencies) - 2)
        End If
        
        GetFormulaDependencies = dependencies
    Else
        GetFormulaDependencies = ""
    End If
    
    Exit Function
ErrorHandler:
    GetFormulaDependencies = "Error analyzing dependencies"
End Function

Function GetPerformanceSection(tbl As ListObject) As String
    Dim output As String
    Dim rowCount As Long
    Dim formulaCount As Integer
    Dim col As ListColumn
    
    rowCount = GetTotalRows(tbl)
    
    ' Count formula columns
    For Each col In tbl.ListColumns
        If GetRealDataType(col) = "Formula" Then
            formulaCount = formulaCount + 1
        End If
    Next col
    
    output = "## PERFORMANCE CONSIDERATIONS" & vbNewLine
    output = output & "- **" & GetSizeCategory(rowCount) & " dataset** (" & Format(rowCount, "#,##0") & " rows) - " & GetPerformanceAdvice(rowCount) & vbNewLine
    
    If formulaCount > 0 Then
        output = output & "- **Formula columns present** (" & formulaCount & ") - these recalculate automatically with data changes" & vbNewLine
    End If
    
    output = output & vbNewLine & vbNewLine
    GetPerformanceSection = output
End Function

' === UTILITY FUNCTIONS ===

Function GetTotalRows(tbl As ListObject) As Long
    If tbl.DataBodyRange Is Nothing Then
        GetTotalRows = 0
    Else
        GetTotalRows = tbl.DataBodyRange.Rows.Count
    End If
End Function

Function GetSizeCategory(rowCount As Long) As String
    If rowCount < 100 Then
        GetSizeCategory = "Small"
    ElseIf rowCount < 1000 Then
        GetSizeCategory = "Medium"
    ElseIf rowCount < 10000 Then
        GetSizeCategory = "Large"
    Else
        GetSizeCategory = "Very Large"
    End If
End Function

Function GetPerformanceAdvice(rowCount As Long) As String
    If rowCount < 1000 Then
        GetPerformanceAdvice = "standard Excel functions work efficiently"
    ElseIf rowCount < 10000 Then
        GetPerformanceAdvice = "structured references recommended"
    Else
        GetPerformanceAdvice = "consider INDEX/MATCH over VLOOKUP for better performance"
    End If
End Function

Function GetOverallPerformanceAdvice(totalRows As Long, tableCount As Integer) As String
    If totalRows < 5000 And tableCount < 5 Then
        GetOverallPerformanceAdvice = "Standard Excel functions should work well"
    ElseIf totalRows < 50000 Then
        GetOverallPerformanceAdvice = "Use structured references and consider performance optimization"
    Else
        GetOverallPerformanceAdvice = "Large dataset - optimize formulas and consider data model approach"
    End If
End Function

Function ChooseOutputLocation() As String
    Dim defaultPath As String
    Dim userPath As String
    
    ' Create default path
    defaultPath = Environ("USERPROFILE") & "\Downloads\AI_Table_Guide_" & Format(Now, "YYYYMMDD_HHMMSS") & ".txt"
    
    ' Simple input box approach (more reliable across Excel versions)
    userPath = InputBox("Enter the full path where you want to save the AI documentation file:" & vbNewLine & vbNewLine & _
                       "Default location:", "Save AI Table Documentation", defaultPath)
    
    ' If user clicked Cancel or entered empty string
    If userPath = "" Then
        ChooseOutputLocation = ""
        Exit Function
    End If
    
    ' Ensure .txt extension
    If Right(LCase(userPath), 4) <> ".txt" Then
        userPath = userPath & ".txt"
    End If
    
    ChooseOutputLocation = userPath
End Function

Sub SaveOutput(content As String, filePath As String)
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open filePath For Output As #fileNum
    Print #fileNum, content
    Close #fileNum
End Sub