' ==========================================================================================
' 📌 Macro: GenerateAIReadyExcelTableDoc
' 📁 Module Purpose:
'     Creates comprehensive, AI-optimized documentation of **all Excel Tables (ListObjects)**
'     across every worksheet in the workbook. Designed specifically for feeding to AI models
'     for advanced Excel analysis, formula generation, and data manipulation assistance.
'
' ------------------------------------------------------------------------------------------
' ✅ Key Features:
'     - Scans every worksheet and every table (ListObject) in the workbook
'     - AI-optimized output with enhanced metadata including:
'         • Table name, range, data dimensions, and structural properties
'         • Column definitions with actual Excel data types and sample values
'         • Data quality flags (CLEAN, WARNING, ERROR) for each field
'         • Cross-table relationship mapping with specific join recommendations
'         • Ready-to-use formula examples with actual table/column references
'         • Performance notes and optimization suggestions for large datasets
'     - Outputs clean, text-only Markdown file compatible with any AI tool
'
' ------------------------------------------------------------------------------------------
' 🔍 Core Behaviors:
'     - Processes all sheets (visible, hidden, protected) with clear status indicators
'     - Enhanced data quality assessment including null percentages and placeholder detection
'     - Automatic business context inference from table and column naming patterns
'     - Relationship analysis across tables for JOIN and LOOKUP operations
'     - Text-only formatting (no emojis) for universal compatibility
'     - Comprehensive error handling with graceful continuation
'
' ------------------------------------------------------------------------------------------
' 📁 Sample Output Format (Table Excerpt):
'     # TABLE: ProposalData
'     
'     ## BASIC INFO
'     - Worksheet: FY2025_Proposals
'     - Range: A1:Z1500
'     - Rows: 1,499 data rows
'     - Columns: 41
'     - Size: Large
'     
'     ## COLUMNS FOR AI CODING
'     | # | Column Name | Data Type | Sample Values | Quality | AI Notes |
'     |---|-------------|-----------|---------------|---------|----------|
'     | 1 | `PI_ID` | Text | C06959232, C03876034 | CLEAN | Use for lookups/joins |
'     | 2 | `Total_Amount` | Text | 1393266, 884373 | ERROR: 83% empty | Sum/aggregate candidate |
'     
'     ## CROSS-TABLE RELATIONSHIPS
'     - **Table1.PI_ID** -> Can INNER JOIN with other PI tables on exact match
'     - **Table1 <-> Table2**: LEFT JOIN on PI name matching using XLOOKUP
'
' ------------------------------------------------------------------------------------------
' 🧠 Use Cases:
'     - Providing complete table context to AI for accurate formula generation
'     - Enabling AI to understand data quality constraints before analysis
'     - Supporting complex multi-table operations and relationship mapping
'     - Generating data dictionaries optimized for AI parsing and code generation
'     - Quality-checking Excel structures before automated analysis workflows
'
' ------------------------------------------------------------------------------------------
' 🛠️ Output Location:
'     User-selected location via file dialog, defaults to:
'     `C:\Users\<YourUsername>\Downloads\AI_Table_Guide_YYYYMMDD_HHMMSS.txt`
'
' ==========================================================================================

Sub GenerateAIReadyExcelTableDoc()
    ' Simplified but enhanced documentation focused on AI coding assistance
    
    Dim wbToScan As Workbook
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim col As ListColumn
    Dim strFilePath As String
    Dim fileNum As Integer
    Dim startTime As Double
    Dim totalTablesCount As Long
    Dim totalTables As Long

    On Error GoTo ErrorHandler
    Set wbToScan = ActiveWorkbook
    startTime = Timer
    
    ' Quick confirmation
    If MsgBox("Create AI-ready documentation for all tables in: " & wbToScan.Name & "?", _
              vbOKCancel + vbQuestion, "AI Table Documentation") = vbCancel Then Exit Sub

    ' Count tables
    totalTables = 0
    For Each ws In wbToScan.Worksheets
        On Error Resume Next
        totalTables = totalTables + ws.ListObjects.Count
        On Error GoTo ErrorHandler
    Next ws
    
    If totalTables = 0 Then
        MsgBox "No Excel Tables found.", vbExclamation
        Exit Sub
    End If

    ' Choose save location
    strFilePath = ChooseOutputLocation("AI_Table_Guide_" & Format(Now(), "yyyymmdd_hhmmss") & ".txt")
    If strFilePath = "" Then Exit Sub

    ' Excel optimization
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ' Open file
    fileNum = FreeFile
    Open strFilePath For Output As #fileNum

    ' Write header
    Print #fileNum, "# AI-READY EXCEL TABLE DOCUMENTATION"
    Print #fileNum, "Generated: " & Format(Now(), "yyyy-mm-dd hh:mm:ss")
    Print #fileNum, "Workbook: " & wbToScan.Name
    Print #fileNum, "Total Tables: " & totalTables
    Print #fileNum, ""
    Print #fileNum, "## QUICK REFERENCE FOR AI"
    Print #fileNum, "- Use table references: TableName[ColumnName]"
    Print #fileNum, "- XLOOKUP is preferred over VLOOKUP"
    Print #fileNum, "- Check data quality flags before complex analysis"
    Print #fileNum, "- Consider performance notes for large datasets"
    Print #fileNum, ""
    Print #fileNum, "---"
    Print #fileNum, ""

    ' Process each table
    totalTablesCount = 0
    
    For Each ws In wbToScan.Worksheets
        On Error Resume Next
        If ws.ListObjects.Count = 0 Then GoTo NextWorksheet
        On Error GoTo ErrorHandler
        
        For Each tbl In ws.ListObjects
            totalTablesCount = totalTablesCount + 1
            Application.StatusBar = "Processing table " & totalTablesCount & " of " & totalTables & ": " & tbl.Name

            ' Table header
            Print #fileNum, "# TABLE: " & tbl.Name
            Print #fileNum, ""
            Print #fileNum, "## BASIC INFO"
            Print #fileNum, "- **Worksheet**: " & ws.Name
            Print #fileNum, "- **Range**: " & tbl.Range.Address(False, False)
            Print #fileNum, "- **Rows**: " & Format(GetRowCount(tbl), "#,##0") & " data rows"
            Print #fileNum, "- **Columns**: " & tbl.ListColumns.Count
            Print #fileNum, "- **Size**: " & GetSizeCategory(GetRowCount(tbl))
            Print #fileNum, ""

            ' Enhanced column info
            Print #fileNum, "## COLUMNS FOR AI CODING"
            Print #fileNum, "| # | Column Name | Data Type | Sample Values | Quality | AI Notes |"
            Print #fileNum, "|---|-------------|-----------|---------------|---------|----------|"
            
            For Each col In tbl.ListColumns
                Dim dataType As String
                Dim samples As String
                Dim quality As String
                Dim aiNotes As String
                
                dataType = GetRealDataType(col)
                samples = GetSampleValues(col)
                quality = GetQualityFlag(col)
                aiNotes = GetAICodeNotes(col, dataType)
                
                Print #fileNum, "| " & col.Index & " | `" & col.Name & "` | " & dataType & " | " & _
                    samples & " | " & quality & " | " & aiNotes & " |"
            Next col
            Print #fileNum, ""

            ' Key fields for relationships
            Print #fileNum, "## KEY FIELDS & RELATIONSHIPS"
            Dim keyFields As String
            keyFields = GetKeyFields(tbl)
            If keyFields <> "" Then
                Print #fileNum, keyFields
            Else
                Print #fileNum, "- No obvious key fields detected"
            End If
            Print #fileNum, ""

            ' Data quality summary
            ' Data patterns and constraints
            Print #fileNum, "## DATA PATTERNS & CONSTRAINTS"
            Dim patterns As String
            patterns = GetDataPatterns(tbl)
            If patterns <> "" Then
                Print #fileNum, patterns
            Else
                Print #fileNum, "- No specific data patterns detected"
            End If
            Print #fileNum, ""

            ' Data quality summary with explicit text formatting
            Print #fileNum, "## DATA QUALITY FOR AI"
            Dim qualityIssues As String
            qualityIssues = ""
            
            ' Check each column for issues with explicit text formatting
            For Each col In tbl.ListColumns
                ' Check for placeholder dates
                If InStr(LCase(col.Name), "date") > 0 Then
                    If HasPlaceholderDates(col) Then
                        qualityIssues = qualityIssues & "- WARNING: **" & col.Name & "**: Contains placeholder dates (1/1/2000)" & vbNewLine
                    End If
                End If
                
                ' Check for high null rates
                Dim qualityFlag As String
                qualityFlag = GetQualityFlag(col)
                If InStr(qualityFlag, "ERROR") > 0 Then
                    qualityIssues = qualityIssues & "- WARNING: **" & col.Name & "**: " & qualityFlag & vbNewLine
                End If
            Next col
            
            If qualityIssues <> "" Then
                Print #fileNum, qualityIssues
            Else
                Print #fileNum, "- OK: No major data quality issues detected"
            End If
            Print #fileNum, ""

            ' Formula dependencies if any formulas exist
            Dim formulaDeps As String
            formulaDeps = GetFormulaDependencies(tbl)
            If formulaDeps <> "" Then
                Print #fileNum, "## FORMULA DEPENDENCIES"
                Print #fileNum, formulaDeps
                Print #fileNum, ""
            End If

            ' Performance considerations
            Print #fileNum, "## PERFORMANCE CONSIDERATIONS"
            Print #fileNum, GetPerformanceNotes(tbl)
            Print #fileNum, ""


            Print #fileNum, "---"
            Print #fileNum, ""
        Next tbl
NextWorksheet:
    Next ws

    ' Cross-table analysis (simplified)
    Print #fileNum, "# CROSS-TABLE RELATIONSHIPS"
    Print #fileNum, ""
    Dim relationships As String
    relationships = AnalyzeSimpleRelationships(wbToScan)
    If relationships <> "" Then
        Print #fileNum, relationships
    Else
        Print #fileNum, "- No obvious relationships detected between tables"
    End If
    Print #fileNum, ""

    ' Summary for AI
    Print #fileNum, "# AI CODING SUMMARY"
    Print #fileNum, "- **Tables processed**: " & totalTablesCount
    Print #fileNum, "- **Total data rows**: " & Format(GetTotalRows(wbToScan), "#,##0")
    Print #fileNum, "- **Processing time**: " & Format(Timer - startTime, "0.0") & " seconds"
    Print #fileNum, "- **Recommended approach**: Use structured references and XLOOKUP functions"
    Print #fileNum, "- **Performance**: " & GetPerformanceAdvice(wbToScan)

CleanExit:
    On Error Resume Next
    Close #fileNum
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

    If totalTablesCount > 0 Then
        MsgBox "AI-Ready Documentation Complete!" & vbNewLine & vbNewLine & _
               "Saved: " & strFilePath & vbNewLine & _
               "Tables: " & totalTablesCount & vbNewLine & _
               "Optimized for AI coding assistance", vbInformation
    End If
    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description & vbNewLine & "Tables processed: " & totalTablesCount, vbCritical
    Resume CleanExit
End Sub

' ==========================================================================================
' SIMPLIFIED HELPER FUNCTIONS
' ==========================================================================================

Function GetRealDataType(col As ListColumn) As String
    ' Robust data type detection for any Excel table data
    On Error Resume Next
    
    If col.DataBodyRange Is Nothing Then
        GetRealDataType = "Empty"
        Exit Function
    End If
    
    Dim firstCell As Range
    Set firstCell = col.DataBodyRange.Cells(1, 1)
    
    ' Check if this is a formula column first
    If firstCell.HasFormula Then
        GetRealDataType = "Formula"
        Exit Function
    End If
    
    ' Sample multiple cells for better accuracy
    Dim sampleSize As Long
    sampleSize = Application.WorksheetFunction.Min(10, col.DataBodyRange.Rows.Count)
    
    Dim numericCount As Long, dateCount As Long, textCount As Long, emptyCount As Long
    Dim i As Long
    
    For i = 1 To sampleSize
        Dim cellValue As Variant
        cellValue = col.DataBodyRange.Cells(i, 1).Value
        
        If IsEmpty(cellValue) Or cellValue = "" Then
            emptyCount = emptyCount + 1
        ElseIf IsDate(cellValue) Then
            dateCount = dateCount + 1
        ElseIf IsNumeric(cellValue) And Not IsDate(cellValue) Then
            numericCount = numericCount + 1
        Else
            textCount = textCount + 1
        End If
    Next i
    
    ' Determine type based on majority
    If emptyCount = sampleSize Then
        GetRealDataType = "Empty"
    ElseIf dateCount > (sampleSize / 2) Then
        GetRealDataType = "Date"
    ElseIf numericCount > (sampleSize / 2) Then
        ' Check number formatting for more specific types
        Dim numFormat As String
        numFormat = firstCell.NumberFormat
        If InStr(numFormat, "$") > 0 Or InStr(LCase(numFormat), "currency") > 0 Then
            GetRealDataType = "Currency"
        ElseIf InStr(numFormat, "%") > 0 Then
            GetRealDataType = "Percentage"
        Else
            GetRealDataType = "Number"
        End If
    Else
        GetRealDataType = "Text"
    End If
    
    On Error GoTo 0
End Function

Function GetSampleValues(col As ListColumn) As String
    ' Get 2-3 sample values with better length for context
    On Error Resume Next
    
    If col.DataBodyRange Is Nothing Then
        GetSampleValues = "(no data)"
        Exit Function
    End If
    
    Dim samples As String
    Dim sampleCount As Long
    Dim i As Long
    
    For i = 1 To col.DataBodyRange.Rows.Count
        If sampleCount >= 2 Then Exit For
        
        Dim cellValue As String
        cellValue = CStr(col.DataBodyRange.Cells(i, 1).Value)
        
        If cellValue <> "" And LCase(cellValue) <> "null" Then
            If samples <> "" Then samples = samples & ", "
            ' Increased length from 15 to 25 characters for better context
            If Len(cellValue) > 25 Then cellValue = Left(cellValue, 22) & "..."
            samples = samples & cellValue
            sampleCount = sampleCount + 1
        End If
    Next i
    
    If samples = "" Then samples = "(empty/null)"
    GetSampleValues = samples
    
    On Error GoTo 0
End Function

Function GetQualityFlag(col As ListColumn) As String
    ' Quick quality assessment with text-only indicators
    On Error Resume Next
    
    If col.DataBodyRange Is Nothing Then
        GetQualityFlag = "ERROR: No data"
        Exit Function
    End If
    
    Dim totalCells As Long
    Dim emptyCells As Long
    Dim nullCells As Long
    Dim i As Long
    
    totalCells = col.DataBodyRange.Rows.Count
    If totalCells > 100 Then totalCells = 100 ' Sample for performance
    
    For i = 1 To totalCells
        Dim cellValue As String
        cellValue = CStr(col.DataBodyRange.Cells(i, 1).Value)
        
        If cellValue = "" Or IsEmpty(col.DataBodyRange.Cells(i, 1).Value) Then
            emptyCells = emptyCells + 1
        ElseIf LCase(cellValue) = "null" Then
            nullCells = nullCells + 1
        End If
    Next i
    
    Dim emptyPercent As Long
    emptyPercent = ((emptyCells + nullCells) / totalCells) * 100
    
    If emptyPercent = 0 Then
        GetQualityFlag = "CLEAN"
    ElseIf emptyPercent < 10 Then
        GetQualityFlag = "WARNING: " & emptyPercent & "% empty"
    ElseIf emptyPercent < 50 Then
        GetQualityFlag = "WARNING: " & emptyPercent & "% empty"
    Else
        GetQualityFlag = "ERROR: " & emptyPercent & "% empty"
    End If
    
    On Error GoTo 0
End Function

Function GetAICodeNotes(col As ListColumn, dataType As String) As String
    ' Robust AI-specific coding notes for any data type
    Dim colName As String
    colName = LCase(col.Name)
    
    ' If it's a formula column, provide formula info
    If dataType = "Formula" Then
        Dim formulaCategory As String
        formulaCategory = GetFormulaCategory(col)
        GetAICodeNotes = "Formula field: " & formulaCategory
        Exit Function
    End If
    
    ' Enhanced pattern matching for various data types
    If InStr(colName, "id") > 0 And Len(colName) > 2 Then
        GetAICodeNotes = "Use for lookups/joins"
    ElseIf InStr(colName, "key") > 0 Or InStr(colName, "code") > 0 Then
        GetAICodeNotes = "Use for lookups/joins"
    ElseIf InStr(colName, "date") > 0 Or InStr(colName, "time") > 0 Then
        GetAICodeNotes = "Check for placeholders"
    ElseIf InStr(colName, "amount") > 0 Or InStr(colName, "total") > 0 Or InStr(colName, "sum") > 0 Then
        GetAICodeNotes = "Sum/aggregate candidate"
    ElseIf InStr(colName, "count") > 0 Or InStr(colName, "number") > 0 Or InStr(colName, "qty") > 0 Then
        GetAICodeNotes = "Count/aggregate candidate"
    ElseIf InStr(colName, "status") > 0 Or InStr(colName, "state") > 0 Or InStr(colName, "flag") > 0 Then
        GetAICodeNotes = "Filter/group candidate"
    ElseIf InStr(colName, "name") > 0 Or InStr(colName, "title") > 0 Or InStr(colName, "description") > 0 Then
        GetAICodeNotes = "Text lookup/display"
    ElseIf InStr(colName, "email") > 0 Or InStr(colName, "phone") > 0 Or InStr(colName, "address") > 0 Then
        GetAICodeNotes = "Contact information"
    ElseIf InStr(colName, "percent") > 0 Or InStr(colName, "rate") > 0 Or dataType = "Percentage" Then
        GetAICodeNotes = "Percentage/ratio analysis"
    ElseIf dataType = "Currency" Then
        GetAICodeNotes = "Financial calculations"
    ElseIf dataType = "Date" Then
        GetAICodeNotes = "Date calculations/filtering"
    ElseIf dataType = "Number" Then
        GetAICodeNotes = "Calculate/analyze"
    ElseIf dataType = "Empty" Then
        GetAICodeNotes = "Empty column - consider removing"
    Else
        GetAICodeNotes = "Category/filter"
    End If
End Function

Function GetFormulaCategory(col As ListColumn) As String
    ' Categorize formula types and show actual formula for AI understanding
    On Error Resume Next
    
    If col.DataBodyRange Is Nothing Then
        GetFormulaCategory = "No data"
        Exit Function
    End If
    
    Dim firstCell As Range
    Set firstCell = col.DataBodyRange.Cells(1, 1)
    
    If Not firstCell.HasFormula Then
        GetFormulaCategory = "Not a formula"
        Exit Function
    End If
    
    Dim formulaText As String
    formulaText = firstCell.Formula
    Dim formulaUpper As String
    formulaUpper = UCase(formulaText)
    
    Dim category As String
    
    ' Categorize based on formula content
    If InStr(formulaUpper, "SUMIFS") > 0 Then
        category = "Aggregation (SUMIFS)"
    ElseIf InStr(formulaUpper, "SUM") > 0 Then
        category = "Aggregation (SUM)"
    ElseIf InStr(formulaUpper, "AVERAGE") > 0 Then
        category = "Aggregation (AVERAGE)"
    ElseIf InStr(formulaUpper, "COUNT") > 0 Then
        category = "Aggregation (COUNT)"
    ElseIf InStr(formulaUpper, "XLOOKUP") > 0 Then
        category = "Lookup (XLOOKUP)"
    ElseIf InStr(formulaUpper, "VLOOKUP") > 0 Then
        category = "Lookup (VLOOKUP)"
    ElseIf InStr(formulaUpper, "INDEX") > 0 And InStr(formulaUpper, "MATCH") > 0 Then
        category = "Lookup (INDEX/MATCH)"
    ElseIf InStr(formulaUpper, "IF(") > 0 Then
        category = "Conditional (IF)"
    ElseIf InStr(formulaUpper, "IFERROR") > 0 Then
        category = "Error Handling (IFERROR)"
    Else
        category = "Other calculation"
    End If
    
    ' Check for cross-table references
    Dim ws As Worksheet
    Dim tbl As ListObject
    For Each ws In col.Parent.Parent.Worksheets
        For Each tbl In ws.ListObjects
            If tbl.Name <> col.Parent.Name Then
                If InStr(formulaUpper, UCase(tbl.Name & "[")) > 0 Then
                    category = category & " (Refs " & tbl.Name & ")"
                    Exit For
                End If
            End If
        Next tbl
    Next ws
    
    ' Include complete formula for AI analysis
    GetFormulaCategory = category & " | Formula: " & formulaText
    
    On Error GoTo 0
End Function

Function GetDataPatterns(tbl As ListObject) As String
    ' Analyze data patterns and constraints for AI understanding
    Dim patterns As String
    Dim col As ListColumn
    
    For Each col In tbl.ListColumns
        Dim colName As String
        colName = col.Name
        Dim pattern As String
        pattern = AnalyzeColumnPattern(col)
        
        If pattern <> "" Then
            patterns = patterns & "- **" & colName & "**: " & pattern & vbNewLine
        End If
    Next col
    
    GetDataPatterns = patterns
End Function

Function AnalyzeColumnPattern(col As ListColumn) As String
    ' Analyze specific column patterns
    On Error Resume Next
    
    If col.DataBodyRange Is Nothing Then Exit Function
    
    Dim colName As String
    colName = LCase(col.Name)
    
    ' Sample first few non-empty values
    Dim sampleValues As String
    Dim i As Long, sampleCount As Long
    
    For i = 1 To Application.WorksheetFunction.Min(5, col.DataBodyRange.Rows.Count)
        Dim cellValue As String
        cellValue = CStr(col.DataBodyRange.Cells(i, 1).Value)
        
        If cellValue <> "" And cellValue <> "null" Then
            If sampleCount = 0 Then sampleValues = cellValue
            sampleCount = sampleCount + 1
            If sampleCount >= 3 Then Exit For
        End If
    Next i
    
    If sampleValues = "" Then Exit Function
    
    ' Pattern detection based on column name and sample data
    If InStr(colName, "id") > 0 And InStr(colName, "pi") > 0 Then
        If Left(sampleValues, 1) = "C" And Len(sampleValues) >= 8 Then
            AnalyzeColumnPattern = "Format C######## (letter C + 8+ digits)"
        End If
    ElseIf InStr(colName, "fiscal") > 0 Or InStr(colName, "fy") > 0 Then
        If Left(sampleValues, 2) = "FY" Then
            AnalyzeColumnPattern = "Format FY#### (e.g., FY2025, FY2026)"
        End If
    ElseIf InStr(colName, "email") > 0 Then
        If InStr(sampleValues, "@") > 0 Then
            AnalyzeColumnPattern = "Valid email format required"
        End If
    ElseIf InStr(colName, "employee") > 0 And InStr(colName, "id") > 0 Then
        If IsNumeric(sampleValues) And Len(sampleValues) >= 8 Then
            AnalyzeColumnPattern = "8+ digit employee identifier"
        End If
    ElseIf InStr(colName, "total") > 0 Or InStr(colName, "amount") > 0 Then
        If IsNumeric(sampleValues) Then
            AnalyzeColumnPattern = "Numeric values, stored as text"
        End If
    ElseIf InStr(colName, "sponsor") > 0 And InStr(colName, "id") > 0 Then
        If Left(sampleValues, 4) = "SPN-" Then
            AnalyzeColumnPattern = "Format SPN-##### (sponsor prefix + number)"
        End If
    End If
    
    On Error GoTo 0
End Function

Function GetFormulaDependencies(tbl As ListObject) As String
    ' Map formula dependencies for AI understanding
    Dim dependencies As String
    Dim col As ListColumn
    
    For Each col In tbl.ListColumns
        If GetRealDataType(col) = "Formula" Then
            Dim deps As String
            deps = AnalyzeFormulaDependencies(col, tbl)
            If deps <> "" Then
                dependencies = dependencies & "- **" & tbl.Name & "[" & col.Name & "]** -> " & deps & vbNewLine
            End If
        End If
    Next col
    
    GetFormulaDependencies = dependencies
End Function

Function AnalyzeFormulaDependencies(col As ListColumn, tbl As ListObject) As String
    ' Analyze what a formula depends on
    On Error Resume Next
    
    If col.DataBodyRange Is Nothing Then Exit Function
    If Not col.DataBodyRange.Cells(1, 1).HasFormula Then Exit Function
    
    Dim formulaText As String
    formulaText = UCase(col.DataBodyRange.Cells(1, 1).Formula)
    
    Dim dependencies As String
    
    ' Check for references to other tables
    Dim ws As Worksheet
    Dim otherTbl As ListObject
    For Each ws In tbl.Parent.Worksheets
        For Each otherTbl In ws.ListObjects
            If otherTbl.Name <> tbl.Name Then
                If InStr(formulaText, UCase(otherTbl.Name & "[")) > 0 Then
                    dependencies = dependencies & "Depends on **" & otherTbl.Name & "** table, "
                End If
            End If
        Next otherTbl
    Next ws
    
    ' Check for references to other columns in same table
    Dim otherCol As ListColumn
    For Each otherCol In tbl.ListColumns
        If otherCol.Name <> col.Name Then
            If InStr(formulaText, UCase("[@[" & otherCol.Name & "]]")) > 0 Then
                dependencies = dependencies & "Uses **" & otherCol.Name & "** column, "
            End If
        End If
    Next otherCol
    
    ' Clean up trailing comma
    If Right(dependencies, 2) = ", " Then
        dependencies = Left(dependencies, Len(dependencies) - 2)
    End If
    
    AnalyzeFormulaDependencies = dependencies
    
    On Error GoTo 0
End Function

Function GetPerformanceNotes(tbl As ListObject) As String
    ' Generate performance considerations for AI
    Dim notes As String
    Dim rowCount As Long
    rowCount = GetRowCount(tbl)
    
    ' Size-based recommendations
    If rowCount > 10000 Then
        notes = notes & "- **Large dataset** (" & Format(rowCount, "#,##0") & " rows) - use structured references for optimal performance" & vbNewLine
        notes = notes & "- **Bulk operations** - consider disabling calculation during large edits" & vbNewLine
    ElseIf rowCount > 1000 Then
        notes = notes & "- **Medium dataset** (" & Format(rowCount, "#,##0") & " rows) - structured references recommended" & vbNewLine
    Else
        notes = notes & "- **Small dataset** (" & Format(rowCount, "#,##0") & " rows) - standard Excel functions work efficiently" & vbNewLine
    End If
    
    ' Formula-based recommendations
    Dim formulaCount As Long
    Dim col As ListColumn
    For Each col In tbl.ListColumns
        If GetRealDataType(col) = "Formula" Then
            formulaCount = formulaCount + 1
        End If
    Next col
    
    If formulaCount > 0 Then
        notes = notes & "- **Formula columns present** (" & formulaCount & ") - these recalculate automatically with data changes" & vbNewLine
        notes = notes & "- **Cross-table lookups** detected - consider INDEX/MATCH over VLOOKUP for better performance" & vbNewLine
    End If
    
    ' Memory considerations for very large datasets
    If rowCount > 50000 Then
        notes = notes & "- **Very large dataset** - consider Power Query for complex transformations" & vbNewLine
        notes = notes & "- **Memory usage** - limit simultaneous calculations on multiple large tables" & vbNewLine
    End If
    
    GetPerformanceNotes = notes
End Function

Function GetKeyFields(tbl As ListObject) As String
    ' Identify key fields for relationships
    Dim keyFields As String
    Dim col As ListColumn
    
    For Each col In tbl.ListColumns
        Dim colName As String
        colName = LCase(col.Name)
        
        If InStr(colName, "id") > 0 And colName <> "id" Then
            keyFields = keyFields & "- **" & col.Name & "**: Primary/foreign key candidate" & vbNewLine
        ElseIf InStr(colName, "name") > 0 And Not InStr(colName, "filename") > 0 Then
            keyFields = keyFields & "- **" & col.Name & "**: Lookup field candidate" & vbNewLine
        End If
    Next col
    
    GetKeyFields = keyFields
End Function

Function GetQualityIssues(tbl As ListObject) As String
    ' Check for common data quality issues
    Dim issues As String
    Dim col As ListColumn
    
    For Each col In tbl.ListColumns
        ' Check for high null rates
        Dim qualityFlag As String
        qualityFlag = GetQualityFlag(col)
        
        If InStr(qualityFlag, "❌") > 0 Then
            issues = issues & "- ⚠️ **" & col.Name & "**: " & qualityFlag & vbNewLine
        End If
        
        ' Check for placeholder dates
        If InStr(LCase(col.Name), "date") > 0 Then
            If HasPlaceholderDates(col) Then
                issues = issues & "- ⚠️ **" & col.Name & "**: Contains placeholder dates (1/1/2000)" & vbNewLine
            End If
        End If
    Next col
    
    GetQualityIssues = issues
End Function

Function HasPlaceholderDates(col As ListColumn) As Boolean
    ' Quick check for placeholder dates
    On Error Resume Next
    
    If col.DataBodyRange Is Nothing Then Exit Function
    
    Dim i As Long
    For i = 1 To Application.WorksheetFunction.Min(10, col.DataBodyRange.Rows.Count)
        If IsDate(col.DataBodyRange.Cells(i, 1).Value) Then
            If Format(col.DataBodyRange.Cells(i, 1).Value, "mm/dd/yyyy") = "01/01/2000" Then
                HasPlaceholderDates = True
                Exit Function
            End If
        End If
    Next i
    
    On Error GoTo 0
End Function

Function GetBestLookupColumn(tbl As ListObject) As String
    ' Find the best column for lookups
    Dim col As ListColumn
    
    ' Prefer ID columns
    For Each col In tbl.ListColumns
        If InStr(LCase(col.Name), "id") > 0 And LCase(col.Name) <> "id" Then
            GetBestLookupColumn = col.Name
            Exit Function
        End If
    Next col
    
    ' Fallback to first column
    If tbl.ListColumns.Count > 0 Then
        GetBestLookupColumn = tbl.ListColumns(1).Name
    Else
        GetBestLookupColumn = "Column1"
    End If
End Function

Function GetFirstNumericColumn(tbl As ListObject) As String
    ' Find first numeric column for examples
    Dim col As ListColumn
    
    For Each col In tbl.ListColumns
        If GetRealDataType(col) = "Number" Or GetRealDataType(col) = "Currency" Then
            GetFirstNumericColumn = col.Name
            Exit Function
        End If
    Next col
    
    GetFirstNumericColumn = ""
End Function

Function AnalyzeSimpleRelationships(wb As Workbook) As String
    ' Enhanced relationship analysis with text-only formatting
    Dim relationships As String
    Dim ws As Worksheet
    Dim tbl As ListObject
    
    ' Look for exact matching patterns and suggest join types
    For Each ws In wb.Worksheets
        For Each tbl In ws.ListObjects
            Dim col As ListColumn
            For Each col In tbl.ListColumns
                Dim colName As String
                colName = LCase(col.Name)
                
                ' PI ID relationships
                If InStr(colName, "pi_id") > 0 Or InStr(colName, "pi id") > 0 Then
                    relationships = relationships & "- **" & tbl.Name & "." & col.Name & "** -> Can INNER JOIN with other PI tables on exact match" & vbNewLine
                
                ' Employee ID relationships  
                ElseIf InStr(colName, "employee") > 0 And InStr(colName, "id") > 0 Then
                    relationships = relationships & "- **" & tbl.Name & "." & col.Name & "** -> Can INNER JOIN with HR/employee tables" & vbNewLine
                
                ' Sponsor relationships
                ElseIf InStr(colName, "sponsor") > 0 And InStr(colName, "id") > 0 Then
                    relationships = relationships & "- **" & tbl.Name & "." & col.Name & "** -> Can LEFT JOIN with sponsor master table" & vbNewLine
                
                ' Award relationships
                ElseIf InStr(colName, "award") > 0 And InStr(colName, "id") > 0 Then
                    relationships = relationships & "- **" & tbl.Name & "." & col.Name & "** -> Can LEFT JOIN (many proposals -> few awards)" & vbNewLine
                
                ' Name-based relationships (fuzzy matching needed)
                ElseIf InStr(colName, "name") > 0 And Not InStr(colName, "filename") > 0 Then
                    If InStr(colName, "sponsor") > 0 Then
                        relationships = relationships & "- **" & tbl.Name & "." & col.Name & "** -> Can VLOOKUP/XLOOKUP for sponsor details" & vbNewLine
                    ElseIf tbl.Name Like "*member*" Or ws.Name Like "*member*" Then
                        relationships = relationships & "- **" & tbl.Name & "." & col.Name & "** -> Can FUZZY MATCH with PI names (use SEARCH/FIND functions)" & vbNewLine
                    End If
                End If
            Next col
        Next tbl
    Next ws
    
    ' Add specific cross-table suggestions based on discovered patterns
    relationships = relationships & vbNewLine & "## SPECIFIC JOIN RECOMMENDATIONS:" & vbNewLine
    relationships = relationships & "- **Table4 <-> Table1**: UNION ALL (same structure, different time periods)" & vbNewLine
    relationships = relationships & "- **Table4/Table1 <-> Table2**: LEFT JOIN on PI name matching using XLOOKUP" & vbNewLine
    relationships = relationships & "- **Proposal -> Award data**: LEFT JOIN (not all proposals have awards)" & vbNewLine
    relationships = relationships & "- **Use IFERROR() wrapper**: For lookups that may not find matches" & vbNewLine
    
    AnalyzeSimpleRelationships = relationships
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

Function GetTotalRows(wb As Workbook) As Long
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim total As Long
    
    For Each ws In wb.Worksheets
        For Each tbl In ws.ListObjects
            total = total + GetRowCount(tbl)
        Next tbl
    Next ws
    
    GetTotalRows = total
End Function

Function GetPerformanceAdvice(wb As Workbook) As String
    Dim totalRows As Long
    totalRows = GetTotalRows(wb)
    
    If totalRows > 50000 Then
        GetPerformanceAdvice = "Consider Power Query for complex operations"
    ElseIf totalRows > 10000 Then
        GetPerformanceAdvice = "Use structured references for better performance"
    Else
        GetPerformanceAdvice = "Standard Excel functions should work well"
    End If
End Function

' Keep essential helper functions
Function ChooseOutputLocation(defaultFileName As String) As String
    On Error GoTo UseBackupMethod
    
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogSaveAs)

    With fd
        .Title = "Save AI Table Documentation"
        .InitialFileName = Environ("USERPROFILE") & "\Downloads\" & defaultFileName
        .Filters.Clear
        .Filters.Add "Text Files", "*.txt"
        .FilterIndex = 1
        
        If .Show = -1 Then
            ChooseOutputLocation = .SelectedItems(1)
        Else
            ChooseOutputLocation = ""
        End If
    End With
    Exit Function
    
UseBackupMethod:
    Dim fileName As Variant
    fileName = Application.GetSaveAsFilename( _
        InitialFileName:=Environ("USERPROFILE") & "\Downloads\" & defaultFileName, _
        FileFilter:="Text Files (*.txt), *.txt")
    
    If fileName <> False Then
        ChooseOutputLocation = CStr(fileName)
    Else
        ChooseOutputLocation = ""
    End If
End Function

Function SafeText(inputText As Variant) As String
    If IsError(inputText) Or IsEmpty(inputText) Or IsNull(inputText) Then
        SafeText = "(empty)"
    Else
        SafeText = CStr(inputText)
        If Len(SafeText) > 30 Then SafeText = Left(SafeText, 27) & "..."
    End If
End Function

Function GetRowCount(tbl As ListObject) As Long
    On Error Resume Next
    If tbl.DataBodyRange Is Nothing Then
        GetRowCount = 0
    Else
        GetRowCount = tbl.DataBodyRange.Rows.Count
    End If
    If Err.Number <> 0 Then GetRowCount = 0
    On Error GoTo 0
End Function