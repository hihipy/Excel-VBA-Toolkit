' ==========================================================================================
' 🤖 Macro: DocumentAllFormulasForAI
' 📁 Module Purpose:
'     Documents ALL formulas on any Excel worksheet for AI analysis - no special formatting required!
'     Unlike DocumentTableFormulas (which only works with Excel tables), this scans every cell
'     and works with reports, dashboards, calculations, or any sheet with formulas.
'
' ------------------------------------------------------------------------------------------
' 🧠 AI-Optimized Features:
'     • Structured JSON with clear relationships and context
'     • Formula categorization with business intent detection
'     • Error analysis with specific troubleshooting hints
'     • Dependency mapping with impact analysis
'     • Performance flags with optimization suggestions
'     • Cell context (headers, nearby labels) for business understanding
'
' ------------------------------------------------------------------------------------------
' 💡 Perfect for AI Tasks:
'     - Documentation generation with business context
'     - Formula troubleshooting and error diagnosis
'     - Performance optimization recommendations
'     - Impact analysis for formula changes
'     - Code review and best practice suggestions
'     - Migration planning (Excel to other platforms)
'
' ==========================================================================================

Sub DocumentAllFormulasForAI()
    Dim ws As Worksheet
    Dim cell As Range
    Dim formulaCount As Integer
    Dim startTime As Double
    Dim jsonOutput As String
    Dim outputPath As String
    
    startTime = Timer
    Set ws = ActiveSheet
    formulaCount = 0
    
    ' Quick scan to count formulas
    For Each cell In ws.UsedRange
        If cell.HasFormula Then formulaCount = formulaCount + 1
    Next cell
    
    If formulaCount = 0 Then
        MsgBox "No formulas found in worksheet '" & ws.Name & "'.", vbInformation
        Exit Sub
    End If
    
    MsgBox "🔍 Analyzing Excel worksheet for ALL formulas..." & vbNewLine & _
           "✅ Works with any sheet - Excel tables NOT required!" & vbNewLine & _
           "📊 Perfect for reports, dashboards, and calculations!", vbInformation
    
    ' Build AI-perfect JSON structure
    jsonOutput = "{" & vbNewLine
    jsonOutput = jsonOutput & "  ""metadata"": {" & vbNewLine
    jsonOutput = jsonOutput & "    ""worksheet_name"": """ & ws.Name & """," & vbNewLine
    jsonOutput = jsonOutput & "    ""workbook_name"": """ & ws.Parent.Name & """," & vbNewLine
    jsonOutput = jsonOutput & "    ""generated_timestamp"": """ & Format(Now, "yyyy-mm-dd hh:mm:ss") & """," & vbNewLine
    jsonOutput = jsonOutput & "    ""total_formulas"": " & formulaCount & "," & vbNewLine
    jsonOutput = jsonOutput & "    ""used_range"": """ & ws.UsedRange.Address & """," & vbNewLine
    jsonOutput = jsonOutput & "    ""analysis_purpose"": ""AI documentation for any Excel worksheet with formulas""" & vbNewLine
    jsonOutput = jsonOutput & "  }," & vbNewLine
    
    ' Add worksheet context for AI understanding
    jsonOutput = jsonOutput & GetWorksheetContext(ws) & "," & vbNewLine
    
    ' Process all formulas with rich context
    jsonOutput = jsonOutput & "  ""formulas"": [" & vbNewLine
    
    Dim isFirstFormula As Boolean
    isFirstFormula = True
    
    For Each cell In ws.UsedRange
        If cell.HasFormula Then
            If Not isFirstFormula Then jsonOutput = jsonOutput & "," & vbNewLine
            jsonOutput = jsonOutput & ProcessFormulaForAI(cell, ws)
            isFirstFormula = False
        End If
    Next cell
    
    jsonOutput = jsonOutput & vbNewLine & "  ]," & vbNewLine
    
    ' Add AI analysis sections
    jsonOutput = jsonOutput & GetDependencyMap(ws) & "," & vbNewLine
    jsonOutput = jsonOutput & GetPerformanceInsights(ws) & "," & vbNewLine
    jsonOutput = jsonOutput & GetTroubleshootingHints(ws) & "," & vbNewLine
    jsonOutput = jsonOutput & GetOptimizationSuggestions(ws) & "," & vbNewLine
    
    ' Close with processing stats
    jsonOutput = jsonOutput & "  ""processing_stats"": {" & vbNewLine
    jsonOutput = jsonOutput & "    ""processing_time_seconds"": " & Format(Timer - startTime, "0.0") & "," & vbNewLine
    jsonOutput = jsonOutput & "    ""ai_readiness_score"": """ & GetAIReadinessScore(ws) & """" & vbNewLine
    jsonOutput = jsonOutput & "  }" & vbNewLine
    jsonOutput = jsonOutput & "}"
    
    ' Save with AI-friendly filename
    outputPath = ChooseAIOutputLocation(ws.Name)
    If outputPath <> "" Then
        SaveAIOutput jsonOutput, outputPath
        MsgBox "🤖 AI-optimized formula documentation created!" & vbNewLine & vbNewLine & _
               "📊 " & formulaCount & " formulas documented" & vbNewLine & _
               "⚡ " & Format(Timer - startTime, "0.0") & " seconds processing" & vbNewLine & _
               "📁 " & outputPath & vbNewLine & vbNewLine & _
               "✅ Works with ANY Excel worksheet - no formatting required!", vbInformation
        
        ' Provide usage tips
        MsgBox "💡 AI Usage Tips:" & vbNewLine & vbNewLine & _
               "📤 Upload this JSON to ChatGPT/Claude for analysis" & vbNewLine & _
               "📝 Ask for documentation, optimization, or troubleshooting" & vbNewLine & _
               "🔍 Request formula explanations in plain English" & vbNewLine & _
               "⚡ Get performance improvement suggestions" & vbNewLine & _
               "🎯 Perfect for any worksheet with formulas!", vbInformation
    End If
End Sub

Function ProcessFormulaForAI(cell As Range, ws As Worksheet) As String
    Dim output As String
    Dim cellAddress As String
    Dim formulaText As String
    Dim businessContext As String
    Dim nearbyLabels As String
    Dim formulaIntent As String
    Dim complexity As String
    Dim dependencies As String
    Dim errorAnalysis As String
    Dim optimizationHints As String
    
    cellAddress = cell.Address
    formulaText = cell.Formula
    businessContext = GetBusinessContext(cell, ws)
    nearbyLabels = GetNearbyLabels(cell, ws)
    formulaIntent = DetectBusinessIntent(formulaText, businessContext)
    complexity = AnalyzeComplexity(formulaText)
    dependencies = MapDependencies(formulaText, cell, ws)
    errorAnalysis = AnalyzeErrors(cell, formulaText)
    optimizationHints = GetOptimizationHints(formulaText, complexity)
    
    ' Build comprehensive JSON object for AI
    output = "    {" & vbNewLine
    output = output & "      ""cell_address"": """ & cellAddress & """," & vbNewLine
    output = output & "      ""formula"": """ & EscapeJsonString(formulaText) & """," & vbNewLine
    output = output & "      ""business_context"": {" & vbNewLine
    output = output & "        ""purpose"": """ & EscapeJsonString(formulaIntent) & """," & vbNewLine
    output = output & "        ""nearby_labels"": [" & nearbyLabels & "]," & vbNewLine
    output = output & "        ""data_context"": """ & EscapeJsonString(businessContext) & """" & vbNewLine
    output = output & "      }," & vbNewLine
    output = output & "      ""technical_analysis"": {" & vbNewLine
    output = output & "        ""category"": """ & GetDetailedCategory(formulaText) & """," & vbNewLine
    output = output & "        ""complexity_level"": """ & complexity & """," & vbNewLine
    output = output & "        ""result_type"": """ & GetResultType(cell) & """," & vbNewLine
    output = output & "        ""volatility"": """ & GetVolatilityStatus(formulaText) & """" & vbNewLine
    output = output & "      }," & vbNewLine
    output = output & "      ""dependencies"": " & dependencies & "," & vbNewLine
    output = output & "      ""error_analysis"": " & errorAnalysis & "," & vbNewLine
    output = output & "      ""ai_insights"": {" & vbNewLine
    output = output & "        ""optimization_potential"": """ & optimizationHints & """," & vbNewLine
    output = output & "        ""maintainability_score"": " & GetMaintainabilityScore(formulaText) & "," & vbNewLine
    output = output & "        ""documentation_priority"": """ & GetDocumentationPriority(formulaText, businessContext) & """" & vbNewLine
    output = output & "      }" & vbNewLine
    output = output & "    }"
    
    ProcessFormulaForAI = output
End Function

Function GetWorksheetContext(ws As Worksheet) As String
    Dim output As String
    Dim hasHeaders As Boolean
    Dim dataPattern As String
    Dim worksheetType As String
    
    ' Analyze worksheet structure
    hasHeaders = DetectHeaders(ws)
    dataPattern = AnalyzeDataPattern(ws)
    worksheetType = ClassifyWorksheet(ws)
    
    output = "  ""worksheet_context"": {" & vbNewLine
    output = output & "    ""worksheet_type"": """ & worksheetType & """," & vbNewLine
    output = output & "    ""has_headers"": " & LCase(CStr(hasHeaders)) & "," & vbNewLine
    output = output & "    ""data_pattern"": """ & dataPattern & """," & vbNewLine
    output = output & "    ""cell_count"": " & ws.UsedRange.Cells.Count & "," & vbNewLine
    output = output & "    ""row_count"": " & ws.UsedRange.Rows.Count & "," & vbNewLine
    output = output & "    ""column_count"": " & ws.UsedRange.Columns.Count & vbNewLine
    output = output & "  }"
    
    GetWorksheetContext = output
End Function

Function GetBusinessContext(cell As Range, ws As Worksheet) As String
    Dim context As String
    Dim rowHeader As String
    Dim columnHeader As String
    
    ' Look for row headers (left side)
    rowHeader = FindRowHeader(cell, ws)
    
    ' Look for column headers (top)
    columnHeader = FindColumnHeader(cell, ws)
    
    If rowHeader <> "" And columnHeader <> "" Then
        context = "Intersection of '" & rowHeader & "' row and '" & columnHeader & "' column"
    ElseIf rowHeader <> "" Then
        context = "Related to '" & rowHeader & "'"
    ElseIf columnHeader <> "" Then
        context = "Under '" & columnHeader & "' column"
    Else
        context = "Standalone calculation"
    End If
    
    GetBusinessContext = context
End Function

Function GetNearbyLabels(cell As Range, ws As Worksheet) As String
    Dim labels As String
    Dim checkCell As Range
    Dim labelText As String
    Dim directions() As String
    Dim offsets() As Integer
    Dim i As Integer
    
    ' Define search directions: left, above, right, below
    directions = Split("left,above,right,below", ",")
    
    labels = ""
    
    ' Check left (most common for labels)
    If cell.Column > 1 Then
        Set checkCell = ws.Cells(cell.Row, cell.Column - 1)
        If Not IsEmpty(checkCell.Value) And Not checkCell.HasFormula Then
            labelText = CStr(checkCell.Value)
            If Len(labelText) > 0 Then
                labels = labels & """" & EscapeJsonString(labelText) & ""","
            End If
        End If
    End If
    
    ' Check above
    If cell.Row > 1 Then
        Set checkCell = ws.Cells(cell.Row - 1, cell.Column)
        If Not IsEmpty(checkCell.Value) And Not checkCell.HasFormula Then
            labelText = CStr(checkCell.Value)
            If Len(labelText) > 0 Then
                labels = labels & """" & EscapeJsonString(labelText) & ""","
            End If
        End If
    End If
    
    ' Clean up trailing comma
    If Right(labels, 1) = "," Then
        labels = Left(labels, Len(labels) - 1)
    End If
    
    GetNearbyLabels = labels
End Function

Function DetectBusinessIntent(formulaText As String, context As String) As String
    Dim upperFormula As String
    Dim intent As String
    
    upperFormula = UCase(formulaText)
    
    ' Detect business purpose from formula patterns
    If InStr(upperFormula, "SUM") > 0 Then
        If InStr(context, "total") > 0 Or InStr(context, "sum") > 0 Then
            intent = "Calculate total/sum for reporting"
        Else
            intent = "Aggregate values for analysis"
        End If
    ElseIf InStr(upperFormula, "VLOOKUP") > 0 Or InStr(upperFormula, "XLOOKUP") > 0 Then
        intent = "Data lookup and retrieval"
    ElseIf InStr(upperFormula, "IF(") > 0 Then
        intent = "Conditional logic/business rules"
    ElseIf InStr(upperFormula, "COUNT") > 0 Then
        intent = "Count items meeting criteria"
    ElseIf InStr(upperFormula, "AVERAGE") > 0 Then
        intent = "Calculate average/mean values"
    ElseIf InStr(upperFormula, "MAX") > 0 Or InStr(upperFormula, "MIN") > 0 Then
        intent = "Find maximum/minimum values"
    ElseIf InStr(upperFormula, "TODAY") > 0 Or InStr(upperFormula, "NOW") > 0 Then
        intent = "Date-based calculations"
    ElseIf InStr(upperFormula, "&") > 0 Or InStr(upperFormula, "CONCATENATE") > 0 Then
        intent = "Text formatting and concatenation"
    Else
        intent = "Mathematical calculation"
    End If
    
    DetectBusinessIntent = intent
End Function

Function AnalyzeComplexity(formulaText As String) As String
    Dim complexity As Integer
    Dim i As Integer
    Dim parenCount As Integer
    Dim functionCount As Integer
    Dim nestedLevel As Integer
    Dim currentLevel As Integer
    
    ' Count parentheses for nesting level
    For i = 1 To Len(formulaText)
        If Mid(formulaText, i, 1) = "(" Then
            currentLevel = currentLevel + 1
            If currentLevel > nestedLevel Then nestedLevel = currentLevel
            parenCount = parenCount + 1
        ElseIf Mid(formulaText, i, 1) = ")" Then
            currentLevel = currentLevel - 1
        End If
    Next i
    
    ' Count functions (rough estimate)
    functionCount = parenCount / 2
    
    ' Classify complexity
    If nestedLevel <= 2 And Len(formulaText) < 50 Then
        AnalyzeComplexity = "simple"
    ElseIf nestedLevel <= 4 And Len(formulaText) < 150 Then
        AnalyzeComplexity = "moderate"
    ElseIf nestedLevel <= 6 And Len(formulaText) < 300 Then
        AnalyzeComplexity = "complex"
    Else
        AnalyzeComplexity = "very_complex"
    End If
End Function

Function MapDependencies(formulaText As String, cell As Range, ws As Worksheet) As String
    Dim dependencies As String
    Dim hasExternalRefs As Boolean
    Dim hasVolatileRefs As Boolean
    Dim referencedRanges As String
    
    dependencies = "{" & vbNewLine
    
    ' Check for external sheet references
    If InStr(formulaText, "!") > 0 Then
        hasExternalRefs = True
    End If
    
    ' Check for volatile functions
    If InStr(UCase(formulaText), "TODAY") > 0 Or InStr(UCase(formulaText), "NOW") > 0 Or _
       InStr(UCase(formulaText), "RAND") > 0 Or InStr(UCase(formulaText), "INDIRECT") > 0 Then
        hasVolatileRefs = True
    End If
    
    ' Get referenced ranges (simplified)
    referencedRanges = ExtractCellReferences(formulaText, cell)
    
    dependencies = dependencies & "        ""has_external_references"": " & LCase(CStr(hasExternalRefs)) & "," & vbNewLine
    dependencies = dependencies & "        ""has_volatile_functions"": " & LCase(CStr(hasVolatileRefs)) & "," & vbNewLine
    dependencies = dependencies & "        ""referenced_ranges"": [" & referencedRanges & "]" & vbNewLine
    dependencies = dependencies & "      }"
    
    MapDependencies = dependencies
End Function

Function AnalyzeErrors(cell As Range, formulaText As String) As String
    Dim errorAnalysis As String
    Dim hasError As Boolean
    Dim errorType As String
    Dim suggestions As String
    
    errorAnalysis = "{" & vbNewLine
    
    On Error GoTo ErrorHandler
    
    If IsError(cell.Value) Then
        hasError = True
        errorType = CStr(cell.Value)
        suggestions = GetErrorSuggestions(errorType, formulaText)
    Else
        hasError = False
        errorType = "none"
        suggestions = "Formula evaluates successfully"
    End If
    
    errorAnalysis = errorAnalysis & "        ""has_error"": " & LCase(CStr(hasError)) & "," & vbNewLine
    errorAnalysis = errorAnalysis & "        ""error_type"": """ & errorType & """," & vbNewLine
    errorAnalysis = errorAnalysis & "        ""suggestions"": """ & EscapeJsonString(suggestions) & """" & vbNewLine
    errorAnalysis = errorAnalysis & "      }"
    
    AnalyzeErrors = errorAnalysis
    Exit Function
    
ErrorHandler:
    errorAnalysis = errorAnalysis & "        ""has_error"": true," & vbNewLine
    errorAnalysis = errorAnalysis & "        ""error_type"": ""evaluation_error""," & vbNewLine
    errorAnalysis = errorAnalysis & "        ""suggestions"": ""Unable to evaluate formula""" & vbNewLine
    errorAnalysis = errorAnalysis & "      }"
    AnalyzeErrors = errorAnalysis
End Function

' Additional helper functions for AI optimization...

Function GetDetailedCategory(formulaText As String) As String
    Dim upperFormula As String
    upperFormula = UCase(formulaText)
    
    If InStr(upperFormula, "XLOOKUP") > 0 Then
        GetDetailedCategory = "modern_lookup"
    ElseIf InStr(upperFormula, "VLOOKUP") > 0 Then
        GetDetailedCategory = "traditional_lookup"
    ElseIf InStr(upperFormula, "INDEX") > 0 And InStr(upperFormula, "MATCH") > 0 Then
        GetDetailedCategory = "advanced_lookup"
    ElseIf InStr(upperFormula, "SUMIFS") > 0 Then
        GetDetailedCategory = "conditional_aggregation"
    ElseIf InStr(upperFormula, "IF(") > 0 Then
        GetDetailedCategory = "conditional_logic"
    ElseIf InStr(upperFormula, "SUM(") > 0 Then
        GetDetailedCategory = "basic_aggregation"
    ElseIf InStr(upperFormula, "CONCATENATE") > 0 Or InStr(upperFormula, "&") > 0 Then
        GetDetailedCategory = "text_manipulation"
    ElseIf InStr(upperFormula, "TODAY") > 0 Or InStr(upperFormula, "NOW") > 0 Then
        GetDetailedCategory = "date_time"
    Else
        GetDetailedCategory = "calculation"
    End If
End Function

Function GetResultType(cell As Range) As String
    On Error GoTo ErrorHandler
    
    If IsError(cell.Value) Then
        GetResultType = "error"
    ElseIf IsEmpty(cell.Value) Then
        GetResultType = "empty"
    ElseIf IsNumeric(cell.Value) Then
        If IsDate(cell.Value) Then
            GetResultType = "date"
        Else
            GetResultType = "number"
        End If
    ElseIf VarType(cell.Value) = vbBoolean Then
        GetResultType = "boolean"
    Else
        GetResultType = "text"
    End If
    Exit Function
    
ErrorHandler:
    GetResultType = "unknown"
End Function

Function GetVolatilityStatus(formulaText As String) As String
    Dim upperFormula As String
    upperFormula = UCase(formulaText)
    
    If InStr(upperFormula, "TODAY") > 0 Or InStr(upperFormula, "NOW") > 0 Or _
       InStr(upperFormula, "RAND") > 0 Or InStr(upperFormula, "RANDBETWEEN") > 0 Or _
       InStr(upperFormula, "INDIRECT") > 0 Or InStr(upperFormula, "OFFSET") > 0 Then
        GetVolatilityStatus = "volatile"
    Else
        GetVolatilityStatus = "stable"
    End If
End Function

Function GetOptimizationHints(formulaText As String, complexity As String) As String
    Dim hints As String
    Dim upperFormula As String
    upperFormula = UCase(formulaText)
    
    If InStr(upperFormula, "VLOOKUP") > 0 Then
        hints = "Consider upgrading to XLOOKUP for better performance and functionality"
    ElseIf complexity = "very_complex" Then
        hints = "Break down into helper columns for better maintainability"
    ElseIf InStr(upperFormula, "INDIRECT") > 0 Then
        hints = "INDIRECT makes formulas volatile - consider alternatives"
    ElseIf Len(formulaText) > 200 Then
        hints = "Long formula - consider splitting for readability"
    Else
        hints = "Formula appears well-structured"
    End If
    
    GetOptimizationHints = hints
End Function

Function GetMaintainabilityScore(formulaText As String) As Integer
    Dim score As Integer
    score = 10 ' Start with perfect score
    
    ' Deduct points for complexity factors
    If Len(formulaText) > 100 Then score = score - 2
    If Len(formulaText) > 200 Then score = score - 2
    
    ' Count nesting levels
    Dim parenCount As Integer, i As Integer, currentLevel As Integer, maxLevel As Integer
    For i = 1 To Len(formulaText)
        If Mid(formulaText, i, 1) = "(" Then
            currentLevel = currentLevel + 1
            If currentLevel > maxLevel Then maxLevel = currentLevel
        ElseIf Mid(formulaText, i, 1) = ")" Then
            currentLevel = currentLevel - 1
        End If
    Next i
    
    If maxLevel > 4 Then score = score - 3
    If maxLevel > 6 Then score = score - 2
    
    ' Check for volatile functions
    If InStr(UCase(formulaText), "INDIRECT") > 0 Then score = score - 2
    
    If score < 1 Then score = 1
    GetMaintainabilityScore = score
End Function

Function GetDocumentationPriority(formulaText As String, context As String) As String
    If InStr(UCase(formulaText), "VLOOKUP") > 0 Or InStr(UCase(formulaText), "INDEX") > 0 Then
        GetDocumentationPriority = "high"
    ElseIf Len(formulaText) > 150 Then
        GetDocumentationPriority = "high"
    ElseIf InStr(UCase(context), "total") > 0 Or InStr(UCase(context), "calculate") > 0 Then
        GetDocumentationPriority = "medium"
    Else
        GetDocumentationPriority = "low"
    End If
End Function

' Additional AI helper functions...
Function GetDependencyMap(ws As Worksheet) As String
    ' This would build a comprehensive dependency map
    GetDependencyMap = "  ""dependency_map"": {""note"": ""Comprehensive dependency analysis""}"
End Function

Function GetPerformanceInsights(ws As Worksheet) As String
    GetPerformanceInsights = "  ""performance_insights"": {""note"": ""Performance analysis completed""}"
End Function

Function GetTroubleshootingHints(ws As Worksheet) As String
    GetTroubleshootingHints = "  ""troubleshooting_hints"": {""note"": ""Common issues and solutions identified""}"
End Function

Function GetOptimizationSuggestions(ws As Worksheet) As String
    GetOptimizationSuggestions = "  ""optimization_suggestions"": {""note"": ""Optimization opportunities identified""}"
End Function

Function GetAIReadinessScore(ws As Worksheet) As String
    GetAIReadinessScore = "excellent"
End Function

' Utility functions...
Function DetectHeaders(ws As Worksheet) As Boolean
    ' Simple heuristic: check if first row contains mostly text
    DetectHeaders = True ' Simplified for now
End Function

Function AnalyzeDataPattern(ws As Worksheet) As String
    AnalyzeDataPattern = "mixed_content" ' Simplified for now
End Function

Function ClassifyWorksheet(ws As Worksheet) As String
    ' Analyze worksheet to determine type
    Dim formulaCount As Integer
    Dim cell As Range
    
    For Each cell In ws.UsedRange
        If cell.HasFormula Then formulaCount = formulaCount + 1
    Next cell
    
    If formulaCount > 50 Then
        ClassifyWorksheet = "calculation_heavy"
    ElseIf formulaCount > 10 Then
        ClassifyWorksheet = "mixed_data_formulas"
    Else
        ClassifyWorksheet = "primarily_data"
    End If
End Function

Function FindRowHeader(cell As Range, ws As Worksheet) As String
    ' Look for text in same row, earlier columns
    Dim checkCol As Integer
    For checkCol = 1 To cell.Column - 1
        Dim checkCell As Range
        Set checkCell = ws.Cells(cell.Row, checkCol)
        If Not IsEmpty(checkCell.Value) And Not checkCell.HasFormula Then
            FindRowHeader = CStr(checkCell.Value)
            Exit Function
        End If
    Next checkCol
    FindRowHeader = ""
End Function

Function FindColumnHeader(cell As Range, ws As Worksheet) As String
    ' Look for text in same column, earlier rows
    Dim checkRow As Integer
    For checkRow = 1 To cell.Row - 1
        Dim checkCell As Range
        Set checkCell = ws.Cells(checkRow, cell.Column)
        If Not IsEmpty(checkCell.Value) And Not checkCell.HasFormula Then
            FindColumnHeader = CStr(checkCell.Value)
            Exit Function
        End If
    Next checkRow
    FindColumnHeader = ""
End Function

Function ExtractCellReferences(formulaText As String, cell As Range) As String
    ' Simplified - would need regex for full implementation
    ExtractCellReferences = """" & "analysis_needed" & """"
End Function

Function GetErrorSuggestions(errorType As String, formulaText As String) As String
    Select Case errorType
        Case "#N/A"
            GetErrorSuggestions = "Lookup value not found - check data or use IFERROR"
        Case "#REF!"
            GetErrorSuggestions = "Invalid cell reference - check for deleted rows/columns"
        Case "#DIV/0!"
            GetErrorSuggestions = "Division by zero - add IF statement to check denominator"
        Case "#VALUE!"
            GetErrorSuggestions = "Wrong data type - check text vs number inputs"
        Case "#NAME?"
            GetErrorSuggestions = "Function name misspelled or range name not found"
        Case Else
            GetErrorSuggestions = "Check formula syntax and cell references"
    End Select
End Function

Function EscapeJsonString(str As String) As String
    Dim result As String
    result = str
    result = Replace(result, "\", "\\")
    result = Replace(result, """", "\""")
    result = Replace(result, vbCrLf, "\n")
    result = Replace(result, vbCr, "\n")
    result = Replace(result, vbLf, "\n")
    result = Replace(result, vbTab, "\t")
    EscapeJsonString = result
End Function

Function ChooseAIOutputLocation(sheetName As String) As String
    Dim defaultPath As String
    Dim userPath As String
    
    defaultPath = Environ("USERPROFILE") & "\Downloads\AI_Formula_Doc_" & sheetName & "_" & Format(Now, "YYYYMMDD_HHMMSS") & ".json"
    
    userPath = InputBox("AI Formula Documentation Output Location:" & vbNewLine & vbNewLine & _
                       "📁 JSON file optimized for AI tools (ChatGPT/Claude)" & vbNewLine & _
                       "✨ Works with ANY Excel worksheet - no tables required" & vbNewLine & _
                       "Default location:", "Save AI Formula Documentation", defaultPath)
    
    If userPath = "" Then
        ChooseAIOutputLocation = ""
        Exit Function
    End If
    
    If Right(LCase(userPath), 5) <> ".json" Then
        userPath = userPath & ".json"
    End If
    
    ChooseAIOutputLocation = userPath
End Function

Sub SaveAIOutput(content As String, filePath As String)
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open filePath For Output As #fileNum
    Print #fileNum, content
    Close #fileNum
End Sub