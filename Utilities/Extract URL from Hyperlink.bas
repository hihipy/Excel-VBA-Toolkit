' ==========================================================================================
' 🔗 Function: URL(cellRef As Range) As String
' 📁 Purpose:
'     Extracts the **true hyperlink URL** from a cell in Excel that contains a hyperlink.
'     Ideal for inventorying links, validating document references, or generating export lists.
'
' ------------------------------------------------------------------------------------------
' ✅ Key Features:
'     - Returns the `Hyperlinks(1).Address` for the specified cell
'     - Gracefully handles:
'         • Empty cells
'         • Cells without hyperlinks
'         • Error states (via `IFERROR`)
'         • Multi-cell input prevention
'     - Compatible with `=URL(A1)` in Excel
'
' ------------------------------------------------------------------------------------------
' 🔍 Example Excel Use Cases:
'     | Cell A1 (Display Text)         | Formula        | Output                    |
'     |-------------------------------|----------------|---------------------------|
'     | Click here (link to example)  | `=URL(A1)`     | https://example.com       |
'     | Plain text (no hyperlink)     | `=URL(A1)`     | ""                        |
'     | =IFERROR(URL(A1), "")         |                | ✅ Safe fallback usage     |
'
' ------------------------------------------------------------------------------------------
' 🧠 When to Use:
'     - Generating lists of external document links
'     - Extracting source URLs from buttons or text labels
'     - Reviewing legacy spreadsheets for outdated web references
'
' ------------------------------------------------------------------------------------------
' ⚠️ Notes:
'     - Returns only the **first** hyperlink if multiple exist
'     - If the cell is not hyperlinked, returns blank string ("")
'     - For clean results, wrap usage in `IFERROR()`
'
' ==========================================================================================
Function URL(cellRef As Range) As String
    ' Extracts the URL from a cell containing a hyperlink
    On Error GoTo ErrorHandler

    If cellRef Is Nothing Then
        URL = "Error: No range provided"
        Exit Function
    End If

    If cellRef.Cells.Count > 1 Then
        URL = "Error: Please select a single cell"
        Exit Function
    End If

    If cellRef.Hyperlinks.Count = 0 Then
        URL = ""  ' No hyperlink found
    Else
        URL = cellRef.Hyperlinks(1).Address
    End If
    Exit Function

ErrorHandler:
    URL = "Error: " & Err.Description
End Function
