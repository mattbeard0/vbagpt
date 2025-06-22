Option Explicit

Private Const OPENAI_ENDPOINT As String = "https://api.openai.com/v1/chat/completions"

Function GetNameValue(nameText As String) As Variant
    Dim nm As Name
    On Error GoTo ErrHandler
    Set nm = ThisWorkbook.Names(nameText)
    GetNameValue = Application.Evaluate(nm.Name)
    Exit Function
ErrHandler:
    GetNameValue = CVErr(xlErrName)  ' name not found
End Function

Public Function CALLGPT(prompt As String) As String
    Dim apiKey As String, model As String
    Dim http   As Object
    Dim payload As String
    Dim respText As String
    Dim json   As Object

    On Error GoTo ErrMissingKey
    apiKey = GetNameValue("OPENAI_API_KEY")
    On Error GoTo ErrMissingModel
    model = GetNameValue("OPENAI_MODEL")
    On Error GoTo ErrRequest
    If Len(Trim(apiKey)) = 0 Then
        CALLGPT = "ERROR: Named range 'OPENAI_API_KEY' is empty; set it to your OpenAI API key."
        Exit Function
    End If

    Set json = JsonConverter.ParseJson("{""model"":'',""messages"":[{""role"":'',""content"":''}]}")

    json("model") = model
    json("messages")(1)("role") = "user"
    json("messages")(1)("content") = prompt

    Set http = CreateObject("MSXML2.XMLHTTP")
    With http
        .Open "POST", OPENAI_ENDPOINT, False
        .setRequestHeader "Content-Type", "application/json"
        .setRequestHeader "Authorization", "Bearer " & apiKey
        .Send JsonConverter.ConvertToJson(json)
        If .Status <> 200 Then
            CALLGPT = "ERROR: HTTP " & .Status & " - " & .StatusText
            Exit Function
        End If
        respText = .responseText
    End With

    Set json = JsonConverter.ParseJson(respText)
    CALLGPT = json("choices")(1)("message")("content")
    Exit Function

ErrMissingKey:
    CALLGPT = "ERROR: Named range 'OpenAI_API_Key' not found; please create it pointing to a cell with your API key or set it in Name Manager."
    Exit Function

ErrMissingModel:
    CALLGPT = "ERROR: Named range 'OPENAI_MODEL' not found; please create it pointing to a cell with your model namem or set it in Name Manager."
    Exit Function

ErrRequest:
    CALLGPT = "ERROR: " & Err.Number & " - " & Err.Description
End Function

Public Function RANGETOMARKDOWN(rng As Range) As String
    Dim row As Range
    Dim cell As Range
    Dim markdown As String
    Dim rowText As String
    
    markdown = ""
    
    For Each row In rng.Rows
        rowText = ""
        For Each cell In row.Cells
            If Len(rowText) > 0 Then
                rowText = rowText & " | "
            End If
            rowText = rowText & cell.Text
        Next cell
        markdown = markdown & "| " & rowText & " |" & vbCrLf
    Next row
    
    RANGETOMARKDOWN = markdown
End Function

Public Function TABLETOMARKDOWN(tableOrRange As Variant) As String
    Dim rng As Range
    Dim headers As Range
    Dim dataRange As Range
    Dim markdown As String
    
    ' Determine if input is a ListObject (table) or Range
    If TypeName(tableOrRange) = "ListObject" Then
        ' It's a table object
        Set rng = tableOrRange.Range
        Set headers = tableOrRange.HeaderRowRange
        Set dataRange = tableOrRange.DataBodyRange
    Else
        ' It's a range - use first row as headers
        Set rng = tableOrRange
        Set headers = rng.Rows(1)
        Set dataRange = rng.Rows(2).Resize(rng.Rows.Count - 1)
    End If
    
    ' Create header row
    markdown = RANGETOMARKDOWN(headers)
    
    ' Add separator row
    Dim headerCell As Range
    Dim separatorRow As String
    separatorRow = "|"
    For Each headerCell In headers.Cells
        separatorRow = separatorRow & " --- |"
    Next headerCell
    markdown = markdown & separatorRow & vbCrLf
    
    ' Add data rows
    If Not dataRange Is Nothing Then
        markdown = markdown & RANGETOMARKDOWN(dataRange)
    End If
    
    TABLETOMARKDOWN = markdown
End Function



