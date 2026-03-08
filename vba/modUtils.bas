Attribute VB_Name = "modUtils"
Option Explicit

Public Function GetOrCreateSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        GetOrCreateSheet.Name = sheetName
    End If
End Function

Public Function GetTable(ByVal sheetName As String, ByVal tableName As String) As ListObject
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)
    Set GetTable = ws.ListObjects(tableName)
End Function

Public Sub EnsureTable(ByVal ws As Worksheet, ByVal tableName As String, ByVal headers As Variant)
    Dim lo As ListObject
    Dim i As Long

    On Error Resume Next
    Set lo = ws.ListObjects(tableName)
    On Error GoTo 0

    If lo Is Nothing Then
        For i = LBound(headers) To UBound(headers)
            ws.Cells(1, i + 1).Value = headers(i)
        Next i
        Set lo = ws.ListObjects.Add(SourceType:=xlSrcRange, Source:=ws.Range(ws.Cells(1, 1), ws.Cells(2, UBound(headers) + 1)), XlListObjectHasHeaders:=xlYes)
        lo.Name = tableName
        lo.DataBodyRange.Rows(1).ClearContents
    End If

    ws.Rows(1).Font.Bold = True
    ws.Columns.AutoFit
End Sub

Public Function NextId(ByVal lo As ListObject, ByVal idColumnName As String) As Long
    Dim maxVal As Long
    Dim c As Range
    maxVal = 0

    If lo.DataBodyRange Is Nothing Then
        NextId = 1
        Exit Function
    End If

    For Each c In lo.ListColumns(idColumnName).DataBodyRange
        If IsNumeric(c.Value) Then
            If CLng(c.Value) > maxVal Then maxVal = CLng(c.Value)
        End If
    Next c

    NextId = maxVal + 1
End Function

Public Function NzDbl(ByVal v As Variant) As Double
    If IsNumeric(v) Then
        NzDbl = CDbl(v)
    Else
        NzDbl = 0
    End If
End Function

Public Function FindRowByValue(ByVal lo As ListObject, ByVal columnName As String, ByVal lookupValue As Variant) As ListRow
    Dim r As ListRow
    For Each r In lo.ListRows
        If CStr(r.Range.Cells(1, lo.ListColumns(columnName).Index).Value) = CStr(lookupValue) Then
            Set FindRowByValue = r
            Exit Function
        End If
    Next r
End Function

Public Function DateSafe(ByVal v As Variant) As Date
    If IsDate(v) Then
        DateSafe = CDate(v)
    Else
        DateSafe = Date
    End If
End Function
