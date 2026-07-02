Function SheetExists(sheetName As String, wb As Workbook) As Boolean
    On Error Resume Next
    SheetExists = Not wb.Sheets(sheetName) Is Nothing
    On Error GoTo 0
End Function