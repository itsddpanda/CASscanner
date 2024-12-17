Attribute VB_Name = "Module3"
Sub autofitcolumns(book As String)
Dim ws As Worksheet
Dim wb As Workbook

Set wb = Workbooks(book)
    For Each ws In wb.Worksheets
        If ws.Name <> "Sheet1" Then
            ws.Cells.EntireColumn.AutoFit
        End If
    Next ws
End Sub
