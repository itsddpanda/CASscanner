Attribute VB_Name = "Module2"
Sub Restart()
Dim ws As Worksheet
Dim wb As Workbook
Unload frmSplash
Application.ScreenUpdating = True
DeleteAllQueriesAndConnections
deleteallsheets
If ThisWorkbook.Sheets.Count = 1 And ThisWorkbook.Sheets(1).name = "Helper" Then
Set ws = ThisWorkbook.Sheets.Add
ws.name = "Restart"
Application.DisplayAlerts = False
ThisWorkbook.Sheets("Helper").Delete
Application.DisplayAlerts = True
Exit Sub
End If
End Sub
Private Sub deleteallsheets()

Dim wb As Workbook
Dim ws As Worksheet
Application.DisplayAlerts = False
Set wb = ThisWorkbook
For Each ws In wb.Sheets
        If ws.name = "Helper" Or wb.Sheets.Count = 1 Then
        'do nothing
        Else
            ws.Delete
        End If
    Next ws
Application.DisplayAlerts = True
End Sub

Private Sub DeleteAllQueriesAndConnections()
    Dim conn As WorkbookConnection
    Dim query As QueryTable
    Dim ws As Worksheet

    ' Delete all Workbook Connections
    For Each conn In ThisWorkbook.Connections
        On Error Resume Next ' In case there are issues with some connections
        conn.Delete
        On Error GoTo 0 ' Reset error handling
    Next conn

    ' Delete all Power Query Queries (if any)
    On Error Resume Next ' Skip if the Queries collection does not exist
    For Each pq In ThisWorkbook.Queries
        pq.Delete
    Next pq
    On Error GoTo 0 ' Reset error handling

    ' Clear any QueryTables in the worksheets (Power Query tables)
    For Each ws In ThisWorkbook.Sheets
        On Error Resume Next ' In case there are no QueryTables on the sheet
        For Each query In ws.QueryTables
            query.Delete
        Next query
        On Error GoTo 0 ' Reset error handling
    Next ws

    'MsgBox "All data queries and connections have been deleted.", vbInformation
End Sub

Private Sub RemoveAllControls()
Dim shp As Shape
On Error GoTo lol
For Each shp In ThisWorkbook.Sheets("Helper").Shapes
shp.Delete
Next shp
lol:
End Sub

Function createlbl(sheetname As String)

End Function
