Attribute VB_Name = "Module3"
Dim targetwb As Workbook
Dim dashboard As Worksheet
Sub autofitcolumns(book As String)
Dim ws As Worksheet
Dim wb As Workbook

Set wb = Workbooks(book)
    For Each ws In wb.Worksheets
        If ws.name <> "Sheet1" Then
            ws.Cells.EntireColumn.AutoFit
            wb.Save
        End If
    Next ws
End Sub
Private Sub set_targetwb()
Dim ws As Worksheet
Dim datafilepath As String

Set ws = ThisWorkbook.Sheets("Helper")
If ws Is Nothing Then 'WAIT is HELPER missing ?!
    MsgBox "Helper Sheet missing.. Can not run step 5"
    Exit Sub
End If
datafilepath = ThisWorkbook.Path & "\" & ws.Range("A2").Value 'should be in same place as created by WriteMFData
If Module1.FileExists(datafilepath) Then
Set targetwb = Workbooks.Open(datafilepath) 'check and open
'MyDebugPrint "DataFile " & datafilepath & vbCr & "TargetWB :" & targetwb.Name
Workbooks.Open ws.Range("A2").Value
Else
MsgBox "Target file not found", vbExclamation, "FILE MISSING"
End If

End Sub
Function stringtodouble(stringnav As String, choice As Integer) As Double
If choice = 1 Then
    If IsNumeric(stringnav) And Not IsDate(Value) Then
        ' Convert the string to a double
        stringtodouble = CDbl(stringnav)
    Else
        ' If not numeric, raise an error
        Err.Raise vbObjectError + 1, , "Value is not a valid number."
    End If
ElseIf choice = 2 Then '2 for date
    If IsDate(stringnav) Then
        'convert to date
        stringtodouble = CDate(stringnav)
    Else
        Err.Raise vbObjectError + 1, , "Value is not a valid date."
    End If
End If
End Function

Private Sub Step5_Dashboard() 'OLD implementation updates NAV of new NAVAll.txt
Dim foundcell As Range
Dim ws As Worksheet
Dim dashexisting As Boolean
Dim isin As String, isinname As String
Dim stringnav As String
Dim NAV, closingbal As Double
Dim lastcol As Long, lastrow As Long
Dim asondate As Date

dashexisting = False
'Open targetwb
set_targetwb
On Error GoTo errohandler

'Check Dashboard existence
On Error Resume Next
For Each ws In targetwb.Sheets
    If ws.name = "Dashboard" Then
    dashexisting = True
    Exit For
    End If
Next ws
On Error GoTo errohandler
If dashexisting Then
    Set dashboard = targetwb.Sheets("Dashboard")
    lastcol = dashboard.Cells(1, ws.Columns.Count).End(xlToLeft).Column + 1
    lastrow = dashboard.Cells(Rows.Count, 1).End(xlUp).row + 1
    'MyDebugPrint lastcol & " : " & lastrow
    dashboard.Cells(1, lastcol).Value = "Next Updated NAV Date"
    dashboard.Cells(1, lastcol + 1).Value = "Closing"
Else
    Set dashboard = targetwb.Sheets.Add(after:=targetwb.Sheets("Log"))
    dashboard.name = "Dashboard"
    targetwb.Save
    dashboard.Range("A1").Value = "ISIN"
    dashboard.Range("B1").Value = "Name"
    dashboard.Range("C1").Value = "As On Date"
    dashboard.Range("D1").Value = "Closing"
    lastrow = 2
    lastcol = 4
End If

'Fill Dashboard
For Each ws In targetwb.Sheets
    If ws.name = "Log" Or ws.name = "Dashboard" Then
    'do nothing
    Else
        isin = ws.Range("B1").Value
        isinname = ws.Range("C1").Value
        stringnav = Module1.GetFundNameByISIN(isin, 4) '4th place is NAV
        NAV = stringtodouble(stringnav, 1) ' 1 for double
        stringnav = Module1.GetFundNameByISIN(isin, 5) '5th place is as on date
        asondate = stringtodouble(stringnav, 2) '2 for date
        'MyDebugPrint "NAV : " & NAV & vbCrLf & "Date : " & Date
        'Add New Closing Date from last
        If dashexisting = True Then
            Set foundcell = dashboard.Cells.Find(What:=isin, after:=dashboard.Cells(1, 1), _
                                  LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
            MyDebugPrint foundcell.Value & " : " & foundcell.Address
            If Not foundcell Is Nothing Then 'Update the existing ISIN
                dashboard.Cells(foundcell.row, lastcol) = asondate
                dashboard.Cells(foundcell.row, lastcol + 1) = NAV
            Else 'New ISIN
                dashboard.Cells(lastrow, 1) = isin
                dashboard.Cells(lastrow, 2) = isinname
                dashboard.Cells(lastrow, lastcol) = asondate
                dashboard.Cells(lastrow, lastcol + 1) = NAV
            End If
        Else 'dashexisting is false i.e. new dashboard
            dashboard.Cells(lastrow, 1) = isin
            dashboard.Cells(lastrow, 2) = isinname
            dashboard.Cells(lastrow, lastcol) = NAV
            dashboard.Cells(lastrow, lastcol - 1) = asondate
        End If
        'Update LastRow
        lastrow = lastrow + 1
    End If

Next ws

' Task Complete Exit
Exit Sub
errohandler:
MsgBox "An error occurred " & Err.Description, vbCritical, "Error " & Err.Number
End Sub

Sub Step5_NewDashboard() 'NEW
Dim foundcell As Range
Dim ws As Worksheet
Dim dashexisting As Boolean
Dim isin As String, isinname As String
Dim stringnav As String
Dim NAV As Double, closingbal As Double
Dim lastcol As Long, lastrow As Long
Dim asondate As Date

dashexisting = False
'Open targetwb
set_targetwb
On Error GoTo errohandler

'Check Dashboard existence
On Error Resume Next
For Each ws In targetwb.Sheets
    If ws.name = "Dashboard" Then
    dashexisting = True
    Exit For
    End If
Next ws
On Error GoTo errohandler
If dashexisting Then
    Set dashboard = targetwb.Sheets("Dashboard")
    lastcol = dashboard.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    lastrow = dashboard.Cells(Rows.Count, 1).End(xlUp).row + 1
    'MyDebugPrint lastcol & " : " & lastrow
'    dashboard.Cells(1, lastcol).Value = "Next Updated NAV Date"
'    dashboard.Cells(1, lastcol + 1).Value = "Closing"
Else
    Set dashboard = targetwb.Sheets.Add(after:=targetwb.Sheets("Log"))
    dashboard.name = "Dashboard"
    targetwb.Save
    dashboard.Range("A1").Value = "ISIN"
    dashboard.Range("B1").Value = "Name"
    dashboard.Range("C1").Value = "As On Date"
    dashboard.Range("D1").Value = "Closing"
    lastrow = 2
    lastcol = 4
End If

'Fill Dashboard
For Each ws In targetwb.Sheets
    If ws.name = "Log" Or ws.name = "Dashboard" Then
    'do nothing
    Else
        isin = ws.Range("B1").Value
        isinname = ws.Range("C1").Value
        closingbal = CDbl(Right(ws.Range("B2").Value, Len(ws.Range("B2").Value) - InStrRev(ws.Range("B2").Value, ":") - 1))
        If closingbal = 0 Then GoTo meh
        Debug.Print closingbal
        stringnav = Module1.GetFundNameByISIN(isin, 4) '4th place is NAV
        NAV = stringtodouble(stringnav, 1) ' 1 for double
        stringnav = Module1.GetFundNameByISIN(isin, 5) '5th place is as on date
        asondate = stringtodouble(stringnav, 2) '2 for date
        'MyDebugPrint "NAV : " & NAV & vbCrLf & "Date : " & Date
        'Add New Closing Date from last
        If dashexisting = True Then
            Set foundcell = dashboard.Cells.Find(What:=isin, after:=dashboard.Cells(1, 1), _
                                  LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
            If Not foundcell Is Nothing Then 'Update the existing ISIN
                dashboard.Cells(foundcell.row, lastcol) = NAV
                dashboard.Cells(foundcell.row, lastcol - 1) = asondate
            End If
        Else 'dashexisting is false i.e. new dashboard
            dashboard.Cells(lastrow, 1) = isin
            dashboard.Cells(lastrow, 2) = isinname
            dashboard.Cells(lastrow, lastcol) = NAV
            dashboard.Cells(lastrow, lastcol - 1) = asondate
        End If
        'Update LastRow
        lastrow = lastrow + 1
    End If
meh:
Next ws
dashboard.Columns.AutoFit
targetwb.Save
' Task Complete Exit
Exit Sub
errohandler:
MsgBox "An error occurred " & Err.Description, vbCritical, "Error " & Err.Number
End Sub

