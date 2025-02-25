Attribute VB_Name = "Module1"
Dim Foliocont As Boolean, ISFound As Boolean, CASFound As Boolean, WBAdd As Boolean
Dim currentFolio, twb As String
Dim isin As String
Dim closingbalanace As Collection, folios As Collection
Dim OUB As Double
Dim transactionStartRow As Long, transactionEndRow As Long


Sub Step1_SelectPDFFile()
' Version 2.0.1 - Step 1: Ask user to select a PDF file and store file name
    Dim goforlaunch As Integer
    On Error GoTo ErrorHandler ' Enable error handling with custom message
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim pdfFilePath As Variant
    Dim FetchFN As Boolean
goforlaunch = 0
FetchFN = False
    ' Initialize workbook
10  Set wb = ThisWorkbook

    If MsgBox(" Starting the process, Hope you have" & vbCrLf & vbCrLf & "1. Unlocked (password removed) PDF File" & vbCrLf & "2. NAVAll.txt file saved from AMFI website" & vbCrLf & vbCrLf & vbCrLf & "If you are not ready hit cancel now to abort" & vbCrLf & vbCrLf & vbCrLf & "Click OK To Proceed", vbOKCancel) = vbCancel Then Exit Sub
    ' Check if "Helper" sheet exists, if not create it
20  On Error Resume Next
30  Set ws = wb.Sheets("Helper")
40  On Error GoTo ErrorHandler
50  If ws Is Nothing Then
60      Set ws = wb.Sheets.Add(after:=wb.Sheets(wb.Sheets.Count))
70      ws.name = "Helper" ' Create and name the Helper sheet
        FetchFN = True
80  End If

    ' Prompt user to select PDF file
    pdfFilePath = False
    
90  If FetchFN Then
    pdfFilePath = Application.GetOpenFilename("PDF Files (*.pdf), *.pdf", , "Select PDF File")
    'On Error Resume Next
100 If pdfFilePath = False Then
110     MsgBox "No file selected. Exiting...", vbExclamation
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
120     Exit Sub
130 End If

    ' Store the selected PDF file path in the Helper sheet (cell A1)
140 ws.Cells(1, 1).Value = pdfFilePath ' Store PDF path for future reference
180 ws.Cells(1, 2).Value = "Source PDF File"
190 MsgBox "PDF file path stored successfully! Name the outputfile next.", vbInformation
    pdfFilePath = Application.GetSaveAsFilename( _
                  InitialFileName:="MyWorkbook.xlsx", _
                  FileFilter:="Excel Files (*.xlsx), *.xlsx", _
                  Title:="Save Workbook As")
    ws.Cells(2, 1).Value = Mid(pdfFilePath, InStrRev(pdfFilePath, "\") + 1)
    ws.Cells(2, 2).Select
    ws.Cells(2, 2).Value = "Target Data File"
    
    If ws.Cells(2, 1).Value = False Then
    Selection.Style = "BAD"
    Else
    Selection.Style = "Good"
    goforlaunch = goforlaunch + 1
    End If

    MsgBox pdfFilePath & " - Output File Name Stored. Locate NAVAll.txt from AMFI website", vbInformation
    pdfFilePath = Application.GetOpenFilename("Text Files (*.txt), *.txt", , "Select NAVAll.txt File")
    ws.Cells(3, 1).Value = pdfFilePath
    ws.Cells(3, 2).Value = "NAVAll.txt"
    ws.Cells(3, 2).Select
    If ws.Cells(3, 1).Value = False Then
    Selection.Style = "BAD"
    Else
    Selection.Style = "Good"
    goforlaunch = goforlaunch + 1
    End If
    ws.Cells(1, 2).Select
    Selection.Style = "Good"
    ws.Cells.EntireColumn.AutoFit
150
If goforlaunch = 2 Then
    If MsgBox("You are go for LAUNCH...All 3 conditions met" & vbCrLf & "Do want to LAUNCH IT?", vbYesNo, "Conditions Met") = vbYes Then
        frmSplash.Show vbModeless
        frmSplash.UpdateCap "LAUNCHING - STEP 2/4 - READ and EXTRACT PDF SCHEMA - KEEP FINGERS CROSSED"
        Step2_ExtractTableIDs
        frmSplash.UpdateCap "LAUNCHING - STEP 3/4 - READ PDF as per the SCHEMA"
        Step3_ReadPortfolioData
        frmSplash.UpdateCap "LAUNCHING - STEP 4/4 (as of now) - FIND Valid Data and Create the OUTPUT file"
        Step4_readsheets
    End If
End If
Else
MsgBox "Helper Sheet exist thus not asking for new file input" & vbCrLf & "Delete Helper sheet to start Afresh", vbExclamation, "Run the sequence"

End If

160 Exit Sub

ErrorHandler:
    ' Display error with line number
170 MsgBox "An error occurred on line " & Erl & ": " & Err.Description, vbCritical, "Error " & Err.Number
End Sub

Sub Step2_ExtractTableIDs()
    ' Version 2.0.3 - Step 2: Extract table IDs from PDF, execute after Step 1
    On Error GoTo ErrorHandler ' Enable error handling
    
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim pdfFilePath As String
    Dim queryName As String
    Dim pqMCode As String
    Dim sheetExists As Boolean

    ' Initialize workbook and worksheet
10  Set wb = ThisWorkbook
On Error Resume Next
20  Set ws = wb.Sheets("Helper") ' Helper sheet to store file path
On Error GoTo ErrorHandler
If ws Is Nothing Then
MsgBox "Helper Sheet missing, Did you run Step 1?", vbExclamation, "Run the sequence"
Exit Sub
End If
    ' Retrieve the file path from the Helper sheet
30  pdfFilePath = ws.Range("A1").Value
40  If pdfFilePath = "" Then
50      MsgBox "No PDF file path found. Exiting...", vbExclamation
60      Exit Sub
70  End If
frmSplash.Show vbModeless
frmSplash.UpdateCap ("LAUNCHING - STEP 2/4 - READ and EXTRACT PDF SCHEMA - KEEP FINGERS CROSSED")
Application.ScreenUpdating = False
frmSplash.UpdateProgressBar 0.1
Application.Wait Now + TimeValue("00:00:01")
    ' Define the Power Query M code
80  pqMCode = _
        "let" & vbCrLf & _
        "    Source = Pdf.Tables(File.Contents(""" & pdfFilePath & """), [Implementation = ""1.1""])," & vbCrLf & _
        "    #""Filtered Rows"" = Table.SelectRows(Source, each Text.Contains([Id], ""Page""))," & vbCrLf & _
        "    TableIDs = Table.SelectColumns(#""Filtered Rows"",{""Id""})," & vbCrLf & _
        "    CleanedData = Table.Distinct(TableIDs)" & vbCrLf & _
        "in" & vbCrLf & _
        "    CleanedData"

    ' Define the query name
90  queryName = "ExtractTableIDs"

    ' Delete the query if it already exists
100 On Error Resume Next
110 wb.Queries(queryName).Delete 'If existing query delete it
120 On Error GoTo ErrorHandler ' Restore error handling

    ' Create a new query to extract Table IDs
130 wb.Queries.Add name:=queryName, Formula:=pqMCode

    ' Check and create PDF_Table_IDs if missing
140 sheetExists = False
150 For Each ws In wb.Sheets
160     If ws.name = "PDF_Table_IDs" Then
170         sheetExists = True
            Set ws = wb.Sheets("PDF_Table_IDs")
            MyDebugPrint vbCrLf & "Clearing OLD Data" & vbCrLf
            frmSplash.UpdateProgressBar 0.2
            frmSplash.UpdateCap "LAUNCHING - STEP 2/4 - READ and EXTRACT PDF SCHEMA - KEEP FINGERS CROSSED"
            Application.Wait Now + TimeValue("00:00:01")
            ws.Range("A:A").Delete
180         Exit For
190     End If
200 Next ws
210 If Not sheetExists Then
220     Set ws = wb.Sheets.Add
230     ws.name = "PDF_Table_IDs"
240 Else
250     Set ws = wb.Sheets("PDF_Table_IDs")
260 End If
MyDebugPrint "Fetching Data for Query :" & pqMCode & vbCrLf & "**********************************" & vvbcrlf
frmSplash.UpdateProgressBar 0.5
frmSplash.UpdateCap "LAUNCHING - STEP 2/4 - READ and EXTRACT PDF SCHEMA - KEEP FINGERS CROSSED"
    ' Load the query into the worksheet
270 With ws.ListObjects.Add(SourceType:=0, Source:="OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & queryName, Destination:=ws.Range("$A$1")).QueryTable
280     .CommandType = xlCmdSql
290     .CommandText = Array("SELECT * FROM [" & queryName & "]")
300     .RowNumbers = False
310     .FillAdjacentFormulas = False
320     .PreserveFormatting = True
330     .RefreshOnFileOpen = False
340     .BackgroundQuery = False
350     .RefreshStyle = xlInsertDeleteCells
360     .SavePassword = False
370     .SaveData = True
380     .AdjustColumnWidth = True
390     .PreserveColumnInfo = True
400     .Refresh BackgroundQuery:=False
410 End With
Application.Wait Now + TimeValue("00:00:01")
frmSplash.UpdateProgressBar 1
frmSplash.UpdateCap ("LAUNCHING - STEP 2/4 - READ and EXTRACT PDF SCHEMA - KEEP FINGERS CROSSED")
frmSplash.lblProgressText.Caption = "COMPLETED...Exiting in 1 sec"
Application.Wait Now + TimeValue("00:00:01")
420 'MsgBox "Table IDs extracted and loaded successfully! Execute Step 3", vbInformation

    wb.Queries(queryName).Delete
    Application.ScreenUpdating = True
    Unload frmSplash
    Set ws = wb.Sheets("Helper")
    If ws.Range("A2") = "" Then
    ws.Cells(2, 2).Value = "Provide Output File Name (with extension ex: .xlsx)"
    ws.Activate
    ws.Cells(2, 2).Select
    Selection.Style = "Bad"
    End If
    If ws.Range("A3") = "" Then
    ws.Range("B3").Value = "[o]Provide NAVAll.txt"
    ws.Range("B3").Select
    Selection.Style = "Bad"
    End If
    ws.Cells.EntireColumn.AutoFit
    ws.Activate
430 Exit Sub

ErrorHandler:
Application.ScreenUpdating = True
Unload frmSplash
440 MsgBox "An error occurred on line " & Erl & ": " & Err.Description, vbCritical, "Error " & Err.Number
End Sub


Function GetFundNameByISIN(ByVal isin As String, Optional ByVal whichone As Integer) As String
    ' Version 1.0.2 - Retrieves fund name by ISIN from navall file
    On Error GoTo ErrorHandler
    
    Dim http As Object
    Dim filepath As String
    Dim saveDialog As FileDialog
    Dim line As String
    Dim fileNum As Integer
    Dim url As String
    Dim todayDate As String
    Dim NAVExist As Boolean
    Dim wb As Workbook
    Set wb = Workbooks("pdf ingestion v0.3.xlsm") 'need to allocate dynamically
filepath = wb.Sheets("Helper").Range("A3").Value
If filepath = "" Then
      MyDebugPrint vbCrLf & "*********************************" & vbCrLf & "No NAV data path found in cell A3 of Helper. Please specify a valid path." & vbCrLf & "*********************************" & vbCrLf
      frmSplash.lblProgressText.Caption = "Error"
      Application.Wait Now + TimeValue("00:00:01")
      'MsgBox "NAV file missing"
      Exit Function
End If

If Dir(filepath) = "" Then
      MsgBox "The specified NAVAll file does not exist. Please check the path.", vbCritical
      Exit Function
End If
    ' Open the file to search for the ISIN
    fileNum = FreeFile
    Open filepath For Input As #fileNum
    Do While Not EOF(fileNum)
        Line Input #fileNum, line
        ' Check if the line contains the ISIN
        If InStr(line, isin) > 0 Then
            Dim fields() As String
            fields = Split(line, ";")
            If UBound(fields) >= 5 Then
                If Not IsNumeric(whichone) Then
                    fundName = fields(3) ' Extract the 4th field (Scheme Name)
                Else
                    fundName = fields(whichone) 'Extract whichone !!
                End If
            Else
                fundName = "Error: Invalid file format"
            End If
            Exit Do
        End If
    Loop
    Close #fileNum

    ' Return the fund name or indicate not found
    If fundName = "" Then
        MsgBox "ISIN Not found"
        GetFundNameByISIN = ""
    Else
        'MsgBox fundName
        GetFundNameByISIN = fundName
    End If
    Exit Function

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
End Function
Sub Step3_ReadPortfolioData()
On Error GoTo ErrorHandler

Dim wb As Workbook
Dim ws, wd, tableWs As Worksheet
Dim j, NoofIDs As Long
Dim Flag, sheetExists As Boolean
Dim tableID As String


'Default boolean is false
'Initialise workbook and worksheet
Set wb = ThisWorkbook

'check if extracted data in step 2 can be ommited in final
For Each wd In wb.Sheets
     If wd.name = "PDF_Table_IDs" Then
         sheetExists = True
         Exit For
     End If
 Next wd
 If Not sheetExists Then
    MsgBox "Run Step 2 first", vbExclamation, "Run the sequence"
    Exit Sub 'Code execution stops here
 End If

Set tableWs = wb.Sheets("PDF_Table_IDs") 'worksheet for table ids

'Read the PDF file path from the Helper sheet, should be fixed ref
filepath = wb.Sheets("Helper").Range("A1").Value
    If filepath = "" Then 'seems like goofup should not be empty
    MsgBox "No PDF file path found. Exiting...", vbExclamation
    Exit Sub
End If

' Loop through each table ID in the "PDF_Table_IDs" worksheet
    NoofIDs = tableWs.Cells(Rows.Count, 1).End(xlUp).row
    
frmSplash.Show vbModeless
Application.ScreenUpdating = False
'Start Master loop to get table data
For j = 2 To tableWs.Cells(Rows.Count, 1).End(xlUp).row
frmSplash.UpdateProgressBar j / NoofIDs
    tableID = tableWs.Cells(j, 1).Value
        If tableID = "" Then 'Empty table id how??? but skip if any
            Exit For
        End If
    'If "TableData_" tableID exist then delete it
    For Each wd In wb.Sheets
        If wd.name = "TableData_" & tableID Then
            Application.DisplayAlerts = False
            wd.Delete
            Application.DisplayAlerts = True
            Exit For
        End If
    Next wd
  Set ws = wb.Sheets.Add ' Create a new worksheet for the current table data
  ws.name = "TableData_" & tableID

' Build the Power Query formula for extracting data from the PDF table
    queryFormula = "let" & vbCrLf & _
        "    Source = Pdf.Tables(File.Contents(""" & filepath & """), [Implementation = ""1.1""])," & vbCrLf & _
        "    TableData = Source{[Id=""" & tableID & """]}[Data]," & vbCrLf & _
        "    CleanedData = Table.Distinct(TableData)" & vbCrLf & _
        "in" & vbCrLf & _
        "    CleanedData"

' Define query name dynamically
    queryName = "Query_" & tableID
On Error Resume Next
    If wb.Queries(queryName) Is Nothing Then
             wb.Queries.Add queryName, queryFormula
                Flag = True
            End If
    On Error GoTo ErrorHandler
    If Flag = False Then
                wb.Queries(queryName).Delete
                wb.Queries.Add queryName, queryFormula
    End If
    MyDebugPrint "Fetching data for :" & queryName & vbCrLf & "QUERY = " & queryFormula & "*************************************" & vbCrLf
    With ws.ListObjects.Add(SourceType:=0, Source:="OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & queryName, Destination:=ws.Range("$A$1")).QueryTable
     .CommandType = xlCmdSql
     .CommandText = Array("SELECT * FROM [" & queryName & "]")
     .RowNumbers = False
     .FillAdjacentFormulas = False
     .PreserveFormatting = True
     .RefreshOnFileOpen = False
     .BackgroundQuery = False
     .RefreshStyle = xlInsertDeleteCells
     .SavePassword = False
     .SaveData = True
     .AdjustColumnWidth = True
     .PreserveColumnInfo = True
     .Refresh BackgroundQuery:=False
    End With
    'Hidden for now Call Step4
    'Step4_readsheet ByVal ws.Name
    DeleteAllQueriesAndConnections
'    Application.DisplayAlerts = False
'    ws.Delete
'    Application.DisplayAlerts = True
Next j
    Application.ScreenUpdating = True
    frmSplash.lblProgressText.Caption = "COMPLETED...Exiting in 1 sec"
    Application.Wait Now + TimeValue("00:00:02")
    Unload frmSplash
    Set ws = wb.Sheets("Helper")
    ws.Select
    ws.Range("A1").Select
    Exit Sub
ErrorHandler:
    Application.ScreenUpdating = True
    frmSplash.lblProgressText.Caption = "ERROR...Exiting in 1 sec"
    Application.Wait Now + TimeValue("00:00:02")
    Unload frmSplash
    MsgBox "An error occurred on line in Step 3 " & ": " & Err.Description, vbCritical, "Error " & Err.Number
End Sub


Sub Step4_readsheets()
'version 3.0.0 starting again 14/12
'Old step 4 is in module 4 and older is in archive
'On Error GoTo ErrorHandler ' Enable error handling

Dim wb As Workbook, wbt As Workbook
Dim wd As Worksheet, ws As Worksheet, tableWs As Worksheet
Dim sheetExists As Boolean
Dim j As Long, i As Long, NoofIDs As Long
Dim tableID As String
Dim ISINC As Collection, openingbalance As Collection, transactionData As Collection
Dim sheetdump As Collection
Dim folioRange As Range
Dim openbal As Double
Dim transaction As Object
Dim TotalSteps As Long
Dim datestart As Date, dateend As Date

Set wb = ThisWorkbook
'check if extracted data in step 2 is found
For Each wd In wb.Sheets
     If wd.name = "PDF_Table_IDs" Then
         sheetExists = True
         Exit For
     End If
 Next wd
 If Not sheetExists Then
    MsgBox "Could Not find PDF Table IDs, Did you run Step 2 (and 1 and 3)", vbExclamation, "Run the sequence"
    Application.ScreenUpdating = True
    Unload frmSplash
    Exit Sub 'Code execution stops here
 End If
 If wb.Sheets("Helper").Range("A3").Value = "" Then MsgBox "Continuing without NAVAll.txt file" & vbCrLf & "Fund Name will not be captured for new folio and you may see an error", vbInformation, "Just an FYI"

'check if MF Data exist
If FileExists(wb.Sheets("Helper").Range("A2").Value) Then
    Set wbt = Workbooks.Open(ThisWorkbook.Path & "\" & wb.Sheets("Helper").Range("A2").Value)
    For Each wd In wbt.Sheets
        If wd.name = "Log" Then
            Set ws = wbt.Sheets("Log")
            Exit For
        End If
    Next wd
    If Not ws Is Nothing Then
        If ws.name = "Log" Then
            FindAndExtractDates datestart, dateend
            WBAdd = False
            If datestart < ws.Cells(ws.Cells(Rows.Count, 1).End(xlUp).row, 3).Value Then
                MsgBox "New CAS date is not new than already imported data, , Exiting", vbCritical, "Can not continue"
                Exit Sub
            Else
                'find last row
                Dim lastrow As Long
                lastrow = ws.Cells(Rows.Count, 1).End(xlUp).row + 1
                ws.Cells(lastrow, 1).Value = Format(Date, "dd/mm/yyyy")
                ws.Cells(lastrow, 2).Value = datestart
                ws.Cells(ws.Cells(Rows.Count, 1).End(xlUp).row, 3).Value = dateend
            End If
        End If
    End If
Else
WBAdd = True
End If
frmSplash.Show vbModeless
Application.ScreenUpdating = False
Set tableWs = wb.Sheets("PDF_Table_IDs") 'worksheet for table ids
NoofIDs = tableWs.Cells(Rows.Count, 1).End(xlUp).row
TotalSteps = NoofIDs
'Start finding folios
Foliocont = False
'Start second row, 1st is header
For j = 2 To NoofIDs
frmSplash.UpdateProgressBar j / TotalSteps
tableID = tableWs.Cells(j, 1).Value  'read the next sheet id
If tableID = "" Then 'Empty table id how??? but skip if any
    MsgBox "Check table ids there is an empty value at count: " & j - 1 & ", Exitting !!"
    Application.ScreenUpdating = True
    Unload frmSplash
    Exit Sub
End If
tableID = "TableData_" & tableID 'create the correct sheet name
Set ws = wb.Sheets(tableID)
If Foliocont = True Then
    If ws.Cells(2, 1).Value = "" Then
            'On Error Resume Next
            MyDebugPrint "FIND Transaction 1 Folio Flag: " & Foliocont & vbCrLf & "ISIN: " & isin & vbCr & "Folios: " & folios(i) & " Sheet: " & tableID
            Set transactionData = FindTransactions(ws, ws.Range("B2"), ws.Range("A1"), ws.Range("A6"))
            If transactionData Is Nothing Then
                Foliocont = False 'Closing
                GoTo lol
            End If
            Set closingbalanace = FindAllOccurrences(ws, "Closing", 1)
            If Not closingbalanace Is Nothing Then Foliocont = False 'MsgBox "closing"
            MyDebugPrint "writemfdata 1 Folio Flag: " & Foliocont & vbCrLf & "ISIN: " & isin & " Sheet: " & tableID & vbCr & "Folios: " & folios(i) & " Sheet: " & tableID
            If WriteMFData(wb.Sheets("Helper").Cells(2, 1).Value, transactionData, isin) = False Then MsgBox "Error writing file"
        Else
            MyDebugPrint "FIND Transaction 2 Folio Flag: " & Foliocont & vbCrLf & "ISIN: " & isin & vbCr & "Folios: " & folios(i) & " Sheet: " & tableID
            Set transactionData = FindTransactions(ws, ws.Range("A2"), ws.Range("A1"), ws.Range("A5"))
            If transactionData Is Nothing Then
                Foliocont = False 'Closing
                GoTo lol
            End If
            Set closingbalanace = FindAllOccurrences(ws, "Closing", 1)
            If Not closingbalanace Is Nothing Then Foliocont = False 'MsgBox "closing"
            MyDebugPrint "writemfdata 2 Folio Flag: " & Foliocont & vbCrLf & "ISIN: " & isin & vbCr & "Folios: " & folios(i) & " Sheet: " & tableID & vbCr & "**************************"
            If WriteMFData(wb.Sheets("Helper").Cells(2, 1).Value, transactionData, isin) = False Then MsgBox "Error writing file"
        End If
End If
    If Foliocont = False Then
    Set folios = FindAllOccurrences(ws, "Folio No", 2)
    If Not folios Is Nothing Then
        For i = 1 To folios.Count
            Set folioRange = folios(i)
            Foliocont = True
            ISFound = False
            'MsgBox folioRange.Address & " : " & folioRange.Value
            Set ISINC = FindAllOccurrences(ws, "ISIN", 1, folioRange)
            isin = ExtractISINFromCollection(ISINC)
            'MsgBox ISINC(1) & " : " & ISIN
            Set openingbalance = FindAllOccurrences(ws, "Opening", 1, ISINC(1))
            If Not openingbalance Is Nothing Then
                openbal = FindDouble(ws, openingbalance(1).Address)
                ISFound = True
                'MsgBox openingbalance(1) & " : " & openingbalance(1).Address & " : " & openbal
            Else
                Set openingbalance = ISINC
            End If
            MyDebugPrint "FIND Transaction 5 Folio Flag: " & Foliocont & vbCrLf & "ISIN: " & isin & vbCr & "Folios: " & folios(i) & " Sheet: " & tableID
            Set transactionData = FindTransactions(ws, folios(i), openingbalance(1), ISINC(1))
            If transactionData Is Nothing Then
                Foliocont = False 'Closing
                GoTo lol
            End If
            Set closingbalanace = FindAllOccurrences(ws, "Closing", 1, folioRange)
            If Not closingbalanace Is Nothing Then
                Foliocont = False 'MsgBox "closing"
                MyDebugPrint "writemfdata 5 Folio Flag: " & Foliocont & vbCr & "ISIN: " & isin & vbCr & "Folios: " & folios(i) & vbCr & "Openbal: " & openbal & " Sheet: " & tableID & vbCr & "**************************"
                If WriteMFData(wb.Sheets("Helper").Cells(2, 1).Value, transactionData, isin, folios(i), openbal) = False Then MsgBox "Error writing file"
                'MsgBox folios(i)
            Else
                Foliocont = True 'continue with folio
                MyDebugPrint "writemfdata 6 Folio Flag: " & Foliocont & vbCrLf & "ISIN: " & isin & vbCrLf & "Folios: " & folios(i) & vbCrLf & "Openbal: " & openbal & " Sheet: " & tableID
                If WriteMFData(wb.Sheets("Helper").Cells(2, 1).Value, transactionData, isin, folios(i), openbal) = False Then MsgBox "Error writing file"
                Exit For
            End If
            'MsgBox closingbalanace(1).Address
        Next i
    End If
    End If
'End If
lol:
Next j
Module3.autofitcolumns (wb.Sheets("Helper").Cells(2, 1).Value)
'Update CAS Dates
If WBAdd = True Then
    Set wb = Workbooks(wb.Sheets("Helper").Cells(2, 1).Value)
    Set ws = wb.Sheets("Sheet1")
    ws.name = "Log"
    ws.Range("A1").Value = "Update Date"
    ws.Range("B1").Value = "Start Date"
    ws.Range("C1").Value = "End Date"
    ws.Range("A2").Value = Format(Date, "dd/mm/yyyy")
    FindAndExtractDates datestart, dateend
    ws.Range("B2").Value = datestart
    ws.Range("C2").Value = dateend
    ws.Cells.EntireColumn.AutoFit
    ws.Activate
    wb.Save
End If
Application.ScreenUpdating = True
frmSplash.lblProgressText.Caption = "COMPLETED...Exiting in 1 sec"
Application.Wait Now + TimeValue("00:00:02")
Unload frmSplash
Exit Sub
'ErrorHandler:
MsgBox "An error occurred: in Step4 " & Err.Description, vbCritical, "Error " & Err.Number
End Sub

Function FindAllOccurrences(ws As Worksheet, searchText As String, Optional ByVal multi As Integer, Optional ByVal after As Range) As Collection
    ' Version 1.0.0 - Find and track all occurrences of a search term (e.g., "Folio")
    'On Error GoTo ErrorHandler ' Enable error handling

    Dim foundcell As Range
    Dim firstAddress As String
    Dim occurrences As New Collection
    Dim startRow As Long

    ' Start searching from the first cell
    'Set foundCell = ws.Cells.Find(What:=searchText, LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
    If after Is Nothing Then
        Set after = ws.Cells(1, 1)
        startRow = 0
    Else
        startRow = after.row
    End If
    
    Set foundcell = ws.Cells.Find(What:=searchText, after:=after, _
                                  LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
    If Not foundcell Is Nothing And multi > 1 Then
        firstAddress = foundcell.Address ' Track the first found cell

        Do
             'Add the found cell to the collection
            occurrences.Add foundcell

            ' Continue searching for the next occurrence
            Set foundcell = ws.Cells.FindNext(foundcell)
        Loop While Not foundcell Is Nothing And foundcell.Address <> firstAddress
    ElseIf Not foundcell Is Nothing And multi = 1 Then
        If foundcell.row >= startRow Then
            occurrences.Add foundcell
        Else
            Set FindAllOccurrences = Nothing
            Exit Function
        End If
    ElseIf foundcell Is Nothing Then
        Set FindAllOccurrences = Nothing
        Exit Function
    End If

    ' Return the collection of found cells
    Set FindAllOccurrences = occurrences
    Exit Function

ErrorHandler:
    ' Handle any errors
    MsgBox "An error occurred: in FindAllOccurrences " & Err.Description, vbCritical, "Error " & Err.Number
    Set FindAllOccurrences = Nothing
End Function

Function ExtractISINFromCollection(occurrences As Collection) As String
    ' Version 1.0.1 - Extract a single 13-character ISIN from a collection of cells
    'On Error GoTo ErrorHandler ' Enable error handling

    Dim cell As Range
    Dim cellContent As String
    Dim regex As Object
    Dim matches As Object

    ' Create a RegExp object for matching ISIN pattern
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "ISIN:[A-Z0-9]{12}" ' Matches ISIN:<12 alphanumeric characters>
    regex.Global = True

    ' Loop through each cell in the collection
    For Each cell In occurrences
        cellContent = cell.Value
        
        ' Match the ISIN pattern in the cell content
        If regex.test(cellContent) Then
            Set matches = regex.Execute(cellContent)
            ' Extract and return the ISIN without the "ISIN:" prefix
            ExtractISINFromCollection = Mid(matches(0), 6, 13) ' Extract the 13-character ISIN
            Exit Function ' Return on the first match
        End If
    Next cell

    ' If no ISIN is found, return an empty string
    ExtractISINFromCollection = ""
    Exit Function

ErrorHandler:
    ' Handle errors
    MsgBox "An error occurred while extracting ISIN: " & Err.Description, vbCritical, "Error " & Err.Number
    ExtractISINFromCollection = ""
End Function

Function FindDouble(ws As Worksheet, cellAddress As String) As Double
    ' Searches for the first Double value in the row of the given address, starting from the column in cellAddress
    Dim searchRow As Long
    Dim startColumn As Long
    Dim cell As Range
    Dim cellValue As Variant
    Dim DoubleCell As Double

    'On Error GoTo ErrorHandler ' Enable error handling

    ' Extract the row and column information from the address
    searchRow = ws.Range(cellAddress).row ' Extract the row from the address
    startColumn = ws.Range(cellAddress).Column ' Extract the column from the address

    ' Loop through all cells in the specified row, starting from the column passed in cellAddress
    For Each cell In ws.Rows(searchRow).Cells
        If cell.Column >= startColumn Then
            cellValue = cell.Value
            ' Check if the cell value is numeric
            If IsNumeric(cellValue) And Not IsEmpty(cellValue) Then
                ' Try to convert the value to Double
                On Error Resume Next
                DoubleCell = cellValue
                'On Error GoTo ErrorHandler
                If VarType(DoubleCell) = vbDouble Then
                    'MsgBox "Double value found: " & CStr(DoubleCell), vbInformation, "Value Found"
                    FindDouble = DoubleCell
                    Exit Function
                End If
            End If
        End If
    Next cell

    ' If no Double value is found
    MsgBox "No Double value found in row " & searchRow & " of worksheet: " & ws.name, vbInformation
    FindDouble = 0 ' Return 0 if no Double value is found
    Exit Function

ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    FindDouble = 0 ' Return 0 on error
End Function

Function FindTransactions(ws As Worksheet, folioRange As Range, startRange As Range, Optional isinrange As Range) As Collection
    ' Version 1.0.1 - Complete function to find transactions with specified criteria
    'On Error GoTo ErrorHandler ' Enable error handling
    
    Dim currentRow As Long
    Dim lastcol As Long
    Dim rowData As Object
    Dim result As New Collection
    Dim row As Range
    Dim keyCounter As Integer
    Dim numFound As Boolean
    Dim colIndex As Integer
    Dim numValue As Double
    Dim rowOutput As String
    Dim spcltrans As Boolean
    Dim datestart As Date, dateend As Date
    
    ' Initialize row counter (starting from startRange row)
    currentRow = WorksheetFunction.Max(folioRange.row, startRange.row, isinrange.row) + 1
    'If startRange.row = isinrange.row And Not ISFound Then currentRow = currentRow + 1
    'check if we need to bump the current row or not
    Set row = ws.Rows(currentRow)
    If Not IsEmpty(row.Cells(folioRange.Column).Value) Then
        If IsDate(row.Cells(folioRange.Column).Value) Then
        Else
            currentRow = currentRow + 1
        End If
'    Else
'        currentRow = currentRow + 1
    End If
    ' Loop through the rows until the last non-empty row in the column
    Do While currentRow <= ws.Cells(ws.Rows.Count, folioRange.Column).End(xlUp).row
        spcltrans = False
        Set row = ws.Rows(currentRow)
        Set rowData = CreateObject("Scripting.Dictionary") ' Dictionary to hold row data

        keyCounter = 0 ' Reset key counter for each row
        lastcol = ws.Cells(currentRow, ws.Columns.Count).End(xlToLeft).Column

        ' Loop through each column starting from folioRange.Column
        For colIndex = folioRange.Column To lastcol
            ' Check if cell is not empty
            If Not IsEmpty(row.Cells(colIndex).Value) Then
                ' Check for date in the first non-empty column
                If keyCounter = 0 And IsDate(row.Cells(colIndex).Value) Then
                    rowData.Add "Date", row.Cells(colIndex).Value
                    keyCounter = keyCounter + 1
                'CHECK For no NEW Transaction
                ElseIf keyCounter = 0 And VarType(row.Cells(colIndex).Value) = vbString Then
                    If InStr(row.Cells(colIndex).Value, "No transactions during this statement period") Then
                        FindAndExtractDates datestart, dateend
                        rowData.Add "Date", dateend
                        rowData.Add "Transaction", row.Cells(colIndex).Value
                        result.Add rowData
                        Set FindTransactions = result
                        Exit Function
                    End If
                ' Check for transaction (text)
                ElseIf keyCounter = 1 And VarType(row.Cells(colIndex).Value) = vbString Then
                    rowData.Add "Transaction", row.Cells(colIndex).Value
                    keyCounter = keyCounter + 1
                    If InStr(row.Cells(colIndex).Value, "***") Then spcltrans = True
                ' Check for amount (double)
                ElseIf keyCounter = 2 And IsNumeric(row.Cells(colIndex).Value) Then
                    numValue = row.Cells(colIndex).Value
                    rowData.Add "Amount", row.Cells(colIndex).Value
                    If spcltrans = True Then keyCounter = 5
'                    If InStr(1, CStr(row.Cells(colIndex).Value), "(") > 0 Then
'                        numValue = -Abs(numValue)
'                        rowData.Add "Amount", numValue
'                    End If
                    keyCounter = keyCounter + 1
                ' Check for units (double)
                ElseIf keyCounter = 3 And IsNumeric(row.Cells(colIndex).Value) Then
                    numValue = row.Cells(colIndex).Value
                    rowData.Add "Units", row.Cells(colIndex).Value
'                    If InStr(1, CStr(row.Cells(colIndex).Value), "(") > 0 Then
'                        numValue = -Abs(numValue)
'                        rowData.Add "Units", numValue
'                    End If
                    keyCounter = keyCounter + 1
                ' Check for price (double)
                ElseIf keyCounter = 4 And IsNumeric(row.Cells(colIndex).Value) Then
                    rowData.Add "Price", row.Cells(colIndex).Value
                    keyCounter = keyCounter + 1
                ' Check for unit balance (double)
                ElseIf keyCounter = 5 And IsNumeric(row.Cells(colIndex).Value) Then
                    rowData.Add "Unit Balance", row.Cells(colIndex).Value
                    keyCounter = keyCounter + 1
                ElseIf keyCounter = 0 And Not IsDate(row.Cells(folioRange.Column).Value) And result.Count > 1 Then
                    'MsgBox "Exiting function: Row " & currentRow & " has a non-date value in the first non-empty cell."
                    MyDebugPrint result.Count & " results found, Exiting function: Row " & currentRow & " has a non-date value in the first non-empty cell."
                    Set FindTransactions = result ' Return the rows processed so far
                    Exit Function
                End If
                ' If all keys are found, exit the loop early
                If keyCounter = 6 Then Exit For
            End If
        Next colIndex

        ' Check if all keys are found
        If keyCounter = 6 Then
            result.Add rowData
            ' Print the successful row output for debugging
            'rowOutput = "Row " & currentRow & " - Date: " & rowData("Date") & "; Transaction: " & rowData("Transaction") & "; Amount: " & rowData("Amount") & "; Units: " & rowData("Units") & "; Price: " & rowData("Price") & "; Unit Balance: " & rowData("Unit Balance")
            'Debug.Print rowOutput
        Else
            ' Error handling: If the first non-empty cell isn't a valid date, exit the function
            If keyCounter = 0 And Not IsDate(row.Cells(folioRange.Column).Value) And result.Count < 1 Then
                'MsgBox "Error: First non-empty value in row " & currentRow & " is not a valid date. Exiting."
                MyDebugPrint "Error: First non-empty value in row " & currentRow & " is not a valid date. No result found."
                Set FindTransactions = Nothing ' Return Nothing if no valid data
                Exit Function
            End If
        End If

        ' Exit condition: If the first non-empty cell in the row is not a date, exit the function
        If keyCounter > 0 And Not IsDate(row.Cells(folioRange.Column).Value) And result.Count > 1 Then
            'MsgBox "Exiting function: Row " & currentRow & " has a non-date value in the first non-empty cell."
            MyDebugPrint "Return function as Row " & currentRow & " has a non-date value in the first non-empty cell."
            Set FindTransactions = result ' Return the rows processed so far
            Exit Function
        End If

        ' Move to the next row
        currentRow = currentRow + 1
    Loop

    ' If no valid rows were found
    Set FindTransactions = result
    Exit Function

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
    Set FindTransactions = Nothing
End Function

Function WriteMFData(fileName As String, transactionData As Collection, Optional isin As String, Optional folio As String, Optional openbal As Double) As Boolean
    'Version 3.2.15 - Writes mutual fund transaction data to a worksheet
    'On Error GoTo ErrorHandler ' Enable error handling

    Dim wb As Workbook
    Dim ws As Worksheet, tempws As Worksheet
    Dim outputRow As Long
    Dim currentTransaction As Object
    Dim colHeaders As Variant
    Dim i As Long, j As Long
    Dim filepath As String
    Dim tsheetFound As Boolean, newtsheet As Boolean
    
    filepath = ThisWorkbook.Path & "\" & fileName

    tsheetFound = False
        For Each wb In Application.Workbooks
            If wb.name = fileName Then
                Exit For
            End If
        Next wb

If wb Is Nothing Then
        On Error Resume Next
        Set wb = Workbooks.Open(filepath)
        'On Error GoTo ErrorHandler
        If wb Is Nothing Then
            'PE1 - create file if not found instead of error
             Set wb = Workbooks.Add
            ' Optionally, save the new workbook with the specified path and name
            wb.SaveAs fileName:=filepath
            'new workbook tasks
            WBAdd = True
        End If
End If

' Look for the sheet with the specified name
For Each ws In wb.Sheets
        If ws.name = isin Then
            tsheetFound = True
            newtsheet = False
            Exit For
        End If
    Next ws
' If not found, create the sheet
If Not tsheetFound Then
    Set ws = wb.Sheets.Add(after:=wb.Sheets(wb.Sheets.Count))
    ws.name = isin
    tsheetFound = True
    newtsheet = True
End If
    ' Successfully opened the file and found/created the sheet
    
    'On Error GoTo ErrorHandler

    ' Prepare the headers for the data table
    colHeaders = Array("Date", "Transaction", "Amount", "Units", "Price", "Unit Balance")
    
    If newtsheet And tsheetFound Then
        ws.Range("A1").Value = folio
        ws.Range("B1").Value = isin
        If GetFundNameByISIN(Right(isin, 12), 3) <> "" Then ws.Range("C1").Value = GetFundNameByISIN(Right(isin, 12), 3)
        ws.Range("A2").Value = "Opening balance :" & openbal
    ' Write headers if the sheet is empty
        For j = LBound(colHeaders) To UBound(colHeaders)
            ws.Cells(3, j + 1).Value = colHeaders(j)
        Next j
        outputRow = 4
    Else
        outputRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row + 1
    End If
    For i = 1 To transactionData.Count
    Set currentTransaction = transactionData.Item(i)
    For j = LBound(colHeaders) To UBound(colHeaders)
    If currentTransaction.Exists(colHeaders(j)) Then
        ws.Cells(outputRow, j + 1).Value = currentTransaction(colHeaders(j))
    Else
        ws.Cells(outputRow, j + 1).Value = "N/A"
    End If
    Next j
    outputRow = outputRow + 1
    Next i
    
    'Add closing bal if flag is false
    If Foliocont = False Then ws.Range("B2").Value = closingbalanace(1).Value
    
    ' Save and close the workbook
    wb.Save

    ' Return success
    WriteMFData = True
    'MsgBox "Data written successfully to " & fileName, vbInformation
    Exit Function

ErrorHandler:
    MsgBox "An error occurred on line " & Erl & ": " & Err.Description, vbCritical, "Error " & Err.Number
    WriteMFData = False
End Function

Function FileExists(filepath As String) As Boolean
    On Error Resume Next
    FileExists = (Dir(filepath) <> "")
    'MsgBox (Dir(filepath) <> "")
    On Error GoTo 0
End Function
Sub MyDebugPrint(message As String)
    Debug.Print message ' Traditional debug print for developer's immediate window turn on if you need to debug
    frmSplash.TextBox1.MultiLine = True
    frmSplash.TextBox1.Value = message & frmSplash.TextBox1.Value ' Update the ListBox in the splash screen
    frmSplash.TextBox1.SelStart = Len(frmSplash.TextBox1.Text)
    frmSplash.Repaint ' Refresh the form to show new messages immediately
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

Sub FindAndExtractDates(ByRef datestart As Date, ByRef dateend As Date)
    Dim ws As Worksheet
    Dim wb As Workbook
    Dim foundcell As Range
    Dim fullText As String
    Dim dateParts() As String
    Dim dateText As String
    Dim startPos As Long
    
    CASFound = False

    Set wb = ThisWorkbook
    'check if extracted data in step 2 is found
    For Each ws In wb.Sheets
    If CASFound = True Then Exit Sub
    If ws.name <> "Sheet1" And ws.name <> "PDF_Table_IDs" And ws.name <> "Helper" Then
        Set foundcell = ws.Cells.Find(What:="Consolidated Account Statement", LookIn:=xlValues, LookAt:=xlPart)
        If Not foundcell Is Nothing Then
        On Error Resume Next
            fullText = foundcell.Value
            If Len(fullText) <> 57 Then Exit Sub
            startPos = InStr(fullText, "Consolidated Account Statement") + Len("Consolidated Account Statement") + 1
            dateText = Mid(fullText, startPos, Len(fullText) - startPos + 1)
            dateParts = Split(dateText, "To")
            datestart = dateValue(Trim(dateParts(0)))
            dateend = dateValue(Trim(dateParts(1)))
            CASFound = True
            'Debug.Print dateStart & "_" & dateEnd
        End If
    End If
    Next ws
End Sub
