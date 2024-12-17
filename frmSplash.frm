VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSplash 
   Caption         =   "UserForm1"
   ClientHeight    =   5775
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12750
   OleObjectBlob   =   "frmSplash.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long

Private Const GWL_STYLE As Long = -16
Private Const WS_CAPTION = &HC00000

' ... rest of your code

Private Sub UserForm_Activate()
    Dim hwnd As Long
    hwnd = FindWindow("ThunderDFrame", Me.Caption)
    SetWindowLong hwnd, GWL_STYLE, GetWindowLong(hwnd, GWL_STYLE) And Not WS_CAPTION
    DrawMenuBar hwnd
    
End Sub
Public Sub UpdateProgressBar(Percentage As Single)
    With Me
        .lblProgressBar.Width = Percentage * (.frmBarFrame.Width - 10) ' Subtract 10 for better look
        .lblProgressText.Caption = Format(Percentage, "0%") & " Complete"
        DoEvents ' Ensures the form updates
    End With
    If Percentage = 100 Then
    Me.lblProgressText.Caption = "COMPLETED"
Application.Wait Now + TimeValue("00:00:05")
End If
End Sub

Public Sub HideProgressBar()
    Me.lblProgressBar.Width = 0
    Me.lblProgressText.Caption = "0% Complete"
    ' If you want to hide the bar entirely, you might consider making these controls invisible instead
End Sub

Private Sub UserForm_Deactivate()


End Sub
