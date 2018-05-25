Attribute VB_Name = "Module1"
Option Explicit

'This section is for the API declares
Private Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long)
Private Declare Function SetWindowLong& Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

'This section is for Progressbar's
Private Const WM_USER = &H400
Private Const PBM_SETBARCOLOR = (WM_USER + 9)
Private Const CCM_FIRST = &H2000
Private Const CCM_SETBKCOLOR = (CCM_FIRST + 1)

'This section is for the Colors of the ProgressBar's
Public Sub PBcolor(PB As ProgressBar, Backcolor As Long, Forecolor As Long)
SendMessage PB.hwnd, CCM_SETBKCOLOR, 0, ByVal Backcolor
SendMessage PB.hwnd, PBM_SETBARCOLOR, 0, ByVal Forecolor
End Sub

