Attribute VB_Name = "movebform"
Option Explicit
 
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
 
Private Const HTCAPTION = 2
Private Const WM_NCLBUTTONDOWN = &HA1
 
Private Const WM_SYSCOMMAND = &H112

Public Function MoveForm(Form As Form)
    ReleaseCapture
    SendMessage Form.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Function

