Attribute VB_Name = "forcewindowtop"
Option Explicit

Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function AttachThreadInput Lib "user32" (ByVal idAttach As Long, ByVal idAttachTo As Long, ByVal fAttach As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
 
Public Sub ForceWindowToTop(hWnd As Long)
   
    Dim lMyPId As Long
    Dim lCurPId As Long
   
    lMyPId = GetWindowThreadProcessId(hWnd, 0)
    lCurPId = GetWindowThreadProcessId(GetForegroundWindow(), 0)
   
    If Not (lMyPId = lCurPId) Then
        AttachThreadInput lCurPId, lMyPId, True
        SetForegroundWindow hWnd
        AttachThreadInput lCurPId, lMyPId, False
    End If
   
    If Not (GetForegroundWindow() = hWnd) Then
        SetForegroundWindow hWnd
    End If
   
End Sub
