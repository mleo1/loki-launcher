Attribute VB_Name = "pidcheckrunning"
Option Explicit
 
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" ( _
    ByVal lpClassName As String, _
    ByVal lpWindowName As String _
) As Long
 
Private Declare Function GetParent Lib "user32" ( _
    ByVal hWnd As Long _
) As Long
 
Private Declare Function GetWindowThreadProcessId Lib "user32" ( _
   ByVal hWnd As Long, _
   lpdwProcessId As Long _
) As Long
Private Declare Function GetWindow Lib "user32.dll" ( _
    ByVal hWnd As Long, _
    ByVal wCmd As Long _
) As Long
 
Private Const GW_HWNDNEXT As Long = 2
    
Public Function isPIDRunning(ByVal Pid As Long) As Boolean
    If Pid = -1 Then
    isPIDRunning = False
    Exit Function
    End If

    Dim tmpHWND As Long
    Dim tmpPID As Long
    Dim tmpID As Long
 
    tmpHWND = FindWindow(vbNullString, vbNullString)
 
    Do While tmpHWND <> 0
        If GetParent(tmpHWND) = 0 Then
            tmpID = GetWindowThreadProcessId(tmpHWND, tmpPID)
            If tmpPID = Pid Then
                    'hwnd = tmpHWND
                    isPIDRunning = True
                Exit Do
            End If
        End If
 
        tmpHWND = GetWindow(tmpHWND, GW_HWNDNEXT)
    Loop
End Function
