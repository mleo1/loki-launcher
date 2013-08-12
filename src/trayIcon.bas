Attribute VB_Name = "trayIcon"
Option Explicit
    
Private Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Private Const NIM_ADD = 0
Private Const NIM_MODIFY = 1
Private Const NIM_DELETE = 2
Private Const NIF_MESSAGE = 1
Private Const NIF_ICON = 2
Private Const NIF_TIP = 4

Private Declare Function Shell_NotifyIconA Lib "SHELL32" _
    (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Integer

Private Function setNOTIFYICONDATA(hWnd As Long, ID As Long, _
    Flags As Long, CallbackMessage As Long, Icon As Long, _
    Tip As String) As NOTIFYICONDATA
    
    Dim nidTemp As NOTIFYICONDATA
    
    nidTemp.cbSize = Len(nidTemp)
    nidTemp.hWnd = hWnd
    nidTemp.uID = ID
    nidTemp.uFlags = Flags
    nidTemp.uCallbackMessage = CallbackMessage
    nidTemp.hIcon = Icon
    nidTemp.szTip = Tip & Chr$(0)
    
    setNOTIFYICONDATA = nidTemp
End Function

Public Function trayIconAdd(Form As Form, tag As String)
    Dim nid As NOTIFYICONDATA
    
    nid = setNOTIFYICONDATA( _
        Form.hWnd, vbNull, _
        NIF_MESSAGE Or NIF_ICON Or NIF_TIP, _
        vbNull, Form.Icon, tag)
    
   Shell_NotifyIconA NIM_ADD, nid
End Function

Public Function trayIconDel(Form As Form)
    Dim nid As NOTIFYICONDATA
    
    nid = setNOTIFYICONDATA( _
        Form.hWnd, vbNull, _
        NIF_MESSAGE Or NIF_ICON Or NIF_TIP, _
        vbNull, Form.Icon, vbNullString)
    
    Shell_NotifyIconA NIM_DELETE, nid
End Function
