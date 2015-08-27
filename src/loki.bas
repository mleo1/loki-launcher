Attribute VB_Name = "loki"
'mleo1

Option Explicit

Private Type LOKI_SETTINGS
    ROExe As String
    MD5 As Boolean
    AddArg As String
    user As String
    Keep As Boolean
End Type

Public LSETTINGS As LOKI_SETTINGS
Public INI_NAME As String
Public EXE_PID As Double
Public LokiDir As String
Public ROExePath As String
Public IniPath As String
Public Mode As String
Public Lokis As Control
Public MX, MY As Single

Sub Main()
    INI_NAME = App.EXEName & ".ldat"
    LokiDir = App.Path & "\"
    IniPath = LokiDir & INI_NAME
    
    If CheckFile(IniPath) = False Then CreateINI

    GetSettings

    ROExePath = LokiDir & LSETTINGS.ROExe
        
    frmLauncher.Show
End Sub

Sub CreateINI()
    Open INI_NAME For Output As #1
    Print #1, "[Loki Launcher]"
    Print #1, "User="
    Print #1, "Keep=false"
    Close #1
End Sub

Sub GetSettings()
    Dim str As String
    Dim tmp() As String
    
    str = StrConv(LoadResData(1, "SETTING"), vbUnicode)
    
    tmp = Split(str, vbCrLf)

    LSETTINGS.ROExe = tmp(0)
    LSETTINGS.MD5 = IIf(tmp(1) = "true", True, False)
    LSETTINGS.AddArg = tmp(2)
    
    LSETTINGS.user = ReadIni("Loki Launcher", "User")

    LSETTINGS.Keep = IIf(LCase(ReadIni("Loki Launcher", "Keep")) = "true", True, False)
End Sub

Function CheckFile(strPath As String) As Boolean
    If Dir(strPath) <> vbNullString Then
        CheckFile = True
    Else
        CheckFile = False
    End If
End Function

Function UnloadAll()
    Dim frm As Form
    For Each frm In Forms
      Unload frm
    Next
End Function

Function HLightTxt(txtbox As TextBox)
    txtbox.SelStart = 0
    txtbox.SelLength = Len(txtbox)
End Function

Function cPass(ByVal str As String) As String
    If LSETTINGS.MD5 = True Then
        cPass = DigestStrToHexStr(str)
    Else
        cPass = str
    End If
End Function
