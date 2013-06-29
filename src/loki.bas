Attribute VB_Name = "loki"
'mleo1

Option Explicit

Public Const debugMode As Integer = 0

Private Type LOKI_SETTINGS
    Title As String
    ROExe As String
    GRF As String
    Patchsite As String
    BG As String
    Mode As String
    MD5 As String
    EXEArg As String
    
    User As String
    Keep As String
End Type

Public INI_SETTINGS As LOKI_SETTINGS
Public INI_NAME As String
Public LokiLoc As String
Public ROExeLoc As String
Public IniLoc As String
Public BGLoc As String

Public Mode As String
Public Lokis As Control

Sub Main()
    'If App.PrevInstance Then Exit Sub
       
    INI_NAME = LCase(App.EXEName & ".ini")
    LokiLoc = App.Path & "\"
    IniLoc = LokiLoc & INI_NAME
        
    getINIData

    ROExeLoc = LokiLoc & INI_SETTINGS.ROExe
    BGLoc = Replace(App.Path & INI_SETTINGS.BG, "/", "\")

    If checkFiles = 0 Then
        unloadALL
        Exit Sub
    End If

If debugMode > 0 Then GoTo dmode
    If INI_SETTINGS.Mode = "2" Then
dmode:
        frmLauncher.Show
    Else
        frmPatcher.Show
    End If
End Sub

Function checkFiles() As Integer
If debugMode > 0 Then GoTo dmode

    If checkFile(IniLoc) = False Then
        MsgBox "Can't find " & IniLoc, vbCritical, "Loki Launcher"
        checkFiles = 0
        Exit Function
    End If
    
    If checkFile(ROExeLoc) = False Then
        MsgBox "Can't find " & ROExeLoc, vbCritical, INI_SETTINGS.Title
        checkFiles = 0
        Exit Function
    End If
    
    If checkFile(BGLoc) = False Then
        MsgBox "Can't find " & BGLoc, vbCritical, INI_SETTINGS.Title
        checkFiles = 0
        Exit Function
    End If
    
    'If checkFile(LokiLoc & "7z.dll") = False Then
    '    MsgBox "Can't find " & LokiLoc & "7z.dll", vbCritical, INI_SETTINGS.Title
    '    checkFiles = 0
    '    Exit Function
    'End If
    '
    'If checkFile(LokiLoc & "VszLib.dll") = False Then
    '    MsgBox "Can't find " & LokiLoc & "VszLib.dll", vbCritical, INI_SETTINGS.Title
    '    checkFiles = 0
    '    Exit Function
    'End If
    '
    'If checkFile(LokiLoc & "GRF.dll") = False Then
    '    MsgBox "Can't find " & LokiLoc & "GRF.dll", vbCritical, INI_SETTINGS.Title
    '    checkFiles = 0
    '    Exit Function
    'End If
    '
    'If checkFile(LokiLoc & "GRFvb6.dll") = False Then
    '    MsgBox "Can't find " & LokiLoc & "GRFvb6.dll", vbCritical, INI_SETTINGS.Title
    '    checkFiles = 0
    '    Exit Function
    'End If
    
dmode:
    checkFiles = 1
End Function

Sub getINIData()
    INI_SETTINGS.Title = ReadIni("Settings", "Title")
    INI_SETTINGS.ROExe = ReadIni("Settings", "Exe")
    'INI_SETTINGS.GRF = ReadIni("Settings", "GRF")
    'INI_SETTINGS.Patchsite = ReadIni("Settings", "Patchsite")
    INI_SETTINGS.BG = ReadIni("Settings", "BG")
    'INI_SETTINGS.Mode = ReadIni("Settings", "Mode")
INI_SETTINGS.Mode = 2 'mleo
    INI_SETTINGS.MD5 = LCase(ReadIni("Settings", "MD5"))
    INI_SETTINGS.EXEArg = ReadIni("Settings", "ExeArg")
    
    INI_SETTINGS.User = ReadIni("Loki Launcher", "User")
    INI_SETTINGS.Keep = LCase(ReadIni("Loki Launcher", "Keep"))
End Sub

Function checkFile(strPath As String) As Boolean
    If Dir$(strPath) <> vbNullString Then
        checkFile = True
    Else
        checkFile = False
    End If
End Function

Function unloadALL()
    Set Lokis = Nothing

    Dim frm As Form
    For Each frm In Forms
      Unload frm
    Next
    Set frm = Nothing
End Function

Function hlightTxt(txtbox As TextBox)
    txtbox.SelStart = 0
    txtbox.SelLength = Len(txtbox)
End Function
