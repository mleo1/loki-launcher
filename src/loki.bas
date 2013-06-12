Attribute VB_Name = "loki"
Option Explicit

Public Const debugMode As Integer = 0

Private Type LOKI_SETTINGS
    Title As String
    ROExe As String
    BG As String
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

Public Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

Sub Main()
    If App.PrevInstance Then Exit Sub
        
    INI_NAME = frmRes.Caption
    LokiLoc = App.Path & "\"
    IniLoc = LokiLoc & INI_NAME
        
    getINIData

    ROExeLoc = LokiLoc & INI_SETTINGS.ROExe
    BGLoc = Replace(App.Path & INI_SETTINGS.BG, "/", "\")

    If checkFiles = 0 Then
        unloadAllForms
        Exit Sub
    End If
    
    frmLauncher.Show
End Sub

Function checkFiles() As Integer
If debugMode > 0 Then GoTo dMode

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
    
dMode:
    checkFiles = 1
End Function

Sub getINIData()
    INI_SETTINGS.Title = ReadIni("Settings", "Title")
    INI_SETTINGS.ROExe = ReadIni("Settings", "Exe")
    INI_SETTINGS.BG = ReadIni("Settings", "BG")
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

Function unloadAllForms()
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
