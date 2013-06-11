Attribute VB_Name = "ini"
Option Explicit

Public INI_NAME As String

Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Function ReadIni(Section As String, Key As String) As String
    Dim RetVal As String * 255, v As Long
    v = GetPrivateProfileString(Section, Key, "", RetVal, 255, App.Path & "\" & INI_NAME)
    ReadIni = Left(RetVal, v)
End Function
 
Public Function ReadIniSection(Section As String) As String
    Dim RetVal As String * 255, v As Long
    v = GetPrivateProfileSection(Section, RetVal, 255, App.Path & "\" & INI_NAME)
    ReadIniSection = Left(RetVal, v - 1)
End Function
 
Public Sub WriteIni(Section As String, Key As String, Value As String)
    WritePrivateProfileString Section, Key, Value, App.Path & "\" & INI_NAME
End Sub
 
Public Sub WriteIniSection(Section As String, Value As String)
    WritePrivateProfileSection Section, Value, App.Path & "\" & INI_NAME
End Sub
