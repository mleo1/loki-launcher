Attribute VB_Name = "Module1"
Option Explicit

Function unloadAllForms()
      Dim frm As Form
      For Each frm In Forms
        Unload frm
      Next
      Set frm = Nothing
End Function

Function hlighttxt(txtbox As TextBox)
    txtbox.SelStart = 0
    txtbox.SelLength = Len(txtbox)
End Function

Function checkPath(strPath As String) As Boolean
    If Dir$(strPath) <> vbNullString Then
        checkPath = True
    Else
        checkPath = False
    End If
End Function

Public Function str2MD5(ByVal str As String) As String
    str2MD5 = DigestStrToHexStr(str)
End Function
