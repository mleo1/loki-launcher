VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Loki Launcher"
   ClientHeight    =   7440
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7515
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   7515
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUser 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Left            =   3075
      MaxLength       =   23
      TabIndex        =   0
      Top             =   4785
      Width           =   1815
   End
   Begin VB.TextBox txtPass 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      IMEMode         =   3  'DISABLE
      Left            =   3075
      MaxLength       =   23
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   5265
      Width           =   1815
   End
   Begin VB.Image imgKeep 
      Height          =   150
      Left            =   5160
      Picture         =   "frmMain.frx":058A
      Top             =   4800
      Width           =   510
   End
   Begin VB.Image imgExit 
      Height          =   300
      Left            =   5160
      Picture         =   "frmMain.frx":0B34
      Top             =   5760
      Width           =   630
   End
   Begin VB.Image imgLogin 
      Height          =   300
      Left            =   4440
      Picture         =   "frmMain.frx":101C
      Top             =   5760
      Width           =   630
   End
   Begin VB.Image imgLoginBox 
      Height          =   1800
      Left            =   1680
      Picture         =   "frmMain.frx":1504
      Top             =   4320
      Width           =   4200
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Keep As Boolean
Private ROEXE As String
Private Const LOKI As String = "Loki Launcher"
Private Lvar As String

Private Sub imgKeep_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Keep Then
        imgKeep.Picture = frmRes.imgCheckOFF.Picture
        Keep = False
        WriteIni "Loki", "Keep", "false"
        WriteIni "Loki", "User", vbNullString
    Else
        imgKeep.Picture = frmRes.imgCheckON.Picture
        Keep = True
        WriteIni "Loki", "Keep", "true"
        WriteIni "Loki", "User", txtUser
    End If
End Sub

Private Sub imgLoginBox_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Lvar = "lb" Then Exit Sub
    imgLogin.Picture = frmRes.imgLoginA.Picture
    imgExit.Picture = frmRes.imgExitA.Picture
    Lvar = "lb"
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Lvar = "f" Then Exit Sub
    imgLogin.Picture = frmRes.imgLoginA.Picture
    imgExit.Picture = frmRes.imgExitA.Picture
    Lvar = "f"
End Sub

Private Sub imgExit_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Lvar = "e" Then Exit Sub
    imgExit.Picture = frmRes.imgExitB.Picture
    Lvar = "e"
End Sub

Private Sub imgLogin_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Lvar = "l" Then Exit Sub
    imgLogin.Picture = frmRes.imgLoginB.Picture
    Lvar = "l"
End Sub

Private Sub imgExit_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
     imgExit.Picture = frmRes.imgExitC.Picture
End Sub

Private Sub imgLogin_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    imgLogin.Picture = frmRes.imgLoginC.Picture
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then imgLogin_Click
End Sub

Private Sub txtUser_GotFocus()
    hlighttxt txtUser
End Sub

Private Sub txtPass_GotFocus()
    hlighttxt txtPass
End Sub

Private Sub imgExit_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    unloadAllForms
End Sub

Function isFilled() As String
    If txtUser = vbNullString Then
        MsgBox "Please fill up username", vbInformation, LOKI
        isFilled = 0
        Exit Function
    End If
    
    If txtPass = vbNullString Then
        MsgBox "Please fill up password", vbInformation, LOKI
        isFilled = 0
        Exit Function
    End If
    
    isFilled = 1
End Function

Private Sub imgLogin_Click()
On Error GoTo hell
    If isFilled = 0 Then Exit Sub
    
    Dim str1 As String
    
    str1 = App.Path & "\" & ROEXE & " " & ReadIni("Settings", "ExeArg") & " -t:" & checkMD5(txtPass) & " " & txtUser & " server"
    
    Shell str1, vbNormalFocus
    
    Unload Me
    
    Exit Sub
hell:
    MsgBox "Error", vbExclamation, LOKI
End Sub

Function checkMD5(ByVal pass As String)
    If UCase(ReadIni("Settings", "MD5")) = "true" Then
        checkMD5 = str2MD5(pass)
    Else
        checkMD5 = pass
    End If
End Function

Private Sub Form_Load()
    On Error Resume Next
    
    INI_NAME = frmRes.Caption
   
    ROEXE = ReadIni("Settings", "Exe")
    
    If checkPath(App.Path & "\" & ROEXE) = False Then
        MsgBox "Can't find " & App.Path & "\" & ROEXE, vbCritical, LOKI
        Unload Me
        Exit Sub
    End If
    
    If LCase(ReadIni("Loki", "Keep")) = "true" Then
        imgKeep.Picture = frmRes.imgCheckON.Picture
        Keep = True
        
        txtUser = ReadIni("Loki", "User")
    Else
        imgKeep.Picture = frmRes.imgCheckOFF.Picture
        Keep = False
    End If
        
    frmMain.Picture = LoadPicture(App.Path & ReadIni("Settings", "BG"))
End Sub

Private Sub txtUser_KeyUp(KeyCode As Integer, Shift As Integer)
    If Keep = True Then WriteIni "Loki", "User", txtUser
End Sub
