VERSION 5.00
Begin VB.Form frmLauncher 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Loki Launcher"
   ClientHeight    =   7440
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7515
   Icon            =   "frmLauncher.frx":0000
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
      Picture         =   "frmLauncher.frx":058A
      Top             =   4800
      Width           =   510
   End
   Begin VB.Image imgExit 
      Height          =   300
      Left            =   5160
      Picture         =   "frmLauncher.frx":0B34
      Top             =   5760
      Width           =   630
   End
   Begin VB.Image imgLogin 
      Height          =   300
      Left            =   4440
      Picture         =   "frmLauncher.frx":101C
      Top             =   5760
      Width           =   630
   End
   Begin VB.Image imgLoginBox 
      Height          =   1800
      Left            =   1680
      Picture         =   "frmLauncher.frx":1504
      Top             =   4320
      Width           =   4200
   End
End
Attribute VB_Name = "frmLauncher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Keep As Boolean
Private Mode As String
Private MX, MY As Integer

Private Sub Form_Load()
On Error Resume Next
    Me.Caption = INI_SETTINGS.Title
    Me.Picture = LoadPicture(BGLoc)
    'Me.Icon = LoadResPicture(101, vbResIcon)
    
    If INI_SETTINGS.Keep = "true" Then
        imgKeep.Picture = frmRes.imgCheckON.Picture
        Keep = True
        
        txtUser = INI_SETTINGS.User
    Else
        imgKeep.Picture = frmRes.imgCheckOFF.Picture
        Keep = False
    End If
End Sub

Private Sub imgKeep_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Keep Then
        imgKeep.Picture = frmRes.imgCheckOFF.Picture
        Keep = False
        WriteIni "Loki Launcher", "Keep", "false"
        WriteIni "Loki Launcher", "User", vbNullString
    Else
        imgKeep.Picture = frmRes.imgCheckON.Picture
        Keep = True
        WriteIni "Loki Launcher", "Keep", "true"
        WriteIni "Loki Launcher", "User", txtUser
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Mode = "f" Then Exit Sub
    imgLogin.Picture = frmRes.imgLoginA.Picture
    imgExit.Picture = frmRes.imgExitA.Picture
    Mode = "f"
End Sub

Private Sub imgExit_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Mode = "e" Then Exit Sub
    imgExit.Picture = frmRes.imgExitB.Picture
    Mode = "e"
End Sub

Private Sub imgLogin_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Mode = "l" Then Exit Sub
    imgLogin.Picture = frmRes.imgLoginB.Picture
    Mode = "l"
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
    hlightTxt txtUser
End Sub

Private Sub txtPass_GotFocus()
    hlightTxt txtPass
End Sub

Private Sub imgExit_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    unloadAllForms
End Sub

Private Sub imgLogin_Click()
On Error GoTo hell
    If txtUser = vbNullString Then
        MsgBox "Please fill up username", vbInformation, INI_SETTINGS.Title
        Exit Sub
    End If
    If txtPass = vbNullString Then
        MsgBox "Please fill up password", vbInformation, INI_SETTINGS.Title
        Exit Sub
    End If
    
    Shell ROExeLoc & " " & INI_SETTINGS.EXEArg & " -t:" & cMD5(txtPass) & " " & txtUser & " server", vbNormalFocus
    
    Unload Me
    Exit Sub
hell:
    MsgBox "Error", vbExclamation, INI_SETTINGS.Title
End Sub

Function cMD5(ByVal Pass As String)
    If INI_SETTINGS.MD5 = "true" Then
        cMD5 = DigestStrToHexStr(Pass)
    Else
        cMD5 = Pass
    End If
End Function

Private Sub txtUser_KeyUp(KeyCode As Integer, Shift As Integer)
    If Keep = True Then WriteIni "Loki Launcher", "User", txtUser
End Sub

Private Sub imgLoginBox_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    MX = x
    MY = Y
    Mode = "md"
End Sub

Private Sub imgLoginBox_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Mode = "md" Then MoveDrag x, Y

    If Mode = "md" Then Exit Sub
    imgLogin.Picture = frmRes.imgLoginA.Picture
    imgExit.Picture = frmRes.imgExitA.Picture
End Sub

Private Sub imgLoginBox_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Mode = "lb"
End Sub

Sub MoveDrag(x As Single, Y As Single)
    Sleep 10
    Dim lokis As Control
    For Each lokis In frmLauncher
        lokis.Left = lokis.Left + (x - MX)
        lokis.Top = lokis.Top + (Y - MY)
    Next
    Set lokis = Nothing
End Sub

