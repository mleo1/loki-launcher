VERSION 5.00
Begin VB.Form frmLauncher 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FF00FF&
   BorderStyle     =   0  'None
   Caption         =   "Loki Launcher"
   ClientHeight    =   7650
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7605
   Icon            =   "frmLauncher.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7650
   ScaleWidth      =   7605
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrCheck 
      Interval        =   100
      Left            =   1080
      Top             =   6120
   End
   Begin VB.TextBox txtUser 
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
      Top             =   5145
      Width           =   1815
   End
   Begin VB.TextBox txtPass 
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
      Top             =   5625
      Width           =   1815
   End
   Begin VB.Image imgReplay 
      Height          =   180
      Left            =   1725
      Picture         =   "frmLauncher.frx":000C
      Top             =   6240
      Width           =   480
   End
   Begin VB.Image imgKeep 
      Height          =   150
      Left            =   5160
      Picture         =   "frmLauncher.frx":04CE
      Top             =   5160
      Width           =   510
   End
   Begin VB.Image imgExit 
      Height          =   300
      Left            =   5160
      Picture         =   "frmLauncher.frx":0A78
      Top             =   6120
      Width           =   630
   End
   Begin VB.Image imgLogin 
      Height          =   300
      Left            =   4440
      Picture         =   "frmLauncher.frx":0F60
      Top             =   6120
      Width           =   630
   End
   Begin VB.Image imgLoginBox 
      Height          =   1800
      Left            =   1680
      Picture         =   "frmLauncher.frx":1448
      Top             =   4680
      Width           =   4200
   End
End
Attribute VB_Name = "frmLauncher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'mleo1

Option Explicit

Private Keep As Boolean
Private MX, MY As Single
Private isRORun As Boolean
Private RO_PID As Long

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then imgLogin_Click
    If KeyCode = vbKeyEscape Then Unload Me
    
    Dim blnIsShift As Boolean
    blnIsShift = Shift And vbAltMask
    If blnIsShift And (KeyCode = vbKeyF4) Then Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
    transparentForm Me

    Me.Caption = INI_SETTINGS.Title
    'Me.Icon = LoadResPicture(101, vbResIcon)
    Me.Icon = frmRes.Icon
If debugMode = 0 Then
    Me.Picture = LoadPicture(BGLoc)
End If
    
    If INI_SETTINGS.Keep = "true" Then
        imgKeep.Picture = frmRes.imgCheckON.Picture
        Keep = True
        
        txtUser = INI_SETTINGS.User
    Else
        imgKeep.Picture = frmRes.imgCheckOFF.Picture
        Keep = False
    End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    moveForm Me
End Sub

Private Sub imgKeep_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
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

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Mode = "f" Then Exit Sub
    imgLogin.Picture = frmRes.imgLoginA.Picture
    imgExit.Picture = frmRes.imgExitA.Picture
    imgReplay.Picture = frmRes.imgReplayA.Picture
    Mode = "f"
End Sub

Private Sub imgLogin_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Mode = "l" Then Exit Sub
    imgLogin.Picture = frmRes.imgLoginB.Picture
imgExit.Picture = frmRes.imgExitA.Picture
    Mode = "l"
End Sub

Private Sub imgLogin_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    imgLogin.Picture = frmRes.imgLoginC.Picture
End Sub

Private Sub imgReplay_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    imgReplay.Picture = frmRes.imgReplayB
End Sub

Private Sub imgReplay_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    imgReplay.Picture = frmRes.imgReplayC
End Sub

Private Sub imgReplay_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    imgReplay.Picture = frmRes.imgReplayA
End Sub

Private Sub imgExit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Mode = "e" Then Exit Sub
    imgExit.Picture = frmRes.imgExitB.Picture
imgLogin.Picture = frmRes.imgLoginA.Picture
    Mode = "e"
End Sub

Private Sub imgExit_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    imgExit.Picture = frmRes.imgExitC.Picture
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
    unloadALL
End Sub

Private Sub tmrCheck_Timer()
    If IsPidRunning(RO_PID) = True Then
        If isRORun = True Then Exit Sub
        isRORun = True
        trayIconAdd Me, INI_SETTINGS.Title & " - " & txtUser
        Me.Hide
    Else
        If isRORun = False Then Exit Sub
        isRORun = False
        trayIconDel Me
        Me.Show
        SetTopMostWindow Me.hWnd, True
        ForceWindowToTop Me.hWnd
        txtPass.SetFocus
    End If
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
    
    'Me.Hide
    Dim strPass As String
    strPass = cMD5(txtPass)
    txtPass = vbNullString
    RO_PID = Shell(ROExeLoc & " " & INI_SETTINGS.EXEArg & " /account:clientinfo.xml -t:" & strPass & " " & txtUser & " server", vbNormalFocus)
    
    Exit Sub
hell:
    MsgBox "Error", vbExclamation, INI_SETTINGS.Title
End Sub

Function cMD5(ByVal pass As String) As String
    If INI_SETTINGS.MD5 = "true" Then
        cMD5 = DigestStrToHexStr(pass)
    Else
        cMD5 = pass
    End If
End Function

Private Sub imgReplay_Click()
    Unload Me
    Shell ROExeLoc & " -Replay", vbNormalFocus
End Sub

Private Sub txtUser_KeyUp(KeyCode As Integer, Shift As Integer)
    If Keep = True Then WriteIni "Loki Launcher", "User", txtUser
End Sub

Private Sub imgLoginBox_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    MX = x
    MY = y
    Mode = "md"
End Sub

Private Sub imgLoginBox_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Mode = "md" Then MoveDrag x, y

    If Mode = "md" Then Exit Sub
    imgLogin.Picture = frmRes.imgLoginA.Picture
    imgExit.Picture = frmRes.imgExitA.Picture
    imgReplay.Picture = frmRes.imgReplayA.Picture
    Mode = "lb"
End Sub

Private Sub imgLoginBox_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Mode = "lb"
End Sub

Public Function MoveDrag(x As Single, y As Single)
    For Each Lokis In frmLauncher
        If Not TypeOf Lokis Is Timer Then Lokis.Move Lokis.Left + (x - MX), Lokis.Top + (y - MY)
    Next
End Function
