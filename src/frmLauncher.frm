VERSION 5.00
Begin VB.Form frmLauncher 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Loki Launcher"
   ClientHeight    =   4950
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5880
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawMode        =   16  'Merge Pen
   Icon            =   "frmLauncher.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   5880
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
      Left            =   2235
      MaxLength       =   26
      TabIndex        =   1
      Top             =   1785
      Width           =   1860
   End
   Begin VB.Timer tmrCheck 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   840
      Top             =   3360
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
      Left            =   2235
      MaxLength       =   26
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   2265
      Width           =   1860
   End
   Begin VB.Label lblLogin 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "login"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   2280
      TabIndex        =   7
      Top             =   2775
      Width           =   405
   End
   Begin VB.Label lblExit 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   3135
      TabIndex        =   8
      Top             =   2775
      Width           =   270
   End
   Begin VB.Label lblX 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   4800
      TabIndex        =   9
      Top             =   2880
      Width           =   135
   End
   Begin VB.Label lblReplay 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "replay"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   960
      TabIndex        =   6
      Top             =   2865
      Width           =   390
   End
   Begin VB.Label lblKeep 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "keep"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   4425
      TabIndex        =   5
      Top             =   1800
      Width           =   315
   End
   Begin VB.Label lblLogON 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "LogOn"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   1320
      Width           =   3855
   End
   Begin VB.Label lblPassword 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   960
      TabIndex        =   3
      Top             =   2235
      Width           =   1095
   End
   Begin VB.Label lblID 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ID"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   960
      TabIndex        =   2
      Top             =   1755
      Width           =   1095
   End
   Begin VB.Image imgReplay 
      Appearance      =   0  'Flat
      Height          =   195
      Left            =   885
      Picture         =   "frmLauncher.frx":000C
      Stretch         =   -1  'True
      Top             =   2865
      Width           =   480
   End
   Begin VB.Image imgKeep 
      Appearance      =   0  'Flat
      Height          =   150
      Left            =   4230
      Picture         =   "frmLauncher.frx":04CE
      Top             =   1800
      Width           =   510
   End
   Begin VB.Image imgExit 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   3000
      Picture         =   "frmLauncher.frx":0A78
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   615
   End
   Begin VB.Image imgLogin 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   2160
      Picture         =   "frmLauncher.frx":122A
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   615
   End
   Begin VB.Image imgLoginWindow 
      Appearance      =   0  'Flat
      Height          =   1800
      Left            =   840
      Picture         =   "frmLauncher.frx":19DC
      Top             =   1320
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

Private lLeft(16) As Integer
Private lTop(16) As Integer
Private fHeight, fWidth As Integer
Private Const BorderLimit As Integer = 10
Private Const BMove As Integer = 100
Private DisableMoveW As Boolean

Private Sub Form_Activate()
    If txtUser.Text = vbNullString Then
        txtUser.SetFocus
    Else
        txtPass.SetFocus
    End If
End Sub

Private Sub Form_Load()
    Me.Caption = App.EXEName
    SetIcon Me.hWnd, 1, True

    Me.BackColor = RGB(255, 0, 255)

    lblID.ForeColor = RGB(90, 111, 153)
    lblPassword.ForeColor = RGB(90, 111, 153)
    txtUser.BackColor = RGB(242, 242, 242)
    txtPass.BackColor = RGB(242, 242, 242)
    lblKeep.ForeColor = RGB(105, 105, 105)
   
    lblLogON.Caption = Left(LoadResString(101), 24)
    lblID.Caption = Left(LoadResString(102), 10)
    lblPassword.Caption = Left(LoadResString(103), 10)
    lblLogin.Caption = Left(LoadResString(104), 12)
    lblExit.Caption = Left(LoadResString(105), 12)
    lblReplay.Caption = Left(LoadResString(106), 12)
    lblKeep.Caption = Left(LoadResString(107), 7)

    lblLogin.FontSize = LFontSize(lblLogin)
    lblExit.FontSize = LFontSize(lblExit)
    imgReplay.Width = lblReplay.Width + 120
    imgLogin.Width = lblLogin.Width + 220
    imgExit.Width = lblExit.Width + 260
    lblX.BackStyle = 0
    lblExit.Left = lblX.Left - lblExit.Width + 60
    imgExit.Left = lblExit.Left - 120
    lblLogin.Left = imgExit.Left - lblLogin.Width - 160
    imgLogin.Left = lblLogin.Left - 120
    
    If LSETTINGS.Keep = True Then
        imgKeep.Picture = frmRes.imgCheckON.Picture
        txtUser = LSETTINGS.user
    Else
        imgKeep.Picture = frmRes.imgCheckOFF.Picture
    End If
 
    Dim sPic As StdPicture
    Set sPic = LoadResPicture("Skin", vbResBitmap)
    If ScaleX(sPic.Width, vbHimetric, vbPixels) < 279 Or _
        ScaleY(sPic.Height, vbHimetric, vbPixels) < 119 Then
        Unload Me
        MsgBox "Please change the skin to atleast 280x120", vbInformation, App.EXEName
        Exit Sub
    End If
    
    Dim X As Integer
    fHeight = Me.Height
    fWidth = Me.Width
    For Each Lokis In Me
        If Not TypeOf Lokis Is Timer Then
            lLeft(X) = Lokis.Left
            lTop(X) = Lokis.Top
            X = X + 1
        End If
    Next

    Me.Width = ScaleX(sPic.Width, vbHimetric, vbTwips)
    Me.Height = ScaleY(sPic.Height, vbHimetric, vbTwips)
    Me.Picture = sPic
    
    X = 0
    If ScaleX(sPic.Width, vbHimetric, vbPixels) < 300 Or _
        ScaleY(sPic.Height, vbHimetric, vbPixels) < 300 Then

        For Each Lokis In Me
            If Not TypeOf Lokis Is Timer Then
                Lokis.Left = lLeft(X) - (fWidth - Me.Width) / 2
                Lokis.Top = (lTop(X) - (fHeight - Me.Height) / 2) + 180
                X = X + 1
            End If
        Next
        DisableMoveW = True
    Else
        For Each Lokis In Me
            If Not TypeOf Lokis Is Timer Then
                Lokis.Left = lLeft(X) - (fWidth - Me.Width) / 2
                Lokis.Top = lTop(X) - (fHeight - Me.Height)
                X = X + 1
            End If
        Next
    End If
    
    TransparentForm Me
    
    EXE_PID = -1
    tmrCheck.Enabled = True
End Sub

Function LFontSize(lbl As Label) As Integer
     Select Case Len(lbl.Caption)
        Case Is >= 9
            LFontSize = 7
            lbl.Top = lbl.Top + 30
        Case Is >= 6
            LFontSize = 8
            lbl.Top = lbl.Top + 20
        Case Is >= 0
            LFontSize = 9
    End Select
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then imgLogin_Click
    If KeyCode = vbKeyEscape Then Unload Me
    If KeyCode = vbKeyF1 Then imgReplay_Click
    If KeyCode = vbKeyF5 Then
        Dim SetupPath As String
        SetupPath = LokiDir & "setup.exe"
        If CheckFile(SetupPath) = True Then
            Me.Hide
            EXE_PID = Shell(SetupPath, vbNormalFocus)
        Else
            MsgBox "Can't find " & SetupPath, vbCritical, App.EXEName
        End If
    End If
    
    Dim xShift As Boolean
    xShift = Shift And vbAltMask
    If xShift And KeyCode = vbKeyF4 Then Unload Me
    
    xShift = Shift And vbCtrlMask
    If xShift And KeyCode = vbKey1 Then imgKeep_Click
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MoveForm Me
End Sub

Private Sub lblKeep_Click()
    imgKeep_Click
End Sub

Private Sub imgKeep_Click()
    If LSETTINGS.Keep = True Then
        imgKeep.Picture = frmRes.imgCheckOFF.Picture
        LSETTINGS.Keep = False
    Else
        imgKeep.Picture = frmRes.imgCheckON.Picture
        LSETTINGS.Keep = True
    End If
    
    LSETTINGS.user = txtUser
    WriteIni "Loki Launcher", "User", LSETTINGS.user
    WriteIni "Loki Launcher", "Keep", IIf(LSETTINGS.Keep, "true", "false")
End Sub

Private Sub txtUser_KeyUp(KeyCode As Integer, Shift As Integer)
    LSETTINGS.user = txtUser
    If LSETTINGS.Keep = True Then WriteIni "Loki Launcher", "User", LSETTINGS.user
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Mode = "f" Then Exit Sub
    imgLogin.Picture = frmRes.imgA.Picture
    imgExit.Picture = frmRes.imgA.Picture
    imgReplay.Picture = frmRes.imgReplayA.Picture
    Mode = "f"
End Sub

'EXIT BUTTON
Private Sub lblExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Mode = "e" Then Exit Sub
    imgExit.Picture = frmRes.imgB.Picture
imgLogin.Picture = frmRes.imgA.Picture
    Mode = "e"
End Sub

Private Sub lblExit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgExit.Picture = frmRes.imgB.Picture
    lblExit.Move lblExit.Left + 20, lblExit.Top + 20
End Sub

Private Sub lblExit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgExit.Picture = frmRes.imgA.Picture
    lblExit.Move lblExit.Left - 20, lblExit.Top - 20
End Sub

Private Sub imgExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Mode = "e" Then Exit Sub
    imgExit.Picture = frmRes.imgB.Picture
imgLogin.Picture = frmRes.imgA.Picture
    Mode = "e"
End Sub

Private Sub imgExit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgExit.Picture = frmRes.imgB.Picture
    lblExit.Move lblExit.Left + 20, lblExit.Top + 20
End Sub

Private Sub imgExit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgExit.Picture = frmRes.imgA.Picture
    lblExit.Move lblExit.Left - 20, lblExit.Top - 20
End Sub

'LOGIN BUTTON
Private Sub lblLogin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Mode = "l" Then Exit Sub
    imgLogin.Picture = frmRes.imgB.Picture
imgExit.Picture = frmRes.imgA.Picture
    Mode = "l"
End Sub

Private Sub lblLogin_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgLogin.Picture = frmRes.imgB.Picture
    lblLogin.Move lblLogin.Left + 20, lblLogin.Top + 20
End Sub

Private Sub lblLogin_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgLogin.Picture = frmRes.imgA.Picture
    lblLogin.Move lblLogin.Left - 20, lblLogin.Top - 20
End Sub

Private Sub imgLogin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Mode = "l" Then Exit Sub
    imgLogin.Picture = frmRes.imgB.Picture
imgExit.Picture = frmRes.imgA.Picture
    Mode = "l"
End Sub

Private Sub imgLogin_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgLogin.Picture = frmRes.imgB.Picture
    lblLogin.Move lblLogin.Left + 20, lblLogin.Top + 20
End Sub

Private Sub imgLogin_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgLogin.Picture = frmRes.imgA.Picture
    lblLogin.Move lblLogin.Left - 20, lblLogin.Top - 20
End Sub

Private Sub lblLogin_Click()
    imgLogin_Click
End Sub

Private Sub imgLogin_Click()
    If CheckFile(ROExePath) = True Then
        If txtUser = vbNullString Or txtPass = vbNullString Then
            MsgBox "Check fields", vbInformation, App.EXEName
            Exit Sub
        End If
    
        Dim strPass As String
        strPass = cPass(txtPass)
        txtPass = vbNullString
        
        Me.Hide
        EXE_PID = Shell(ROExePath & " " & LSETTINGS.AddArg & " -t:" & strPass & " " & txtUser & " server", vbNormalFocus)
    Else
        MsgBox "Can't find " & ROExePath, vbCritical, App.EXEName
    End If
End Sub

'REPLAY BUTTON
Private Sub lblReplay_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Mode = "r" Then Exit Sub
    imgReplay.Picture = frmRes.imgReplayB
    Mode = "r"
End Sub

Private Sub lblReplay_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgReplay.Picture = frmRes.imgReplayB
    lblReplay.Move lblReplay.Left + 20, lblReplay.Top + 20
End Sub

Private Sub lblReplay_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgReplay.Picture = frmRes.imgReplayA
    lblReplay.Move lblReplay.Left - 20, lblReplay.Top - 20
End Sub

Private Sub imgReplay_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Mode = "r" Then Exit Sub
    imgReplay.Picture = frmRes.imgReplayB
    Mode = "r"
End Sub

Private Sub imgReplay_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgReplay.Picture = frmRes.imgReplayA
    lblReplay.Move lblReplay.Left - 20, lblReplay.Top - 20
End Sub

Private Sub imgReplay_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgReplay.Picture = frmRes.imgReplayB
    lblReplay.Move lblReplay.Left + 20, lblReplay.Top + 20
End Sub

Private Sub lblReplay_Click()
    imgReplay_Click
End Sub

Private Sub imgReplay_Click()
    If CheckFile(ROExePath) = True Then
        Me.Hide
        EXE_PID = Shell(ROExePath & " -Replay", vbNormalFocus)
    Else
        MsgBox "Can't find " & ROExePath, vbCritical, App.EXEName
    End If
End Sub

Private Sub txtUser_GotFocus()
    HLightTxt txtUser
End Sub

Private Sub txtPass_GotFocus()
    HLightTxt txtPass
End Sub

Private Sub lblExit_Click()
    imgExit_Click
End Sub

Private Sub imgExit_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnloadAll
End Sub

Private Sub tmrCheck_Timer()
    If DisableMoveW = False Then
        If imgLoginWindow.Left < BorderLimit Then
            For Each Lokis In Me
                If Not TypeOf Lokis Is Timer Then Lokis.Left = Lokis.Left + BMove
            Next
        End If
        If imgLoginWindow.Left + imgLoginWindow.Width > Me.Width - BorderLimit Then
            For Each Lokis In Me
                If Not TypeOf Lokis Is Timer Then Lokis.Left = Lokis.Left - BMove
            Next
        End If
    
        If imgLoginWindow.Top < BorderLimit Then
            For Each Lokis In Me
                If Not TypeOf Lokis Is Timer Then Lokis.Top = Lokis.Top + BMove
            Next
        End If
        If imgLoginWindow.Top + imgLoginWindow.Height > Me.Height - BorderLimit Then
            For Each Lokis In Me
                If Not TypeOf Lokis Is Timer Then Lokis.Top = Lokis.Top - BMove
            Next
        End If
    End If
    
    If isPIDRunning(EXE_PID) = True Then
        TrayIconAdd frmRes, App.EXEName '& " - " & txtUser
    Else
        If Me.Visible = False Then
            EXE_PID = -1
            TrayIconDel frmRes
            Me.Show
        End If
    End If
End Sub

Private Sub imgLoginWindow_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MX = X
    MY = Y
    Mode = "md"
End Sub

Private Sub imgLoginWindow_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Mode = ""
End Sub

Private Sub imgLoginWindow_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Mode = "md" Then MoveWindow X, Y
    
    If Mode = "md" Then Exit Sub
    imgLogin.Picture = frmRes.imgA.Picture
    imgReplay.Picture = frmRes.imgReplayA.Picture
    imgExit.Picture = frmRes.imgA.Picture
    Mode = "lw"
End Sub

Private Sub lblID_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MX = X
    MY = Y
    Mode = "md"
End Sub

Private Sub lblID_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Mode = "md" Then MoveWindow X, Y
End Sub

Private Sub lblID_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Mode = ""
End Sub

Private Sub lblPassword_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MX = X
    MY = Y
    Mode = "md"
End Sub

Private Sub lblPassword_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Mode = "md" Then MoveWindow X, Y
End Sub

Private Sub lblPassword_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Mode = ""
End Sub

Private Sub lblLogON_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MX = X
    MY = Y
    Mode = "md"
End Sub

Private Sub lblLogON_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Mode = "md" Then MoveWindow X, Y
End Sub

Private Sub lblLogON_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Mode = ""
End Sub

Sub MoveWindow(X As Single, Y As Single)
    If isLimit = False Then
        For Each Lokis In Me
            If Not TypeOf Lokis Is Timer Then Lokis.Move Lokis.Left + (X - MX), Lokis.Top + (Y - MY)
        Next
    Else
        Mode = "mf"
        MoveForm Me
    End If
End Sub

Function isLimit() As Boolean
    If imgLoginWindow.Left < BorderLimit Or _
        imgLoginWindow.Left + imgLoginWindow.Width > Me.Width - BorderLimit Or _
        imgLoginWindow.Top < BorderLimit Or _
        imgLoginWindow.Top + imgLoginWindow.Height > Me.Height - BorderLimit Then
        isLimit = True
    Else
        isLimit = False
    End If
End Function

