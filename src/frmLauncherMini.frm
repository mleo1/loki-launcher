VERSION 5.00
Begin VB.Form frmLauncherMini 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "LogOn"
   ClientHeight    =   1710
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1010.324
   ScaleMode       =   0  'User
   ScaleWidth      =   4394.266
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPass 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1440
      MaxLength       =   26
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   600
      Width           =   1860
   End
   Begin VB.TextBox txtUser 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1440
      MaxLength       =   26
      TabIndex        =   1
      Top             =   120
      Width           =   1860
   End
   Begin VB.CommandButton cmdReplay 
      Caption         =   "replay"
      Height          =   305
      Left            =   120
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1339
      Width           =   735
   End
   Begin VB.CheckBox chkAuto 
      Caption         =   "runbg"
      Height          =   255
      Left            =   3480
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   457
      Width           =   975
   End
   Begin VB.CheckBox chkKeep 
      Caption         =   "keep"
      Height          =   255
      Left            =   3480
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   171
      Width           =   975
   End
   Begin VB.Timer tmrCheck 
      Interval        =   40
      Left            =   1560
      Top             =   1200
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "login"
      Default         =   -1  'True
      Height          =   457
      Left            =   2280
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1200
      Width           =   1065
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "exit"
      Height          =   457
      Left            =   3480
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1200
      Width           =   1065
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      X1              =   112.673
      X2              =   4281.593
      Y1              =   638.099
      Y2              =   638.099
   End
   Begin VB.Label lblID 
      Alignment       =   1  'Right Justify
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
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   165
      Width           =   1095
   End
   Begin VB.Label lblPassword 
      Alignment       =   1  'Right Justify
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
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   645
      Width           =   1095
   End
End
Attribute VB_Name = "frmLauncherMini"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'mleo1

Option Explicit

Private blnHasLoaded As Boolean

Private Sub Form_Load()
    SetIcon Me.hWnd, "MAIN", True

    Me.Caption = App.EXEName
    lblID.Caption = Left(LoadResString(102), 10)
    lblPassword.Caption = Left(LoadResString(103), 10)
    cmdLogin.Caption = Left(LoadResString(104), 12)
    cmdExit.Caption = Left(LoadResString(105), 12)
    cmdReplay.Caption = Left(LoadResString(106), 12)
    chkKeep.Caption = Left(LoadResString(107), 7)
    chkAuto.Caption = Left(LoadResString(108), 7)
    
    If INI_SETTINGS.Keep = "true" Then
        chkKeep.Value = 1
        txtUser = INI_SETTINGS.User
    Else
        chkKeep.Value = 0
    End If
    If INI_SETTINGS.Auto = "true" Then
        chkAuto.Value = 1
    Else
        chkAuto.Value = 0
    End If
    
    blnHasLoaded = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cmdLogin_Click
    If KeyCode = vbKeyEscape Then Unload Me
    If KeyCode = vbKeyF1 Then cmdReplay_Click
    If KeyCode = vbKeyF5 Then
        Dim SetupPath As String
        SetupPath = LokiDir & "setup.exe"
        If CheckFile(SetupPath) = True Then
            If INI_SETTINGS.Auto = "true" Then
                Me.Hide
                EXE_PID = Shell(SetupPath, vbNormalFocus)
            Else
                Unload Me
                Shell SetupPath, vbNormalFocus
            End If
        Else
            MsgBox "Can't find " & SetupPath, vbCritical, App.EXEName
        End If
    End If
    
    Dim xShift As Boolean
    xShift = Shift And vbAltMask
    If xShift And KeyCode = vbKeyF4 Then Unload Me
    
    xShift = Shift And vbCtrlMask
    If xShift And KeyCode = vbKey1 Then chkKeep_Click
    If xShift And KeyCode = vbKey2 Then chkAuto_Click
End Sub

Private Sub chkAuto_Click()
    If blnHasLoaded = False Then Exit Sub
    
    If INI_SETTINGS.Auto = "true" Then
        chkAuto.Value = 0
        INI_SETTINGS.Auto = "false"
    Else
        chkAuto.Value = 1
        INI_SETTINGS.Auto = "true"
    End If
    WriteIni "Loki Launcher", "Auto", INI_SETTINGS.Auto
End Sub

Private Sub chkKeep_Click()
    If blnHasLoaded = False Then Exit Sub

    If INI_SETTINGS.Keep = "true" Then
        chkKeep.Value = 0
        INI_SETTINGS.User = vbNullString
        INI_SETTINGS.Keep = "false"
    Else
        chkKeep.Value = 1
        INI_SETTINGS.User = txtUser
        INI_SETTINGS.Keep = "true"
    End If
    WriteIni "Loki Launcher", "User", INI_SETTINGS.User
    WriteIni "Loki Launcher", "Keep", INI_SETTINGS.Keep
End Sub

Private Sub cmdLogin_Click()
    If CheckFile(ROExePath) = True Then
        If txtUser = vbNullString Or txtPass = vbNullString Then
            MsgBox "Check fields", vbInformation, App.EXEName
            Exit Sub
        End If
    
        Dim strPass As String
        strPass = cMD5(txtPass)
        txtPass = vbNullString
        
        If INI_SETTINGS.Auto = "true" Then
            Me.Hide
            EXE_PID = Shell(ROExePath & " " & INI_SETTINGS.EXEArg & " /account:clientinfo.xml -t:" & strPass & " " & txtUser & " server", vbNormalFocus)
        Else
            Unload Me
            Shell ROExePath & " " & INI_SETTINGS.EXEArg & " /account:clientinfo.xml -t:" & strPass & " " & txtUser & " server", vbNormalFocus
        End If
    Else
        MsgBox "Can't find " & ROExePath, vbCritical, App.EXEName
    End If
End Sub

Private Sub cmdReplay_Click()
    If CheckFile(ROExePath) = True Then
        If INI_SETTINGS.Auto = "true" Then
            Me.Hide
            EXE_PID = Shell(ROExePath & " " & INI_SETTINGS.EXEArg & " -Replay", vbNormalFocus)
        Else
            Unload Me
            Shell ROExePath & " " & INI_SETTINGS.EXEArg & " -Replay", vbNormalFocus
        End If
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

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnloadAll
End Sub

Private Sub tmrCheck_Timer()
    If isPIDRunning(EXE_PID) = True Then
        TrayIconAdd frmRes, App.EXEName
    Else
        If Me.Visible = False Then
            TrayIconDel frmRes
            Me.Show
        End If
    End If
End Sub

Private Sub txtUser_KeyUp(KeyCode As Integer, Shift As Integer)
    INI_SETTINGS.User = txtUser
    If INI_SETTINGS.Keep = "true" Then WriteIni "Loki Launcher", "User", INI_SETTINGS.User
End Sub

