VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.Ocx"
Begin VB.Form frmPatcher 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FF00FF&
   BorderStyle     =   0  'None
   Caption         =   "Loki Patcher"
   ClientHeight    =   7680
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7620
   Icon            =   "frmPatcher.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   512
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   508
   StartUpPosition =   2  'CenterScreen
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   5760
      Top             =   6600
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.PictureBox picPatch 
      Height          =   735
      Left            =   360
      ScaleHeight     =   675
      ScaleWidth      =   5235
      TabIndex        =   1
      Top             =   6480
      Width           =   5295
      Begin MSComctlLib.ProgressBar prgPatch 
         Height          =   240
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   423
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Label lblPatch 
         Caption         =   "Status: No Updates"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   4815
      End
   End
   Begin SHDocVwCtl.WebBrowser webNotice 
      Height          =   2535
      Left            =   360
      TabIndex        =   0
      Top             =   3840
      Width           =   5295
      ExtentX         =   9340
      ExtentY         =   4471
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Image imgExit 
      Height          =   495
      Left            =   5760
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Image imgStart 
      Height          =   495
      Left            =   5760
      Top             =   3900
      Width           =   1575
   End
End
Attribute VB_Name = "frmPatcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'mleo1

Option Explicit

Private picTemp As PictureBox

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then imgStart_Click
    If KeyCode = vbKeyEscape Then Unload Me
    
    Dim blnIsShift As Boolean
    blnIsShift = Shift And vbAltMask
    If blnIsShift And (KeyCode = vbKeyF4) Then Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = INI_SETTINGS.Title
    'Me.Icon = LoadResPicture(101, vbResIcon)
    Me.Icon = frmRes.Icon
If debugMode = 0 Then
    Me.Picture = LoadPicture(BGLoc)
End If

    Set picTemp = Controls.Add("vb.picturebox", "picTemp")
    picTemp.ScaleMode = 3
    picTemp.AutoRedraw = True
    picTemp.Top = 0
    picTemp.Left = 0
   
    lokiMagic imgStart, frmRes.picStartA
    lokiMagic imgExit, frmRes.picExitA
    
    transparentForm Me
    
    'webNotice.Navigate "http://xvideos.com/notice.html"
End Sub

Function lokiMagic(ByRef imgTO As Image, ByRef picFROM As PictureBox)
    picFROM.BorderStyle = 1
    picTemp.Width = picFROM.Width
    picTemp.Height = picFROM.Height
    picTemp.ScaleHeight = picFROM.ScaleHeight
    picTemp.ScaleWidth = picFROM.ScaleWidth
    
    picTemp.Picture = picFROM.Picture
    
    TransparentBlt _
        Me.hDC, _
        imgTO.Left, imgTO.Top, _
        picTemp.ScaleWidth, _
        picTemp.ScaleHeight, _
        picTemp.hDC, _
        0, 0, _
        picTemp.ScaleWidth, _
        picTemp.ScaleHeight, _
        RGB(255, 0, 255)
        
    imgTO.Height = picTemp.Height
    imgTO.Width = picTemp.Width
End Function

Private Function patch()
    'debug
       LokiLoc = App.Path & "\"
        
    '

    
        '
        Dim DLLoc, DLFile As String
        Dim str1() As String
        DLLoc = "http://tae.com/php.ini"
        str1 = Split(DLLoc, "/")
        DLFile = str1(UBound(str1))
        lblPatch = "Status: Downloading " & DLFile
        downloadFile DLLoc, LokiLoc & DLFile
    
    '
    prgPatch.Value = 0

    lblPatch = "Status: Finished"
End Function

Private Function downloadFile(ByVal strURL As String, ByVal strDestination As String)
    Const CHUNK_SIZE As Long = 1024
    Dim intFile As Integer
    Dim lngBytesReceived As Long
    Dim lngFileLength As Long
    Dim strHeader As String
    Dim b() As Byte
    Dim i As Integer
    
    prgPatch.Value = 0
    
    DoEvents
        
    With Inet1
        
        .URL = strURL
        .Execute , "GET", , "Range: bytes=" & CStr(lngBytesReceived) & "-" & vbCrLf
                
        While .StillExecuting
            DoEvents
        Wend
        
        strHeader = .GetHeader
    End With
        
        
    strHeader = Inet1.GetHeader("Content-Length")
    lngFileLength = Val(strHeader)
    
    DoEvents
        
    lngBytesReceived = 0
    
    intFile = FreeFile()
    
    Open strDestination For Binary Access Write As #intFile
    
    Do
        b = Inet1.GetChunk(CHUNK_SIZE, icByteArray)
        Put #intFile, , b
        lngBytesReceived = lngBytesReceived + UBound(b, 1) + 1
        
        prgPatch.Value = (Round((lngBytesReceived / lngFileLength) * 100))
        DoEvents
    Loop While UBound(b, 1) > 0

    Close #intFile
End Function

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    moveForm Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set picTemp = Nothing
    
    unloadALL
End Sub

Private Sub imgExit_Click()
    Unload Me
End Sub

Private Sub imgStart_Click()
    Exit Sub
    Unload Me
    If INI_SETTINGS.Mode = "1" Then frmLauncher.Show
    If INI_SETTINGS.Mode = "3" Then Shell ROExeLoc
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Mode = "f" Then Exit Sub
    lokiMagic imgStart, frmRes.picStartA
    lokiMagic imgExit, frmRes.picExitA
    Mode = "f"
End Sub

Private Sub imgExit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Mode = "e" Then Exit Sub
        lokiMagic imgExit, frmRes.picExitB
lokiMagic imgStart, frmRes.picStartA
    Mode = "e"
End Sub

Private Sub imgStart_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Mode = "s" Then Exit Sub
        lokiMagic imgStart, frmRes.picStartB
lokiMagic imgExit, frmRes.picExitA
    Mode = "s"
End Sub

Private Sub imgExit_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    lokiMagic imgExit, frmRes.picExitC
End Sub

Private Sub imgStart_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    lokiMagic imgStart, frmRes.picStartC
End Sub
