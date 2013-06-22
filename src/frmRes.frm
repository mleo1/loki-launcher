VERSION 5.00
Begin VB.Form frmRes 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FF00FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmRes"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4410
   Icon            =   "frmRes.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   288
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   294
   StartUpPosition =   3  'Windows Default
   Tag             =   "patch.ini"
   Begin VB.PictureBox picStartC 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   480
      Picture         =   "frmRes.frx":058A
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   107
      TabIndex        =   5
      Top             =   3600
      Width           =   1605
   End
   Begin VB.PictureBox picExitC 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   2160
      Picture         =   "frmRes.frx":2D0A
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   107
      TabIndex        =   4
      Top             =   3600
      Width           =   1605
   End
   Begin VB.PictureBox picStartB 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   480
      Picture         =   "frmRes.frx":548A
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   107
      TabIndex        =   3
      Top             =   3000
      Width           =   1605
   End
   Begin VB.PictureBox picExitB 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   2160
      Picture         =   "frmRes.frx":7C0A
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   107
      TabIndex        =   2
      Top             =   3000
      Width           =   1605
   End
   Begin VB.PictureBox picStartA 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   480
      Picture         =   "frmRes.frx":A38A
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   107
      TabIndex        =   1
      Top             =   2400
      Width           =   1605
   End
   Begin VB.PictureBox picExitA 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   2160
      Picture         =   "frmRes.frx":CB0A
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   107
      TabIndex        =   0
      Top             =   2400
      Width           =   1605
   End
   Begin VB.Image imgReplayA 
      Appearance      =   0  'Flat
      Height          =   180
      Left            =   2520
      Picture         =   "frmRes.frx":F28A
      Top             =   240
      Width           =   480
   End
   Begin VB.Image imgReplayB 
      Appearance      =   0  'Flat
      Height          =   180
      Left            =   2520
      Picture         =   "frmRes.frx":F74C
      Top             =   600
      Width           =   480
   End
   Begin VB.Image imgReplayC 
      Appearance      =   0  'Flat
      Height          =   180
      Left            =   2520
      Picture         =   "frmRes.frx":FC0E
      Top             =   960
      Width           =   480
   End
   Begin VB.Image imgCheckOFF 
      Appearance      =   0  'Flat
      Height          =   150
      Left            =   600
      Picture         =   "frmRes.frx":100D0
      Top             =   1920
      Width           =   510
   End
   Begin VB.Image imgCheckON 
      Appearance      =   0  'Flat
      Height          =   150
      Left            =   600
      Picture         =   "frmRes.frx":1067A
      Top             =   1560
      Width           =   510
   End
   Begin VB.Image imgExitC 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1440
      Picture         =   "frmRes.frx":10C24
      Top             =   960
      Width           =   630
   End
   Begin VB.Image imgExitB 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1440
      Picture         =   "frmRes.frx":113D6
      Top             =   600
      Width           =   630
   End
   Begin VB.Image imgExitA 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1440
      Picture         =   "frmRes.frx":11B86
      Top             =   240
      Width           =   630
   End
   Begin VB.Image imgLoginC 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   480
      Picture         =   "frmRes.frx":1206E
      Top             =   960
      Width           =   630
   End
   Begin VB.Image imgLoginB 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   480
      Picture         =   "frmRes.frx":1281E
      Top             =   600
      Width           =   630
   End
   Begin VB.Image imgLoginA 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   480
      Picture         =   "frmRes.frx":12FCE
      Top             =   240
      Width           =   630
   End
End
Attribute VB_Name = "frmRes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Me.Hide
End Sub

