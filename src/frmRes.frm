VERSION 5.00
Begin VB.Form frmRes 
   Caption         =   "loki.ini"
   ClientHeight    =   2490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3975
   LinkTopic       =   "Form2"
   ScaleHeight     =   2490
   ScaleWidth      =   3975
   StartUpPosition =   3  'Windows Default
   Begin VB.Image imgCheckOFF 
      Height          =   150
      Left            =   600
      Picture         =   "frmRes.frx":0000
      Top             =   1920
      Width           =   510
   End
   Begin VB.Image imgCheckON 
      Height          =   150
      Left            =   600
      Picture         =   "frmRes.frx":05AA
      Top             =   1560
      Width           =   510
   End
   Begin VB.Image imgExitC 
      Height          =   300
      Left            =   2640
      Picture         =   "frmRes.frx":0B54
      Top             =   960
      Width           =   630
   End
   Begin VB.Image imgExitB 
      Height          =   300
      Left            =   2640
      Picture         =   "frmRes.frx":1306
      Top             =   600
      Width           =   630
   End
   Begin VB.Image imgExitA 
      Height          =   300
      Left            =   2640
      Picture         =   "frmRes.frx":1AB6
      Top             =   240
      Width           =   630
   End
   Begin VB.Image imgLoginC 
      Height          =   300
      Left            =   480
      Picture         =   "frmRes.frx":1F9E
      Top             =   960
      Width           =   630
   End
   Begin VB.Image imgLoginB 
      Height          =   300
      Left            =   480
      Picture         =   "frmRes.frx":274E
      Top             =   600
      Width           =   630
   End
   Begin VB.Image imgLoginA 
      Height          =   300
      Left            =   480
      Picture         =   "frmRes.frx":2EFE
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
