VERSION 5.00
Begin VB.Form frmRes 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Res"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2190
   Icon            =   "frmRes.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   159
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   146
   StartUpPosition =   3  'Windows Default
   Tag             =   "patch.ini"
   Begin VB.Image imgReplayA 
      Appearance      =   0  'Flat
      Height          =   180
      Left            =   1320
      Picture         =   "frmRes.frx":058A
      Top             =   240
      Width           =   480
   End
   Begin VB.Image imgReplayB 
      Appearance      =   0  'Flat
      Height          =   180
      Left            =   1320
      Picture         =   "frmRes.frx":0A4C
      Top             =   600
      Width           =   480
   End
   Begin VB.Image imgCheckOFF 
      Appearance      =   0  'Flat
      Height          =   150
      Left            =   600
      Picture         =   "frmRes.frx":0F0E
      Top             =   1920
      Width           =   510
   End
   Begin VB.Image imgCheckON 
      Appearance      =   0  'Flat
      Height          =   150
      Left            =   600
      Picture         =   "frmRes.frx":14B8
      Top             =   1560
      Width           =   510
   End
   Begin VB.Image imgB 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   480
      Picture         =   "frmRes.frx":1A62
      Top             =   600
      Width           =   630
   End
   Begin VB.Image imgA 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   480
      Picture         =   "frmRes.frx":2214
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

