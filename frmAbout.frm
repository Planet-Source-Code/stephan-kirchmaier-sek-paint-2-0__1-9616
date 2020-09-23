VERSION 5.00
Begin VB.Form frmAbout 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "About SEK - Paint 2.0"
   ClientHeight    =   2475
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   3360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   165
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   224
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   2040
      Width           =   3135
   End
   Begin VB.Timer timScroll 
      Interval        =   30
      Left            =   3360
      Top             =   480
   End
   Begin VB.TextBox txtScroll 
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
   Begin VB.Image img1 
      Height          =   1440
      Left            =   600
      Picture         =   "frmAbout.frx":0000
      Stretch         =   -1  'True
      Top             =   480
      Width           =   2160
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private sTextScr As String

Private Sub cmdBack_Click()
    frmAbout.Hide
End Sub

Private Sub Form_Load()
    sTextScr = "          !SEK - Paint 2.0, Made by Stephan Kirchmaier in Y2K! Vote FOR it, please!! THANX!          "
    For i = 0 To frmAbout.ScaleHeight * 2
        Line (i, 0)-(i, frmAbout.ScaleWidth * 2), i
    Next i
End Sub

Private Sub timScroll_Timer()
    If Len(sTextScr) = 0 Then
        sTextScr = "          !SEK - Paint 2.0, Made by Stephan Kirchmaier in Y2K! Vote FOR it, please!! THANX!          "
    End If
    sTextScr = Right(sTextScr, Len(sTextScr) - 1)
    txtScroll.Text = sTextScr
End Sub
