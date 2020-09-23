VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRotate 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Rotate the Picture"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   5265
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   5265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin MSComctlLib.ProgressBar ctlRotProgBar 
      Height          =   135
      Left            =   2040
      TabIndex        =   5
      Top             =   2520
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.PictureBox picPrev 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   2295
      Left            =   2040
      ScaleHeight     =   149
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   197
      TabIndex        =   4
      Top             =   120
      Width           =   3015
   End
   Begin VB.ComboBox cboAngle 
      Height          =   315
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton cmdRotate 
      Caption         =   "Rotate"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton cmdSmallPrev 
      Caption         =   "Small Preview"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label lblAngle 
      AutoSize        =   -1  'True
      Caption         =   "Angle:"
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   450
   End
End
Attribute VB_Name = "frmRotate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdRotate_Click()
    Dim angle As Single
    
    If Not (cboAngle.ListIndex = -1) Then
        angle = cboAngle.List(cboAngle.ListIndex)
        angle = Pi * angle / 180
        With frmMain
            .picRotate.Height = .picMain.Height
            .picRotate.Width = .picMain.Width
            Call bmp_rotate(.picMain, .picRotate, angle)
            .picRotate.Picture = .picRotate.Image
            .picNew.Height = .picMain.Height
            .picNew.Width = .picMain.Width
            .picMain.PaintPicture .picNew.Image, 0, 0
            IsBitmap = False
            .picMain.PaintPicture .picRotate.Picture, 0, 0, , , , , , , vbSrcCopy
        End With
        frmRotate.Hide
    End If
End Sub

Private Sub cmdSmallPrev_Click()
    Dim angle As Single
    
    If Not (cboAngle.ListIndex = -1) Then
        angle = cboAngle.List(cboAngle.ListIndex)
        angle = Pi * angle / 180
        Call bmp_rotate(frmMain.picMain, picPrev, angle)
    End If
End Sub

Private Sub cmdCancel_Click()
    frmRotate.Hide
End Sub

Private Sub Form_Load()
    For i = 0 To 360
        cboAngle.AddItem i
    Next i
    cboAngle.ListIndex = 90
End Sub
