VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPrint 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Print the picture"
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   5475
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   5475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2880
      TabIndex        =   18
      Top             =   2640
      Width           =   2415
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   495
      Left            =   2880
      TabIndex        =   17
      Top             =   1920
      Width           =   2415
   End
   Begin VB.TextBox txtTop 
      Alignment       =   1  'Rechts
      Height          =   285
      Left            =   4560
      TabIndex        =   12
      Text            =   "3"
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox txtLeft 
      Height          =   285
      Left            =   4560
      TabIndex        =   11
      Text            =   "3"
      Top             =   1320
      Width           =   735
   End
   Begin MSComctlLib.Slider ctlNumCopies 
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2880
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   450
      _Version        =   393216
      Min             =   1
   End
   Begin VB.Frame fraPQu 
      Caption         =   "Print Quality"
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
      Begin VB.OptionButton optBad 
         Caption         =   "Bad Quality"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton optGood 
         Caption         =   "Good Quality"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   1335
      End
      Begin VB.OptionButton optBetter 
         Caption         =   "Better Quality"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   1335
      End
      Begin VB.OptionButton optBest 
         Caption         =   "Best Quality"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1440
         Width           =   1215
      End
      Begin VB.OptionButton optUDef 
         Caption         =   "User defined:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   1800
         Width           =   1275
      End
      Begin VB.TextBox txtDpi 
         Alignment       =   1  'Rechts
         Height          =   285
         Left            =   1560
         MaxLength       =   4
         TabIndex        =   1
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label lblDpi 
         AutoSize        =   -1  'True
         Caption         =   "dpi"
         Height          =   195
         Left            =   2160
         TabIndex        =   7
         Top             =   1920
         Width           =   210
      End
   End
   Begin VB.Label lblTop 
      AutoSize        =   -1  'True
      Caption         =   "Distance from the Top of the page"
      Height          =   195
      Left            =   2880
      TabIndex        =   16
      Top             =   120
      Width           =   2430
   End
   Begin VB.Label lblPix1 
      AutoSize        =   -1  'True
      Caption         =   "in Centimeters:"
      Height          =   195
      Left            =   2880
      TabIndex        =   15
      Top             =   480
      Width           =   1035
   End
   Begin VB.Label lblLeft 
      AutoSize        =   -1  'True
      Caption         =   "Distance from the Left of the page"
      Height          =   195
      Left            =   2880
      TabIndex        =   14
      Top             =   960
      Width           =   2415
   End
   Begin VB.Label lblPix2 
      AutoSize        =   -1  'True
      Caption         =   "in Centimeters:"
      Height          =   195
      Left            =   2880
      TabIndex        =   13
      Top             =   1320
      Width           =   1035
   End
   Begin VB.Label lblNumCopEx 
      AutoSize        =   -1  'True
      Caption         =   "Number of copies:"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   2520
      Width           =   1290
   End
   Begin VB.Label lblNumCop 
      AutoSize        =   -1  'True
      Caption         =   "01"
      Height          =   195
      Left            =   2400
      TabIndex        =   8
      Top             =   2520
      Width           =   180
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    frmPrint.Hide
End Sub

Private Sub cmdPrint_Click()
    Dim iQuality As Integer

    On Error Resume Next
    If Not IsNumeric(txtLeft.Text) Then
        MsgBox "Type in a number, LEFT.", vbCritical, "Error"
        Exit Sub
    End If
    If Not IsNumeric(txtTop.Text) Then
        MsgBox "Type in a number, TOP.", vbCritical, "Error"
        Exit Sub
    End If
    If optUDef.Value And Not IsNumeric(txtDpi.Text) Then
        MsgBox "Type in a valid number, DPI.", vbCritical, "Error"
        Exit Sub
    End If
    
    If optBad.Value Then iQuality = -1
    If optGood.Value Then iQuality = -2
    If optBetter.Value Then iQuality = -3
    If optBest.Value Then iQuality = -4
    If optUDef.Value Then iQuality = txtDpi.Text
    
    With Printer
        .Copies = ctlNumCopies.Value
        .PrintQuality = iQuality
        .ScaleMode = vbCentimeters
        .Height = frmMain.picMain.ScaleHeight
        .Width = frmMain.picMain.ScaleWidth
        .ScaleLeft = txtLeft.Text
        .ScaleTop = txtTop.Text
    End With
    frmMain.picPrint.Picture = frmMain.picMain.Image
    
    Printer.PaintPicture frmMain.picPrint.Picture, 0, 0
    Printer.EndDoc
    frmPrint.Hide
End Sub

Private Sub ctlNumCopies_Click()
    If ctlNumCopies.Value < 10 Then
        lblNumCop.Caption = "0" & CStr(ctlNumCopies.Value)
    Else
        lblNumCop.Caption = CStr(ctlNumCopies.Value)
    End If
End Sub

Private Sub Form_Load()
    ctlNumCopies.Value = 1
    optGood.Value = True
End Sub
