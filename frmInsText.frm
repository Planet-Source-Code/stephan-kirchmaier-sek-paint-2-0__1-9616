VERSION 5.00
Begin VB.Form frmInsText 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Insert Text"
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   7290
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   7290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.ComboBox cboFonts 
      Height          =   315
      Left            =   5520
      TabIndex        =   16
      Top             =   3600
      Width           =   1695
   End
   Begin VB.ComboBox cboWidth 
      Height          =   315
      Left            =   3720
      TabIndex        =   15
      Top             =   3600
      Width           =   1695
   End
   Begin VB.ComboBox cboHeight 
      Height          =   315
      Left            =   1920
      TabIndex        =   14
      Top             =   3600
      Width           =   1695
   End
   Begin VB.ComboBox cboBold 
      Height          =   315
      Left            =   120
      TabIndex        =   13
      Top             =   3600
      Width           =   1695
   End
   Begin VB.TextBox txtText 
      Height          =   285
      Left            =   120
      TabIndex        =   11
      Text            =   "SEK - Paint 2.0"
      Top             =   2400
      Width           =   3255
   End
   Begin VB.CheckBox chbItalic 
      Caption         =   "Italic"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2880
      Width           =   735
   End
   Begin VB.CheckBox chbUnderlined 
      Caption         =   "Underlined"
      Height          =   255
      Left            =   1080
      TabIndex        =   9
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CheckBox chbStrikeout 
      Caption         =   "Strikeout"
      Height          =   255
      Left            =   2400
      TabIndex        =   8
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   1800
      TabIndex        =   7
      Top             =   1560
      Width           =   1575
   End
   Begin VB.CommandButton cmdDrawPic 
      Caption         =   "Draw in the pic"
      Height          =   495
      Left            =   1800
      TabIndex        =   6
      Top             =   840
      Width           =   1575
   End
   Begin VB.CommandButton cmdDrawPrev 
      Caption         =   "Draw Preview"
      Height          =   495
      Left            =   1800
      TabIndex        =   5
      Top             =   120
      Width           =   1575
   End
   Begin VB.PictureBox picBackBuffer 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   120
      ScaleHeight     =   93
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   93
      TabIndex        =   3
      Top             =   360
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.PictureBox picAngle 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   120
      ScaleHeight     =   93
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   93
      TabIndex        =   2
      Top             =   360
      Width           =   1455
   End
   Begin VB.PictureBox picTest 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   2895
      Left            =   3600
      ScaleHeight     =   189
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   236
      TabIndex        =   0
      Top             =   360
      Width           =   3600
   End
   Begin VB.Label lblFonts 
      AutoSize        =   -1  'True
      Caption         =   "Fonts:"
      Height          =   195
      Left            =   5640
      TabIndex        =   20
      Top             =   3360
      Width           =   435
   End
   Begin VB.Label lblWeight 
      AutoSize        =   -1  'True
      Caption         =   "Weight"
      Height          =   195
      Left            =   120
      TabIndex        =   19
      Top             =   3360
      Width           =   510
   End
   Begin VB.Label lblWidth 
      AutoSize        =   -1  'True
      Caption         =   "Width"
      Height          =   195
      Left            =   3720
      TabIndex        =   18
      Top             =   3360
      Width           =   420
   End
   Begin VB.Label lblH 
      AutoSize        =   -1  'True
      Caption         =   "Height:"
      Height          =   195
      Left            =   1920
      TabIndex        =   17
      Top             =   3360
      Width           =   510
   End
   Begin VB.Label lblText 
      AutoSize        =   -1  'True
      Caption         =   "Text:"
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   2160
      Width           =   360
   End
   Begin VB.Label lblAngle 
      AutoSize        =   -1  'True
      Caption         =   "Angle:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   450
   End
   Begin VB.Label lblPreview 
      AutoSize        =   -1  'True
      Caption         =   "Preview:"
      Height          =   195
      Left            =   3600
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmInsText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const t = 93

Dim det As Boolean, aX, aY

Private Sub cmdCancel_Click()
    frmInsText.Hide
End Sub

Private Sub cmdDrawPic_Click()
    Dim an As Double, l As Double, u As Double, k As Double
    Dim iWeight As Integer
    
    l = Sqr((t / 2 - aX) ^ 2 + (t / 2 - aY) ^ 2)
    u = Sqr((t / 2 - aX) ^ 2)
    k = u / l
    
    If k = 1 Then
        If (aY = t / 2) And (aX < t / 2) Then
            an = 180
        Else
            an = 0
        End If
    ElseIf k = 0 Then
        If (aY > t / 2) Then
            an = 270
        Else
            an = 90
        End If
    Else
        an = (180 * (Atn(-k / Sqr(-k * k + 1)) + 2 * Atn(1))) / (4 * Atn(1))
    End If
    If (aX > t / 2) And (aY > t / 2) Then
        an = 360 - an
    ElseIf (aX < t / 2) And (aY < t / 2) Then
        an = 180 - an
    ElseIf (aX < t / 2) And (aY > t / 2) Then
        an = 180 + an
    End If
    Select Case cboBold.ListIndex
        Case 0: iWeight = 900
        Case 1: iWeight = 800
        Case 2: iWeight = 700
        Case 3: iWeight = 600
        Case 4: iWeight = 500
        Case 5: iWeight = 400
        Case 6: iWeight = 300
        Case 7: iWeight = 200
        Case 8: iWeight = 100
    End Select
    Call DrawText(frmMain.picMain, txtText.Text, sX, sY, chbUnderlined.Value, chbItalic.Value, chbStrikeout.Value, cboHeight.List(cboHeight.ListIndex), cboWidth.List(cboWidth.ListIndex), an, iWeight, cboFonts.List(cboFonts.ListIndex))
    frmInsText.Hide
End Sub

Private Sub cmdDrawPrev_Click()
    Dim an As Double, l As Double, u As Double, k As Double
    Dim iWeight As Integer
    
    l = Sqr((t / 2 - aX) ^ 2 + (t / 2 - aY) ^ 2)
    u = Sqr((t / 2 - aX) ^ 2)
    k = u / l
    
    If k = 1 Then
        If (aY = t / 2) And (aX < t / 2) Then
            an = 180
        Else
            an = 0
        End If
    ElseIf k = 0 Then
        If (aY > t / 2) Then
            an = 270
        Else
            an = 90
        End If
    Else
        an = (180 * (Atn(-k / Sqr(-k * k + 1)) + 2 * Atn(1))) / (4 * Atn(1))
    End If
    If (aX > t / 2) And (aY > t / 2) Then
        an = 360 - an
    ElseIf (aX < t / 2) And (aY < t / 2) Then
        an = 180 - an
    ElseIf (aX < t / 2) And (aY > t / 2) Then
        an = 180 + an
    End If
    Select Case cboBold.ListIndex
        Case 0: iWeight = 900
        Case 1: iWeight = 800
        Case 2: iWeight = 700
        Case 3: iWeight = 600
        Case 4: iWeight = 500
        Case 5: iWeight = 400
        Case 6: iWeight = 300
        Case 7: iWeight = 200
        Case 8: iWeight = 100
    End Select
    Call DrawText(picTest, txtText.Text, picTest.ScaleWidth / 2, picTest.ScaleHeight / 2, chbUnderlined.Value, chbItalic.Value, chbStrikeout.Value, cboHeight.List(cboHeight.ListIndex), cboWidth.List(cboWidth.ListIndex), an, iWeight, cboFonts.List(cboFonts.ListIndex))
End Sub

Private Sub Form_Load()
    Dim i As Long
    
    det = False
    For i = 0 To Printer.FontCount - 1
        cboFonts.AddItem Printer.Fonts(i)
    Next i
    cboFonts.ListIndex = 1
    For i = 2 To 150
        cboHeight.AddItem i
        cboWidth.AddItem i
    Next i
    cboHeight.ListIndex = 20
    cboWidth.ListIndex = 20
    With cboBold
        .AddItem "Heavy"
        .AddItem "Extrabold"
        .AddItem "Bold"
        .AddItem "Semibold"
        .AddItem "Medium"
        .AddItem "Normal"
        .AddItem "Light"
        .AddItem "Extralight"
        .AddItem "Thin"
        .ListIndex = 5
    End With
    Call picAngle_MouseMove(1, 1, t, t / 2)
    aX = t
    aY = t / 2
End Sub

Private Sub picAngle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not det Then
        det = True
        aX = X
        aY = Y
    Else
        det = False
    End If
End Sub

Private Sub picAngle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not det Then
        picBackBuffer.Circle (t / 2, t / 2), (t / 2) - 8, vbBlack
        picBackBuffer.Line (t / 2, 4)-(t / 2, 12), vbBlack
        picBackBuffer.Line (t / 2, t - 12)-(t / 2, t - 4), vbBlack
        picBackBuffer.Line (t - 12, t / 2)-(t - 4, t / 2), vbBlack
        picBackBuffer.Line (4, t / 2)-(12, t / 2), vbBlack
        picBackBuffer.Line (t / 2, t / 2)-(X, Y), vbRed
        picAngle.Picture = picBackBuffer.Image
        picBackBuffer.Cls
    End If
End Sub

Private Sub picTest_Click()
    picTest.Cls
End Sub

Private Sub DrawText(Obj As Object, Text, X As Long, Y As Long, Underlined As Boolean, Italic As Boolean, Strike As Boolean, Height As Integer, Width As Integer, ByVal Angle As Integer, ByVal FWidth As Long, FName As String)
    Dim lHFont As Long, lTFont As Long

    lHFont = CreateFont(Height, Width, Angle * 10, Angle * 10, FWidth, CLng(Italic), CLng(Underlined), CLng(Strike), DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, PROOF_QUALITY, FF_DONTCARE, FName)
    lTFont = SelectObject(Obj.hdc, lHFont)
    Obj.CurrentX = X
    Obj.CurrentY = Y
    Obj.ForeColor = ForeCol
    Obj.Print Text
    SelectObject Obj.hdc, lTFont
    DeleteObject lHFont
End Sub
