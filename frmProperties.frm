VERSION 5.00
Begin VB.Form frmProperties 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Picture - Information"
   ClientHeight    =   1860
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   5295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1860
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Label lblPath 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblBitDepth 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   480
   End
   Begin VB.Label lblWidth 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   480
   End
   Begin VB.Label lblHeight 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   480
   End
   Begin VB.Label lblRLE 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   480
   End
   Begin VB.Label lblPlanes 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   480
   End
   Begin VB.Label lblSize 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   480
   End
End
Attribute VB_Name = "frmProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type BITMAPFILEHEADER
    bfType As Integer
    bfSize As Long
    bfReserved1 As Integer
    bfReserved2 As Integer
    bfOffBits As Long
End Type

Private Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Private Sub Form_Load()
    Dim iFreeNum As Integer
    Dim tBmpFileHeader As BITMAPFILEHEADER
    Dim tBmpInfoHeader As BITMAPINFOHEADER
    
    On Error Resume Next
    iFreeNum = FreeFile()
    If IsBitmap And FileExist(sBmpPath) Then
        Open sBmpPath For Binary Access Read Lock Write As iFreeNum
            Get iFreeNum, 1, tBmpFileHeader
            Get iFreeNum, , tBmpInfoHeader
        Close iFreeNum
    Else
        sBmpPath = App.Path & "\TEMP.bmp"
        SavePicture frmMain.picMain.Image, sBmpPath
        Open sBmpPath For Binary Access Read Lock Write As iFreeNum
            Get iFreeNum, 1, tBmpFileHeader
            Get iFreeNum, , tBmpInfoHeader
        Close iFreeNum
        Kill sBmpPath
    End If
    lblBitDepth.Caption = "Number of usable colors: " & 2 ^ tBmpInfoHeader.biBitCount
    lblHeight.Caption = "Height in Pixels: " & tBmpInfoHeader.biHeight
    lblWidth.Caption = "Width in Pixels: " & tBmpInfoHeader.biWidth
    lblPath.Caption = "Path: " & sBmpPath
    lblPlanes.Caption = "Number of Planes: " & tBmpInfoHeader.biPlanes
    lblRLE.Caption = "Is RLE used? " & CBool(tBmpInfoHeader.biCompression)
    lblSize.Caption = "Size: " & tBmpFileHeader.bfSize \ 1024 & " KB"
End Sub
