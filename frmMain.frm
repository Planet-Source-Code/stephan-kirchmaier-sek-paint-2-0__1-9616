VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "SEK - Paint 2.0"
   ClientHeight    =   7845
   ClientLeft      =   165
   ClientTop       =   480
   ClientWidth     =   9135
   LinkTopic       =   "Form1"
   ScaleHeight     =   523
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   609
   StartUpPosition =   3  'Windows-Standard
   Begin MSComDlg.CommonDialog ctlCommDiag 
      Left            =   9600
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picBackCol 
      Height          =   255
      Left            =   8400
      ScaleHeight     =   195
      ScaleWidth      =   315
      TabIndex        =   11
      Top             =   7560
      Width           =   375
   End
   Begin VB.PictureBox picForeCol 
      Height          =   255
      Left            =   7920
      ScaleHeight     =   195
      ScaleWidth      =   315
      TabIndex        =   10
      Top             =   7560
      Width           =   375
   End
   Begin MSComctlLib.Toolbar ctlTools 
      Align           =   1  'Oben ausrichten
      Height          =   420
      Left            =   0
      TabIndex        =   7
      Top             =   420
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ctlImgTools"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   16
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmPencil"
            Object.ToolTipText     =   "Draw with a pencil"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmStar"
            Object.ToolTipText     =   "Draw a star"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmVertLine"
            Object.ToolTipText     =   "Draw a vertical line"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmHorzLine"
            Object.ToolTipText     =   "Draw a horizontal line"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmDiagLineRL"
            Object.ToolTipText     =   "Draw a diagonal line (/)"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmDiagLineLR"
            Object.ToolTipText     =   "Draw a diagonal line (\)"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmCross"
            Object.ToolTipText     =   "Draw a cross"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmDiagCross"
            Object.ToolTipText     =   "Draw a diagonal cross"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmPolygon"
            Object.ToolTipText     =   "Draw a polygon"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmGetCol"
            Object.ToolTipText     =   "Change the drawing color"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmErase"
            Object.ToolTipText     =   "Erase some pixel"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmInsertText"
            Object.ToolTipText     =   "Shows the Insert Text - Dialog."
            ImageIndex      =   15
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmBrush"
            Object.ToolTipText     =   "Draw with a brush"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmStLine"
            Object.ToolTipText     =   "Draw a straight line"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmFillRgn"
            Object.ToolTipText     =   "Fill Regions with the drawing - color"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmCircRect"
            Object.ToolTipText     =   "Draw circles or rects"
            ImageIndex      =   1
         EndProperty
      EndProperty
      Begin VB.ComboBox cboEffects 
         Height          =   315
         Left            =   7440
         TabIndex        =   9
         Top             =   0
         Width           =   1695
      End
      Begin VB.ComboBox cboFilters 
         Height          =   315
         Left            =   5640
         TabIndex        =   8
         Top             =   0
         Width           =   1695
      End
   End
   Begin MSComctlLib.ProgressBar ctlProgBar 
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   7200
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.Toolbar ctlToolBar 
      Align           =   1  'Oben ausrichten
      Height          =   420
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ctlImgList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmNew"
            Object.ToolTipText     =   "Create a new picture"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmOpen"
            Object.ToolTipText     =   "Open a picture"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmSave"
            Object.ToolTipText     =   "Save the Picture to the .bmp-File-Format"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmCopy"
            Object.ToolTipText     =   "Copy a part of the picture to the clipboard"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmCut"
            Object.ToolTipText     =   "Cut out a part of the picture"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmPaste"
            Object.ToolTipText     =   "Paste the content of the clipboard into a new picture"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmProps"
            Object.ToolTipText     =   "Show the Properties - Dialog"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmUndo"
            Object.ToolTipText     =   "Undo the last action"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmPrint"
            Object.ToolTipText     =   "Show the Printing - Dialog"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
   Begin VB.VScrollBar scrVert 
      Height          =   6015
      Left            =   8880
      TabIndex        =   4
      Top             =   840
      Width           =   255
   End
   Begin VB.HScrollBar scrHorz 
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   6840
      Width           =   8895
   End
   Begin MSComctlLib.StatusBar ctlStatBar 
      Align           =   2  'Unten ausrichten
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   7485
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Key             =   "mnuMouseC"
            Object.ToolTipText     =   "Mouse - Coords"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Key             =   "mnuCurTool"
            Object.ToolTipText     =   "Current Tool"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picBack 
      Height          =   6015
      Left            =   0
      ScaleHeight     =   397
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   589
      TabIndex        =   0
      Top             =   840
      Width           =   8895
      Begin VB.PictureBox picRotate 
         Appearance      =   0  '2D
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   6360
         ScaleHeight     =   23
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   135
         TabIndex        =   18
         Top             =   3000
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.PictureBox picCutCopy 
         Appearance      =   0  '2D
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   6360
         ScaleHeight     =   23
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   135
         TabIndex        =   17
         Top             =   2520
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.PictureBox picPrint 
         Appearance      =   0  '2D
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   6360
         ScaleHeight     =   23
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   135
         TabIndex        =   16
         Top             =   2040
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.PictureBox picInvert 
         Appearance      =   0  '2D
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   6360
         ScaleHeight     =   23
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   135
         TabIndex        =   15
         Top             =   1560
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.PictureBox picFlip 
         Appearance      =   0  '2D
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   6360
         ScaleHeight     =   23
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   135
         TabIndex        =   14
         Top             =   1080
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.PictureBox picNew 
         Appearance      =   0  '2D
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   6360
         ScaleHeight     =   23
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   135
         TabIndex        =   13
         Top             =   600
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.PictureBox picUndo 
         Appearance      =   0  '2D
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   6360
         ScaleHeight     =   23
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   135
         TabIndex        =   12
         Top             =   120
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.PictureBox picMain 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   4725
         Left            =   0
         ScaleHeight     =   311
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   351
         TabIndex        =   1
         Top             =   0
         Width           =   5325
      End
   End
   Begin MSComctlLib.ImageList ctlImgTools 
      Left            =   9600
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":011C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0438
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0754
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0A70
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0D8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":12FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1618
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1934
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A50
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1EA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":22F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2614
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2A68
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":31BC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ctlImgList1 
      Left            =   9600
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3610
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3724
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3838
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":394C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3A60
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3B74
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3C88
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3D9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3EB0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Test"
      Visible         =   0   'False
      Begin VB.Menu mnuGetPicInfo 
         Caption         =   "Get Picture Information"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu mnuCut 
         Caption         =   "Cut"
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "Paste"
      End
      Begin VB.Menu mnuProps 
         Caption         =   "Properties"
         Begin VB.Menu mnuDrawWidth 
            Caption         =   "Set the DrawWidth"
         End
         Begin VB.Menu mnuDrawStyle 
            Caption         =   "DrawStyle"
            Begin VB.Menu mnuDSFilled 
               Caption         =   "Filled"
               Checked         =   -1  'True
            End
            Begin VB.Menu mnuDSLine 
               Caption         =   "Line"
            End
            Begin VB.Menu mnuDSPoint 
               Caption         =   "Point"
            End
            Begin VB.Menu mnuDSLinePoint 
               Caption         =   "Line-Point"
            End
            Begin VB.Menu mnuDSLinePointPoint 
               Caption         =   "Line-Point-Point"
            End
         End
         Begin VB.Menu mnuFillStyle 
            Caption         =   "FillStyle"
            Begin VB.Menu mnuFSTFilled 
               Caption         =   "Filled"
               Checked         =   -1  'True
            End
            Begin VB.Menu mnuFSTHorzLine 
               Caption         =   "Horizontal Line"
            End
            Begin VB.Menu mnuFSTVertLine 
               Caption         =   "Vertical Line"
            End
            Begin VB.Menu mnuFSTDiagLineLR 
               Caption         =   "Diagonal Line (\)"
            End
            Begin VB.Menu mnuFSTDiagLineRL 
               Caption         =   "Diagonal Line (/)"
            End
            Begin VB.Menu mnuFSTCross 
               Caption         =   "Cross"
            End
            Begin VB.Menu mnuFSTDiagCross 
               Caption         =   "Diagonal Cross"
            End
         End
      End
      Begin VB.Menu mnuTools 
         Caption         =   "Tools"
         Begin VB.Menu mnuPencil 
            Caption         =   "Pencil"
         End
         Begin VB.Menu mnuStar 
            Caption         =   "Star"
         End
         Begin VB.Menu mnuHorzLine 
            Caption         =   "Horizontal Line"
         End
         Begin VB.Menu mnuVertLine 
            Caption         =   "Vertical Line"
         End
         Begin VB.Menu mnuCross 
            Caption         =   "Cross"
         End
         Begin VB.Menu mnuDiagCross 
            Caption         =   "Diagonal Cross"
         End
         Begin VB.Menu mnuDiagLineLR 
            Caption         =   "Diagonal Line (\)"
         End
         Begin VB.Menu mnuDiagLineRL 
            Caption         =   "Diagonal Line (/)"
         End
         Begin VB.Menu mnuUserDefPolygon 
            Caption         =   "Userdefined Polygon"
         End
         Begin VB.Menu mnuText 
            Caption         =   "Insert Text"
         End
         Begin VB.Menu mnuStraightLine 
            Caption         =   "Straight Line"
         End
         Begin VB.Menu mnuBrush 
            Caption         =   "Brush"
         End
         Begin VB.Menu mnuPolygon 
            Caption         =   "Polygon"
         End
         Begin VB.Menu mnuRect 
            Caption         =   "Rect"
         End
         Begin VB.Menu mnuCircle 
            Caption         =   "Circle"
         End
         Begin VB.Menu mnuFilledRect 
            Caption         =   "Filled Rect"
         End
         Begin VB.Menu mnuFilledCircle 
            Caption         =   "Filled Circle"
         End
         Begin VB.Menu mnuFillRgn 
            Caption         =   "Fill Regions"
         End
         Begin VB.Menu mnuErase 
            Caption         =   "Erase"
         End
         Begin VB.Menu mnuChooseCol 
            Caption         =   "Choose Color"
         End
      End
      Begin VB.Menu mnuFilters 
         Caption         =   "Filters"
         Begin VB.Menu mnuEmboss 
            Caption         =   "Emboss"
         End
         Begin VB.Menu mnuSharpen 
            Caption         =   "Sharpen"
         End
         Begin VB.Menu mnuDiffuse 
            Caption         =   "Diffuse"
         End
         Begin VB.Menu mnuRects 
            Caption         =   "Rects"
         End
         Begin VB.Menu mnuBrightness 
            Caption         =   "Brightness"
         End
         Begin VB.Menu mnuIce 
            Caption         =   "Ice"
         End
         Begin VB.Menu mnuDark 
            Caption         =   "Dark"
         End
         Begin VB.Menu mnuHeat 
            Caption         =   "Heat"
         End
         Begin VB.Menu mnuStrange 
            Caption         =   "Strange"
         End
         Begin VB.Menu mnuAqua 
            Caption         =   "Aqua"
         End
         Begin VB.Menu mnuNight 
            Caption         =   "Night"
         End
         Begin VB.Menu mnuCrazyLines 
            Caption         =   "Crazy Lines"
         End
         Begin VB.Menu mnuAfrica 
            Caption         =   "Africa"
         End
         Begin VB.Menu mnuBlur 
            Caption         =   "Blur"
         End
         Begin VB.Menu mnuInvert 
            Caption         =   "Invert"
         End
         Begin VB.Menu mnuGreyscale 
            Caption         =   "Greyscale"
         End
         Begin VB.Menu mnuComic 
            Caption         =   "Comic"
         End
         Begin VB.Menu mnuBnW 
            Caption         =   "Black and White"
         End
      End
      Begin VB.Menu mnuEffects 
         Caption         =   "Effects"
         Begin VB.Menu mnuFlip1 
            Caption         =   "Flip1"
         End
         Begin VB.Menu mnuFlip2 
            Caption         =   "Flip2"
         End
         Begin VB.Menu mnuFlip3 
            Caption         =   "Flip3"
         End
         Begin VB.Menu mnuRepColor 
            Caption         =   "Replace Color"
         End
         Begin VB.Menu mnuWave 
            Caption         =   "Wave"
         End
         Begin VB.Menu mnuHammer 
            Caption         =   "Hammer"
         End
         Begin VB.Menu mnuHook 
            Caption         =   "Hook"
         End
         Begin VB.Menu mnuRotate 
            Caption         =   "Rotate"
         End
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
   Begin VB.Menu mnuShapes 
      Caption         =   "Shapes"
      Visible         =   0   'False
      Begin VB.Menu mnusRect 
         Caption         =   "Rect"
      End
      Begin VB.Menu mnusCircle 
         Caption         =   "Circle"
      End
      Begin VB.Menu mnusFRect 
         Caption         =   "Filled Rect"
      End
      Begin VB.Menu mnusFCircle 
         Caption         =   "Filled Circle"
      End
   End
   Begin VB.Menu mnuPolGon 
      Caption         =   "PolyGon"
      Visible         =   0   'False
      Begin VB.Menu mnuNormPolygon 
         Caption         =   "Normal"
      End
      Begin VB.Menu mnuUDefPolygon 
         Caption         =   "Userdefined"
      End
   End
   Begin VB.Menu mnuChooseCol2 
      Caption         =   "ChooseCol"
      Visible         =   0   'False
      Begin VB.Menu mnuOnPic 
         Caption         =   "On the Picture"
      End
      Begin VB.Menu mnuWithDialog 
         Caption         =   "With the Dialog"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboEffects_Click()
    If Marked Then
        MarkRect picMain, StartX, StartY, OldX, OldY
        Marked = False
        NowMove = False
    End If
    Select Case cboEffects.ListIndex
        Case 0:
            Exit Sub
        Case 6:
            curTool = Tools.EfHammer
            Call RenPanel
        Case 7:
            curTool = Tools.EfHook
            Call RenPanel
        Case 4:
            curTool = Tools.RepCol
            Call RenPanel
        Case 1:
            Call mnuFlip1_Click
        Case 2:
            Call mnuFlip2_Click
        Case 3:
            Call mnuFlip3_Click
        Case 5:
            Call mnuWave_Click
        Case 8:
            Call mnuRotate_Click
    End Select
    cboEffects.ListIndex = 0
End Sub


Private Sub cboFilters_Click()
    If Marked Then
        MarkRect picMain, StartX, StartY, OldX, OldY
        Marked = False
        NowMove = False
    End If
    Select Case cboFilters.ListIndex
        Case 0:
            Exit Sub
        Case 1:
            Call mnuEmboss_Click
        Case 2:
            Call mnuSharpen_Click
        Case 3:
            Call mnuDiffuse_Click
        Case 4:
            Call mnuRects_Click
        Case 5:
            Call mnuBrightness_Click
        Case 6:
            Call mnuIce_Click
        Case 7:
            Call mnuDark_Click
        Case 8:
            Call mnuHeat_Click
        Case 9:
            Call mnuStrange_Click
        Case 10:
            Call mnuAqua_Click
        Case 11:
            Call mnuNight_Click
        Case 12:
            Call mnuCrazyLines_Click
        Case 13:
            Call mnuAfrica_Click
        Case 14:
            Call mnuBlur_Click
        Case 15:
            Call mnuInvert_Click
        Case 16:
            Call mnuGreyscale_Click
        Case 17:
            Call mnuComic_Click
        Case 18:
            Call mnuBnW_Click
    End Select
    cboFilters.ListIndex = 0
End Sub

Private Sub ctlToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim sFName As String, iFreeNum As Integer, iIsBmp As Integer
    On Error Resume Next
    
    Select Case Button.Key
        Case "cmProps":
            PopupMenu mnuProps
        
        Case "cmCopy":
            Call mnuCopy_Click
            
        Case "cmCut":
            Call mnuCut_Click
            
        Case "cmNew":
            With picNew
                .Height = picMain.Height
                .Width = picMain.Width
            End With
            picMain.PaintPicture picNew.Image, 0, 0
            IsBitmap = False
            
        Case "cmPrint":
            frmPrint.Show 1, Me
        
        Case "cmUndo":
            picMain.Picture = picUndo.Image
        
        Case "cmOpen":
            ctlCommDiag.Filter = "*.bmp;*.jpg;*.gif;*.wmf;*.cur;*.ico"
            ctlCommDiag.ShowOpen
            sFName = ctlCommDiag.filename
            
            If sFName = vbNullString Then Exit Sub
            
            If Not FileExist(sFName) Then
                MsgBox "File doesn't exist.", vbCritical, "Error"
                Exit Sub
            End If
            iFreeNum = FreeFile()
            Open sFName For Binary Access Read Lock Write As iFreeNum
                Get iFreeNum, 1, iIsBmp
            Close iFreeNum
            IsBitmap = (iIsBmp = 19778)
            sBmpPath = sFName
            picMain.Picture = LoadPicture(sFName)
            Call PrepPic
            Call Save
        
        Case "cmSave":
            ctlCommDiag.Filter = "*.bmp"
            ctlCommDiag.ShowSave
            sFName = ctlCommDiag.filename
            
            If sFName = vbNullString Then Exit Sub
            
            If FileExist(sFName) Then
                MsgBox "The File already exist!", vbCritical, "Error"
                Exit Sub
            End If
            
            If (LCase$(Right$(sFName, 4)) = ".bmp") Then
                SavePicture picMain.Image, sFName
            Else
                sFName = sFName & ".bmp"
                SavePicture picMain.Image, sFName
            End If
            
        Case "cmPaste":
            Call mnuPaste_Click

    End Select
End Sub

Private Sub Form_Load()

    Call Save
    With cboFilters
        .AddItem "Filters"
        .AddItem "Emboss"
        .AddItem "Sharpen"
        .AddItem "Diffuse"
        .AddItem "Rects"
        .AddItem "Brightness"
        .AddItem "Ice"
        .AddItem "Darkness"
        .AddItem "Heat"
        .AddItem "Strange"
        .AddItem "Aqua"
        .AddItem "Night"
        .AddItem "Crazy Lines"
        .AddItem "Africa"
        .AddItem "Blur"
        .AddItem "Invert"
        .AddItem "Greyscale"
        .AddItem "Comic"
        .AddItem "Black and White"
        .ListIndex = 0
    End With
    With cboEffects
        .AddItem "Effects"
        .AddItem "Flip1"
        .AddItem "Flip2"
        .AddItem "Flip3"
        .AddItem "Replace Color"
        .AddItem "Wave"
        .AddItem "Hammer"
        .AddItem "Hook"
        .AddItem "Rotate"
        .ListIndex = 0
    End With
    For i = 0 To 360
        sinTab(i) = Sin(i)
        cosTab(i) = Cos(i)
    Next i
    IsCurOK = IsCurOKFirst
    ForeCol = vbBlack
    FillCol = vbRed
    picBackCol.BackColor = vbRed
    picForeCol.BackColor = vbBlack
    propFillStyle = 0
    curTool = Tools.Pencil
    picMain.FillStyle = 0
    picMain.DrawStyle = 0
    Call PrepPic
    Call RenPanel
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ctlStatBar.Panels(1).Text = vbNullString
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    picBack.Width = frmMain.ScaleWidth - scrVert.Width
    picBack.Height = frmMain.ScaleHeight - (ctlStatBar.Height + ctlProgBar.Height + scrHorz.Height + 2 * ctlTools.Height)
    scrVert.Left = picBack.Width
    scrVert.Height = picBack.Height
    scrHorz.Top = 2 * ctlTools.Height + picBack.Height
    scrHorz.Width = picBack.Width
    ctlProgBar.Top = scrHorz.Top + scrHorz.Height
    ctlProgBar.Width = picBack.Width + scrVert.Width
    picBackCol.Top = frmMain.ScaleHeight - (5 + picBackCol.Height)
    picBackCol.Left = scrHorz.Width - picBackCol.Width
    picForeCol.Left = picBackCol.Left - (5 + picForeCol.Width)
    picForeCol.Top = picBackCol.Top
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show 1, Me
End Sub

Private Sub mnuAfrica_Click()
    Call Save
    For i = 0 To curX
        For j = 0 To curY
            tCOl = GetPixel(picMain.hdc, i, j)
            r = tCOl Mod 256
            g = (tCOl \ 256) Mod 256
            b = tCOl \ 256 \ 256
            r = Abs((g * b) / 256)
            g = Abs((b * r) / 256)
            b = Abs((r * g) / 256)
            SetPixel picMain.hdc, i, j, RGB(r, g, b)
        Next j
        ctlProgBar.Value = i * 100 \ (curX - 1)
    Next i
    picMain.Refresh
    ctlProgBar.Value = 0
End Sub

Private Sub mnuAqua_Click()
    Call Save
    For i = 0 To curX
        For j = 0 To curY
            tCOl = GetPixel(picMain.hdc, i, j)
            r = tCOl Mod 256
            g = (tCOl \ 256) Mod 256
            b = tCOl \ 256 \ 256
            r = (g - b) ^ 2 / 125
            g = (r - b) ^ 2 / 125
            b = (r - g) ^ 2 / 125
            SetPixel picMain.hdc, i, j, RGB(r, g, b)
        Next j
        ctlProgBar.Value = i * 100 \ (curX - 1)
    Next i
    picMain.Refresh
    ctlProgBar.Value = 0
End Sub

Private Sub mnuBlur_Click()
    Call Save
    Call PrepareImg
    For i = 1 To curX - 1
        For j = 1 To curY - 1
            r = Abs(lArrCol(0, i - 1, j - 1) + lArrCol(0, i, j - 1) + lArrCol(0, i + 1, j - 1) + lArrCol(0, i - 1, j) + lArrCol(0, i, j) + lArrCol(0, i + 1, j) + lArrCol(0, i - 1, j + 1) + lArrCol(0, i, j + 1) + lArrCol(0, i + 1, j + 1))
            g = Abs(lArrCol(1, i - 1, j - 1) + lArrCol(1, i, j - 1) + lArrCol(1, i + 1, j - 1) + lArrCol(1, i - 1, j) + lArrCol(1, i, j) + lArrCol(1, i + 1, j) + lArrCol(1, i - 1, j + 1) + lArrCol(1, i, j + 1) + lArrCol(1, i + 1, j + 1))
            b = Abs(lArrCol(2, i - 1, j - 1) + lArrCol(2, i, j - 1) + lArrCol(2, i + 1, j - 1) + lArrCol(2, i - 1, j) + lArrCol(2, i, j) + lArrCol(2, i + 1, j) + lArrCol(2, i - 1, j + 1) + lArrCol(2, i, j + 1) + lArrCol(2, i + 1, j + 1))
            SetPixel picMain.hdc, i, j, RGB(r / 10, g / 10, b / 10)
        Next j
        ctlProgBar.Value = i * 100 \ (curX - 1)
    Next i
    picMain.Refresh
    ctlProgBar.Value = 0
End Sub

Private Sub mnuBnW_Click()
    Call Save
    For i = 0 To curX
        For j = 0 To curY
            tCOl = GetPixel(picMain.hdc, i, j)
            r = tCOl Mod 256
            g = (tCOl Mod 256) \ 256
            b = tCOl \ 256 \ 256
            
            If r < 200 And g < 200 And b < 200 Then
                tCOl = vbBlack
            Else
                tCOl = vbWhite
            End If
            SetPixel picMain.hdc, i, j, tCOl
        Next j
        ctlProgBar.Value = i * 100 \ (curX - 1)
    Next i
    picMain.Refresh
    ctlProgBar.Value = 0
End Sub

Private Sub mnuBrightness_Click()
    Dim c As Long

    Call Save
    For i = 0 To curX
        For j = 0 To curY
            tCOl = GetPixel(picMain.hdc, i, j)
            r = tCOl Mod 256
            g = (tCOl \ 256) Mod 256
            b = tCOl \ 256 \ 256
            c = Abs((r + g + b) \ 3)
            r = Abs(r + c)
            g = Abs(g + c)
            b = Abs(b + c)
            SetPixel picMain.hdc, i, j, RGB(r, g, b)
        Next j
        ctlProgBar.Value = i * 100 \ (curX - 1)
    Next i
    picMain.Refresh
    ctlProgBar.Value = 0
End Sub

Private Sub mnuBrush_Click()
    curTool = Tools.Brush
    cboEffects.ListIndex = 0
    Call RenPanel
End Sub

Private Sub mnuChooseCol_Click()
    curTool = Tools.pinp
    cboEffects.ListIndex = 0
    Call RenPanel
End Sub

Private Sub mnuCircle_Click()
    curTool = Tools.tCircle
    cboEffects.ListIndex = 0
    Call RenPanel
End Sub

Private Sub mnuComic_Click()
    Call Save
    For i = 0 To curX
        For j = 0 To curY
            tCOl = GetPixel(picMain.hdc, i, j)
            r = Abs(tCOl Mod 256)
            g = Abs((tCOl \ 256) Mod 256)
            b = Abs(tCOl \ 256 \ 256)
            r = Abs(r * (g - b + g + r)) / 256
            g = Abs(r * (b - g + b + r)) / 256
            b = Abs(g * (b - g + b + r)) / 256
            tCOl = RGB(r, g, b)
            r = Abs(tCOl Mod 256)
            g = Abs((tCOl \ 256) Mod 256)
            b = Abs(tCOl \ 256 \ 256)
            r = (r + g + b) / 3
            tCOl = RGB(r, r, r)
            SetPixel picMain.hdc, i, j, tCOl
        Next j
        ctlProgBar.Value = i * 100 \ (curX - 1)
    Next i
    picMain.Refresh
    ctlProgBar.Value = 0
End Sub

Private Sub mnuCopy_Click()
    curTool = Tools.btCopy
    Call RenPanel
    If cboEffects.ListIndex = 6 Or cboEffects.ListIndex = 7 Then cboEffects.ListIndex = 0
End Sub

Private Sub mnuCrazyLines_Click()
    Dim tColA As Long, tColB As Long, tColC As Long, tColD As Long, tColE As Long, tColF As Long, tColG As Long, tColH As Long

    Call Save
    For i = 0 To curX
        For j = 0 To curY
            tColA = GetPixel(picMain.hdc, i - 1, j - 1)
            tColB = GetPixel(picMain.hdc, i, j - 1)
            tColC = GetPixel(picMain.hdc, i + 1, j - 1)
            tColD = GetPixel(picMain.hdc, i - 1, j)
            tColE = GetPixel(picMain.hdc, i + 1, j)
            tColF = GetPixel(picMain.hdc, i - 1, j + 1)
            tColG = GetPixel(picMain.hdc, i, j + 1)
            tColH = GetPixel(picMain.hdc, i + 1, j + 1)

            tCOl = (tColA + tColB + tColC + tColD + tColE + tColF + tColG + tColH) / 8
            SetPixel picMain.hdc, i, j, tCOl
        Next j
        ctlProgBar.Value = i * 100 \ (curX - 1)
    Next i
    picMain.Refresh
    ctlProgBar.Value = 0
End Sub

Private Sub mnuCross_Click()
    curTool = Tools.Cross
    cboEffects.ListIndex = 0
    Call RenPanel
End Sub

Private Sub mnuCut_Click()
    curTool = Tools.btCut
    Call RenPanel
    If cboEffects.ListIndex = 6 Or cboEffects.ListIndex = 7 Then cboEffects.ListIndex = 0
End Sub

Private Sub mnuDark_Click()
    Call Save
    For i = 0 To curX
        For j = 0 To curY
            tCOl = GetPixel(picMain.hdc, i, j)
            r = tCOl Mod 256
            g = (tCOl \ 256) Mod 256
            b = tCOl \ 256 \ 256
            r = Abs(r - 64)
            g = Abs(r - 64)
            b = Abs(r - 64)
            SetPixel picMain.hdc, i, j, RGB(r, g, b)
        Next j
        ctlProgBar.Value = i * 100 \ (curX - 1)
    Next i
    picMain.Refresh
    ctlProgBar.Value = 0
End Sub

Private Sub mnuDiagCross_Click()
    curTool = Tools.DiagCross
    cboEffects.ListIndex = 0
    Call RenPanel
End Sub

Private Sub mnuDiagLineLR_Click()
    curTool = Tools.DiagLineLR
    cboEffects.ListIndex = 0
    Call RenPanel
End Sub

Private Sub mnuDiagLineRL_Click()
    curTool = Tools.DiagLineRL
    cboEffects.ListIndex = 0
    Call RenPanel
End Sub

Private Sub mnuDiffuse_Click()
    Dim nP1 As Integer, nP2 As Integer, nP3 As Integer
    Dim tColA As Long, tColB As Long, tColC As Long
    
    Call Save
    For i = 2 To curX - 3
        For j = 2 To curY - 3
            nP1 = Int(Rnd * 5 - 2)
            nP2 = Int(Rnd * 5 - 2)
            nP3 = Int(Rnd * 5 - 2)
            tColA = GetPixel(picMain.hdc, i, j + nP1)
            tColB = GetPixel(picMain.hdc, i + nP2, j)
            tColC = GetPixel(picMain.hdc, i + nP3, j + nP3)
            r = Abs(tColA Mod 256)
            g = Abs((tColB \ 256) Mod 256)
            b = Abs(tColC \ 256 \ 256)
            SetPixel picMain.hdc, i, j, RGB(r, g, b)
        Next j
        ctlProgBar.Value = i * 100 \ (curX - 1)
    Next i
    picMain.Refresh
    ctlProgBar.Value = 0
End Sub

Private Sub mnuDrawWidth_Click()
    Dim iCool As Integer
    
    On Error Resume Next
    iCool = InputBox("Type in the desired Draw Width!", "SEK - Paint 2.0", picMain.DrawWidth)
    If Not IsNumeric(iCool) Or (iCool < 0) Then
        MsgBox "You must type in a valid number!", vbCritical, "Error"
        Exit Sub
    End If
    picMain.DrawWidth = iCool
End Sub

Private Sub mnuDSFilled_Click()
    picMain.DrawStyle = 0
    mnuDSFilled.Checked = True
    mnuDSLine.Checked = False
    mnuDSLinePoint.Checked = False
    mnuDSLinePointPoint.Checked = False
    mnuDSPoint.Checked = False
End Sub

Private Sub mnuDSLine_Click()
    picMain.DrawStyle = 1
    mnuDSFilled.Checked = False
    mnuDSLine.Checked = True
    mnuDSLinePoint.Checked = False
    mnuDSLinePointPoint.Checked = False
    mnuDSPoint.Checked = False
End Sub

Private Sub mnuDSLinePoint_Click()
    picMain.DrawStyle = 3
    mnuDSFilled.Checked = False
    mnuDSLine.Checked = False
    mnuDSLinePoint.Checked = True
    mnuDSLinePointPoint.Checked = False
    mnuDSPoint.Checked = False
End Sub

Private Sub mnuDSLinePointPoint_Click()
    picMain.DrawStyle = 4
    mnuDSFilled.Checked = False
    mnuDSLine.Checked = False
    mnuDSLinePoint.Checked = False
    mnuDSLinePointPoint.Checked = True
    mnuDSPoint.Checked = False
End Sub

Private Sub mnuDSPoint_Click()
    picMain.DrawStyle = 2
    mnuDSFilled.Checked = False
    mnuDSLine.Checked = False
    mnuDSLinePoint.Checked = False
    mnuDSLinePointPoint.Checked = False
    mnuDSPoint.Checked = True
End Sub

Private Sub mnuEmboss_Click()
    Dim tColA As Long
    Dim r1 As Long, g1 As Long, b1 As Long
    
    Call Save
    For i = 0 To curX - 1
        For j = 0 To curY - 1
            tCOl = GetPixel(picMain.hdc, i, j)
            tColA = GetPixel(picMain.hdc, i + 1, j + 1)
            r = tCOl Mod 256
            g = (tCOl \ 256) Mod 256
            b = tCOl \ 256 \ 256
            r1 = tColA Mod 256
            g1 = (tColA \ 256) Mod 256
            b1 = tColA \ 256 \ 256
            r = Abs(r - r1 + 128)
            g = Abs(g - g1 + 128)
            b = Abs(b - b1 + 128)
            SetPixel picMain.hdc, i, j, RGB(r, g, b)
        Next j
        ctlProgBar.Value = i * 100 \ (curX - 1)
    Next i
    picMain.Refresh
    ctlProgBar.Value = 0
End Sub

Private Sub mnuErase_Click()
    curTool = Tools.tErase
    cboEffects.ListIndex = 0
    Call RenPanel
End Sub

Private Sub mnuFilledCircle_Click()
    curTool = Tools.FilledCircle
    cboEffects.ListIndex = 0
    Call RenPanel
End Sub

Private Sub mnuFilledRect_Click()
    curTool = Tools.FilledRect
    cboEffects.ListIndex = 0
    Call RenPanel
End Sub

Private Sub mnuFillRgn_Click()
    curTool = Tools.FillRgn
    cboEffects.ListIndex = 0
    Call RenPanel
End Sub

Private Sub mnuFlip1_Click()
    Call Save
    picFlip.Picture = picMain.Image
    picMain.PaintPicture picFlip.Picture, 0, picMain.ScaleHeight - 1, picMain.ScaleWidth, -picMain.ScaleHeight, , , , , vbSrcCopy
End Sub

Private Sub mnuFlip2_Click()
    Call Save
    picFlip.Picture = picMain.Image
    picMain.PaintPicture picFlip.Picture, picMain.ScaleWidth - 1, 0, -picMain.ScaleWidth, picMain.ScaleHeight, , , , , vbSrcCopy
End Sub

Private Sub mnuFlip3_Click()
    Call Save
    picFlip.Picture = picMain.Image
    picMain.PaintPicture picFlip.Picture, picMain.ScaleWidth - 1, picMain.ScaleHeight - 1, -picMain.ScaleWidth, -picMain.ScaleHeight, , , , , vbSrcCopy
End Sub

Private Sub mnuFSTCross_Click()
    propFillStyle = 6
    mnuFSTFilled.Checked = False
    mnuFSTCross.Checked = True
    mnuFSTDiagCross.Checked = False
    mnuFSTDiagLineRL.Checked = False
    mnuFSTDiagLineLR.Checked = False
    mnuFSTHorzLine.Checked = False
    mnuFSTVertLine.Checked = False
End Sub

Private Sub mnuFSTDiagCross_Click()
    propFillStyle = 7
    mnuFSTFilled.Checked = False
    mnuFSTCross.Checked = False
    mnuFSTDiagCross.Checked = True
    mnuFSTDiagLineRL.Checked = False
    mnuFSTDiagLineLR.Checked = False
    mnuFSTHorzLine.Checked = False
    mnuFSTVertLine.Checked = False
End Sub

Private Sub mnuFSTDiagLineLR_Click()
    propFillStyle = 5
    mnuFSTFilled.Checked = False
    mnuFSTCross.Checked = False
    mnuFSTDiagCross.Checked = False
    mnuFSTDiagLineRL.Checked = False
    mnuFSTDiagLineLR.Checked = True
    mnuFSTHorzLine.Checked = False
    mnuFSTVertLine.Checked = False
End Sub

Private Sub mnuFSTDiagLineRL_Click()
    propFillStyle = 4
    mnuFSTFilled.Checked = False
    mnuFSTCross.Checked = False
    mnuFSTDiagCross.Checked = False
    mnuFSTDiagLineRL.Checked = True
    mnuFSTDiagLineLR.Checked = False
    mnuFSTHorzLine.Checked = False
    mnuFSTVertLine.Checked = False
End Sub

Private Sub mnuFSTFilled_Click()
    propFillStyle = 0
    mnuFSTFilled.Checked = True
    mnuFSTCross.Checked = False
    mnuFSTDiagCross.Checked = False
    mnuFSTDiagLineRL.Checked = False
    mnuFSTDiagLineLR.Checked = False
    mnuFSTHorzLine.Checked = False
    mnuFSTVertLine.Checked = False
End Sub

Private Sub mnuFSTHorzLine_Click()
    propFillStyle = 2
    mnuFSTFilled.Checked = False
    mnuFSTCross.Checked = False
    mnuFSTDiagCross.Checked = False
    mnuFSTDiagLineRL.Checked = False
    mnuFSTDiagLineLR.Checked = False
    mnuFSTHorzLine.Checked = True
    mnuFSTVertLine.Checked = False
End Sub

Private Sub mnuFSTVertLine_Click()
    propFillStyle = 3
    mnuFSTFilled.Checked = False
    mnuFSTCross.Checked = False
    mnuFSTDiagCross.Checked = False
    mnuFSTDiagLineRL.Checked = False
    mnuFSTDiagLineLR.Checked = False
    mnuFSTHorzLine.Checked = False
    mnuFSTVertLine.Checked = True
End Sub

Private Sub mnuGetPicInfo_Click()
    frmProperties.Show 1, Me
End Sub

Private Sub mnuGreyscale_Click()
    Dim c As Integer
    
    Call Save
    For i = 0 To curX
        For j = 0 To curY
            tCOl = GetPixel(picMain.hdc, i, j)
            r = tCOl Mod 256
            g = (tCOl \ 256) Mod 256
            b = tCOl \ 256 \ 256
            c = r * 0.3 + g * 0.59 + b * 0.11
            SetPixel picMain.hdc, i, j, RGB(c, c, c)
        Next j
        ctlProgBar.Value = i * 100 \ (curX - 1)
    Next i
    picMain.Refresh
    ctlProgBar.Value = 0
End Sub

Private Sub mnuHammer_Click()
    curTool = Tools.EfHammer
    Call RenPanel
End Sub

Private Sub mnuHeat_Click()
    Call Save
    For i = 0 To curX
        For j = 0 To curY
            tCOl = GetPixel(picMain.hdc, i, j)
            r = tCOl Mod 256
            g = (tCOl \ 256) Mod 256
            b = tCOl \ 256 \ 256
            
            r = Abs(((r ^ 2) / ((b + g) + 10)) * 128)
            b = Abs(((b ^ 2) / ((g + r) + 10)) * 128)
            g = Abs(((g ^ 2) / ((r + b) + 10)) * 128)
nOK:
            If r > 32767 Then
                r = r - 32767
            ElseIf g > 32767 Then
                g = g - 32767
            ElseIf b > 32767 Then
                b = b - 32767
            End If
            If r > 32767 Or g > 32767 Or b > 32767 Then
                GoTo nOK
            End If
            SetPixel picMain.hdc, i, j, RGB(r, g, b)
        Next j
        ctlProgBar.Value = i * 100 \ (curX - 1)
    Next i
    ctlProgBar.Value = 0
    picMain.Refresh
End Sub

Private Sub mnuHook_Click()
    curTool = Tools.EfHook
    Call RenPanel
End Sub

Private Sub mnuHorzLine_Click()
    curTool = Tools.HorzLine
    cboEffects.ListIndex = 0
    Call RenPanel
End Sub

Private Sub mnuIce_Click()
    Call Save
    For i = 0 To curX
        For j = 0 To curY
            tCOl = GetPixel(picMain.hdc, i, j)
            r = tCOl Mod 256
            g = (tCOl \ 256) Mod 256
            b = tCOl \ 256 \ 256
            r = Abs((r - g - b) * 1.5)
            g = Abs((g - b - r) * 1.5)
            b = Abs((b - r - g) * 1.5)
            SetPixel picMain.hdc, i, j, RGB(r, g, b)
        Next j
        ctlProgBar.Value = i * 100 \ (curX - 1)
    Next i
    picMain.Refresh
    ctlProgBar.Value = 0
End Sub

Private Sub mnuInvert_Click()
    Dim tH As Long, tW As Long
    
    Call Save
    picInvert.Picture = picMain.Image
    tH = picInvert.Height
    tW = picInvert.Width
    picMain.PaintPicture picInvert.Picture, 0, 0, tW, tH, 0, 0, tW, tH, vbNotSrcCopy
End Sub

Private Sub mnuNight_Click()
    Call Save
    For i = 0 To curX
        For j = 0 To curY
            tCOl = GetPixel(picMain.hdc, i, j)
            r = tCOl Mod 256
            g = (tCOl \ 256) Mod 256
            b = tCOl \ 256 \ 256
            r = Abs(r * r) / 256
            g = Abs(g * g) / 256
            b = Abs(b * b) / 256
            SetPixel picMain.hdc, i, j, RGB(r, g, b)
        Next j
        ctlProgBar.Value = i * 100 \ (curX - 1)
    Next i
    picMain.Refresh
    ctlProgBar.Value = 0
End Sub

Private Sub mnuNormPolygon_Click()
    Call mnuPolygon_Click
End Sub

Private Sub mnuOnPic_Click()
    curTool = Tools.pinp
    cboEffects.ListIndex = 0
    Call RenPanel
End Sub

Private Sub mnuPaste_Click()
    If Clipboard.GetFormat(vbCFBitmap) Then
        picMain.Picture = Clipboard.GetData(vbCFBitmap)
        Call PrepPic
    End If
End Sub

Private Sub mnuPencil_Click()
    curTool = Tools.Pencil
    cboEffects.ListIndex = 0
    Call RenPanel
End Sub

Private Sub mnuPolygon_Click()
    curTool = Tools.Polygon
    cboEffects.ListIndex = 0
    Call RenPanel
End Sub

Private Sub mnuRect_Click()
    curTool = Tools.Rect
    cboEffects.ListIndex = 0
    Call RenPanel
End Sub

Private Sub mnuRects_Click()
    Dim tColR1 As Long, tColR2 As Long, tColR3 As Long, tColR4 As Long, tColR5 As Long
    
    Call Save
    For i = 0 To curX
        For j = 0 To curY
            tColR1 = GetPixel(picMain.hdc, i, j)
            tColR2 = GetPixel(picMain.hdc, i + 1, j)
            tColR3 = GetPixel(picMain.hdc, i - 1, j)
            tColR4 = GetPixel(picMain.hdc, i, j + 1)
            tColR5 = GetPixel(picMain.hdc, i, j - 1)
            SetPixel picMain.hdc, i, j, (Abs(tColR1) - (Abs(tColR2 + tColR3 + tColR4 + tColR5) / 4))
        Next j
        ctlProgBar.Value = i * 100 \ (curX - 1)
    Next i
    ctlProgBar.Value = 0
    picMain.Refresh
End Sub

Private Sub mnuRepColor_Click()
    curTool = Tools.RepCol
    Call RenPanel
End Sub

Private Sub mnuRotate_Click()
    frmRotate.Show 1, Me
End Sub

Private Sub mnusCircle_Click()
    Call mnuCircle_Click
End Sub

Private Sub mnusFCircle_Click()
    Call mnuFilledCircle_Click
End Sub

Private Sub mnusFRect_Click()
    Call mnuFilledRect_Click
End Sub

Private Sub mnuSharpen_Click()
    Dim tColA As Long
    Dim r1 As Long, g1 As Long, b1 As Long
    
    Call Save
    For i = 1 To curX
        For j = 1 To curY
            tCOl = GetPixel(picMain.hdc, i, j)
            r = tCOl Mod 256
            g = (tCOl \ 256) Mod 256
            b = tCOl \ 256 \ 256
            tColA = GetPixel(picMain.hdc, i - 1, j - 1)
            r1 = tColA Mod 256
            g1 = (tColA \ 256) Mod 256
            b1 = tColA \ 256 \ 256
            r = r + 0.5 * (r - r1)
            g = g + 0.5 * (g - g1)
            b = b + 0.5 * (b - b1)
            
            If r > 255 Then r = 255
            If r < 0 Then r = 0
            If g > 255 Then g = 255
            If g < 0 Then g = 0
            If b > 255 Then b = 255
            If b < 0 Then b = 0

            SetPixel picMain.hdc, i, j, RGB(r, g, b)
        Next j
        ctlProgBar.Value = i * 100 \ (curX - 1)
    Next i
    picMain.Refresh
    ctlProgBar.Value = 0
End Sub

Private Sub mnusRect_Click()
    Call mnuRect_Click
End Sub

Private Sub mnuStar_Click()
    curTool = Tools.Star
    cboEffects.ListIndex = 0
    Call RenPanel
End Sub

Private Sub mnuStraightLine_Click()
    curTool = Tools.StraightLine
    cboEffects.ListIndex = 0
    Call RenPanel
End Sub

Private Sub mnuStrange_Click()
    Call Save
    For i = 0 To curX
        For j = 0 To curY
            tCOl = GetPixel(picMain.hdc, i, j)
            r = tCOl Mod 256
            g = (tCOl \ 256) Mod 256
            b = tCOl \ 256 \ 256
            If (g = 0) Or (b = 0) Then
                g = 1
                b = 1
            End If
            r = Abs(Sin(Atn(g / b)) * 125 + 20)
            g = Abs(Sin(Atn(r / b)) * 125 + 20)
            b = Abs(Sin(Atn(r / g)) * 125 + 20)
            SetPixel picMain.hdc, i, j, RGB(r, g, b)
        Next j
        ctlProgBar.Value = i * 100 \ (curX - 1)
    Next i
    picMain.Refresh
    ctlProgBar.Value = 0
End Sub

Private Sub mnuText_Click()
    curTool = Tools.InsText
    cboEffects.ListIndex = 0
    Call RenPanel
End Sub

Private Sub mnuUDefPolygon_Click()
    Call mnuUserDefPolygon_Click
End Sub

Private Sub mnuUserDefPolygon_Click()
    curTool = Tools.UdefPoly
    cboEffects.ListIndex = 0
    Call RenPanel
End Sub

Private Sub mnuVertLine_Click()
    curTool = Tools.VertLine
    cboEffects.ListIndex = 0
    Call RenPanel
End Sub

Private Sub mnuWave_Click()
    Dim coli() As Long, posy() As Double
    
    Call Save
    ReDim coli(curX, curY)
    ReDim posy(curX, curY)
    
    For i = 0 To curX
        For j = 0 To curY
            coli(i, j) = GetPixel(picMain.hdc, i, j)
            posy(i, j) = Sin(i) * 6 + (j - 3)
        Next j
        ctlProgBar.Value = i * 100 \ (curX - 1)
    Next i
    For i = 0 To curX
        For j = 0 To curY
            SetPixel picMain.hdc, i, posy(i, j), coli(i, j)
        Next j
        ctlProgBar.Value = i * 100 \ (curX - 1)
    Next i
    picMain.Refresh
    ctlProgBar.Value = 0
End Sub

Private Sub mnuWithDialog_Click()
    Dim tColA As Long
    
    On Error Resume Next
    Err.Clear
    ctlCommDiag.CancelError = True
    ctlCommDiag.ShowColor
    If Not (Err.Number > 0) Then
        tColA = ctlCommDiag.Color
        picForeCol.BackColor = tColA
        ForeCol = tColA
    End If
End Sub

Private Sub picBack_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        picMain.Width = X
        picMain.Height = Y
        picUndo.Width = X
        picUndo.Height = Y
        picMain.Picture = picMain.Image
        IsBitmap = False
        Call PrepPic
    End If
End Sub

Private Sub picBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ctlStatBar.Panels(1).Text = vbNullString
End Sub

Private Sub picBackCol_Click()
    Dim tColA As Long
    
    On Error Resume Next
    Err.Clear
    ctlCommDiag.CancelError = True
    ctlCommDiag.ShowColor
    If Not (Err.Number > 0) Then
        tColA = ctlCommDiag.Color
        picBackCol.BackColor = tColA
        FillCol = tColA
    End If
End Sub

Private Sub picForeCol_Click()
    Dim tColA As Long
    
    On Error Resume Next
    Err.Clear
    ctlCommDiag.CancelError = True
    ctlCommDiag.ShowColor
    If Not (Err.Number > 0) Then
        tColA = ctlCommDiag.Color
        picForeCol.BackColor = tColA
        ForeCol = tColA
    End If
End Sub

Private Sub picMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim rv As Long, tPoints(0 To 2) As POINTAPI
    
    Select Case Button
        Case 2:
            If curTool = Tools.pinp Then
                FillCol = GetPixel(picMain.hdc, X, Y)
                picBackCol.BackColor = FillCol
                picMain.FillColor = FillCol
            ElseIf curTool = Tools.btCopy Or curTool = Tools.btCut Then
                If Marked Then
                  MarkRect picMain, StartX, StartY, OldX, OldY
                  Marked = False
                  NowMove = False
                End If
            Else
                PopupMenu mnuPopup
            End If
        Case 1:
            Call Save
            Select Case curTool
                Case Tools.EfHammer:
                    Call DrawBCircle(X, Y)
                
                Case Tools.EfHook:
                    Call DrawBCircle(X, Y)
                    
                Case Tools.Pencil:
                    sX = X
                    sY = Y
                    
                Case Tools.FilledCircle:
                    tX = X
                    tY = Y
                    
                Case Tools.FilledRect:
                    tX = X
                    tY = Y
                    
                Case Tools.FillRgn:
                    Call Filling(picMain.Point(X, Y), propFillStyle, X, Y)
                    
                Case Tools.InsText:
                    sX = X
                    sY = Y
                    frmInsText.Show 1, Me

                Case Tools.pinp:
                    ForeCol = GetPixel(picMain.hdc, X, Y)
                    picForeCol.BackColor = ForeCol
                    picMain.ForeColor = ForeCol
                    
                Case Tools.Polygon:
                    If Not stat Then
                        stat = True
                        cuX = X
                        cuY = Y
                    End If
                    sX = X
                    sY = Y
                    If Shift = 1 Then
                        picMain.Line (wX1, wY1)-(cuX, cuY), ForeCol
                        stat1 = False
                        stat = False
                        Exit Sub
                    End If
                    If Not stat1 Then
                        picMain.PSet (X, Y), ForeCol
                        wX1 = X
                        wY1 = Y
                        stat1 = True
                    Else
                        picMain.Line (wX1, wY1)-(X, Y), ForeCol
                    End If
                    Exit Sub
                    
                Case Tools.Rect:
                    tX = X
                    tY = Y
                    
                Case Tools.Star:
                    sX = X
                    sY = Y

                Case Tools.StraightLine:
                    tX = X
                    tY = Y
                    
                Case Tools.tCircle:
                    tX = X
                    tY = Y
                    
                Case Tools.tErase:
                    sX = X
                    sY = Y
                    
                Case Tools.UdefPoly:
                    sX = X
                    sY = Y
                    tX = X
                    tY = Y
                    
                Case Tools.RepCol:
                    tCOl = GetPixel(picMain.hdc, X, Y)
                    For i = 0 To curX
                        For j = 0 To curY
                            If (GetPixel(picMain.hdc, i, j) = tCOl) Then
                                SetPixel picMain.hdc, i, j, ForeCol
                            End If
                        Next j
                        ctlProgBar.Value = i * 100 \ (curX - 1)
                    Next i
                    ctlProgBar.Value = 0
                    picMain.Refresh
                
                Case Tools.btCut:
                    MoveX = X
                    MoveY = Y
                    Select Case Marked
                        Case True
                            NowMove = PtInRegion(X, Y, StartX, StartY, OldX, OldY)
                        Case Else
                            NowMove = False
                            StartX = X
                            StartY = Y
                            OldX = X
                            OldY = Y
                            MoveX = X
                            MoveY = Y
                    End Select
                
                Case Tools.btCopy:
                    MoveX = X
                    MoveY = Y
                    Select Case Marked
                        Case True
                            NowMove = PtInRegion(X, Y, StartX, StartY, OldX, OldY)
                        Case Else
                            NowMove = False
                            StartX = X
                            StartY = Y
                            OldX = X
                            OldY = Y
                            MoveX = X
                            MoveY = Y
                    End Select
                    
            End Select
    End Select
End Sub

Private Sub picMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim tPosA As Long, tPosB As Long
    
    ctlStatBar.Panels(1).Text = X & " X " & Y
    If IsCurOK Then
        Select Case curTool
            Case Tools.EfHammer:
                picMain.MousePointer = 99
                picMain.MouseIcon = LoadPicture(App.Path & "\Hammer.cur")
            Case Tools.EfHook:
                picMain.MousePointer = 99
                picMain.MouseIcon = LoadPicture(App.Path & "\Hook.cur")
            Case Tools.Brush:
                picMain.MousePointer = 99
                picMain.MouseIcon = LoadPicture(App.Path & "\Brush.cur")
            Case Tools.tErase:
                picMain.MousePointer = 99
                picMain.MouseIcon = LoadPicture(App.Path & "\Erase.cur")
            Case Tools.InsText:
                picMain.MousePointer = 99
                picMain.MouseIcon = LoadPicture(App.Path & "\Text.cur")
            Case Tools.FillRgn:
                picMain.MousePointer = 99
                picMain.MouseIcon = LoadPicture(App.Path & "\FillRgn.cur")
            Case Tools.pinp:
                picMain.MousePointer = 99
                picMain.MouseIcon = LoadPicture(App.Path & "\pinp.cur")
            Case Tools.RepCol:
                picMain.MousePointer = 99
                picMain.MouseIcon = LoadPicture(App.Path & "\pinp.cur")
            Case Tools.btCopy:
                picMain.MousePointer = 99
                picMain.MouseIcon = LoadPicture(App.Path & "\Copy.cur")
            Case Tools.btCut:
                picMain.MousePointer = 99
                picMain.MouseIcon = LoadPicture(App.Path & "\Cut.cur")
            Case Else:
                picMain.MousePointer = 99
                picMain.MouseIcon = LoadPicture(App.Path & "\NormDraw.cur")
        End Select
    End If
    If Button = 1 Then
        Select Case curTool
            Case Tools.Brush:
                For i = 0 To 25
                    tPosA = Int(Rnd * 14 - 7)
                    tPosB = Int(Rnd * 14 - 7)
                    SetPixel picMain.hdc, X + tPosA, Y + tPosB, ForeCol
                Next i
                picMain.Refresh
                
            Case Tools.Cross:
                picMain.Line (X - 5, Y)-(X + 5, Y), ForeCol
                picMain.Line (X, Y - 5)-(X, Y + 5), ForeCol
                
            Case Tools.DiagCross:
                picMain.Line (X - 5, Y - 5)-(X + 5, Y + 5), ForeCol
                picMain.Line (X + 5, Y - 5)-(X - 5, Y + 5), ForeCol
                
            Case Tools.DiagLineLR:
                picMain.Line (X, Y)-(X + 5, Y + 5), ForeCol
            
            Case Tools.DiagLineRL:
                picMain.Line (X, Y)-(X - 5, Y + 5), ForeCol
                
            Case Tools.EfHook:
                Call DrawBCircle(X, Y)
            
            Case Tools.UdefPoly:
                picMain.Line (sX, sY)-(X, Y), ForeCol
                sX = X
                sY = Y
                
            Case Tools.HorzLine:
                picMain.Line (X - 5, Y)-(X + 5, Y), ForeCol
                
            Case Tools.Pencil:
                picMain.Line (sX, sY)-(X, Y), ForeCol
                sX = X
                sY = Y
                
            Case Tools.Star:
                picMain.Line (sX, sY)-(X, Y), ForeCol
                
            Case Tools.tErase:
                picMain.Line (sX, sY)-(X, Y), vbWhite
                sX = X
                sY = Y
                
            Case Tools.VertLine:
                picMain.Line (X, Y - 5)-(X, Y + 5), ForeCol
            
            Case Tools.btCopy:
                Select Case Marked
                Case True
                    MarkRect picMain, StartX, StartY, OldX, OldY
                    Select Case NowMove
                        Case True
                            StartX = StartX + (X - MoveX)
                            StartY = StartY + (Y - MoveY)
                            OldX = OldX + (X - MoveX)
                            OldY = OldY + (Y - MoveY)
                            MoveX = X
                            MoveY = Y
                        Case Else
                            OldX = X
                            OldY = Y
                    End Select
                End Select
                MarkRect picMain, StartX, StartY, OldX, OldY
                Marked = True
            
            Case Tools.btCut:
                Select Case Marked
                Case True
                    MarkRect picMain, StartX, StartY, OldX, OldY
                    Select Case NowMove
                        Case True
                            StartX = StartX + (X - MoveX)
                            StartY = StartY + (Y - MoveY)
                            OldX = OldX + (X - MoveX)
                            OldY = OldY + (Y - MoveY)
                            MoveX = X
                            MoveY = Y
                        Case Else
                            OldX = X: OldY = Y
                    End Select
                End Select
                MarkRect picMain, StartX, StartY, OldX, OldY
                Marked = True
            
            End Select
    End If
    wX = X
    wY = Y
End Sub

Private Sub picMain_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim r As Double, bCCT(2) As Boolean, lCCHeight As Long, lCCWidth As Long
    Dim lFst As Long, lCol As Long
    
    If Button = 1 Then
        Select Case curTool
            Case Tools.UdefPoly:
                picMain.Line (tX, tY)-(X, Y), ForeCol
                
            Case Tools.tCircle:
                picMain.FillStyle = 1
                r = Sqr((tX - X) * (tX - X) + (tY - Y) * (tY - Y))
                picMain.Circle (tX, tY), r, ForeCol
                
            Case Tools.Rect:
                picMain.FillStyle = 1
                picMain.Line (tX, tY)-(X, Y), ForeCol, B
                
            Case Tools.FilledCircle:
                picMain.FillStyle = 0
                picMain.FillColor = FillCol
                r = Sqr((tX - X) * (tX - X) + (tY - Y) * (tY - Y))
                picMain.Circle (tX, tY), r, FillCol
                picMain.FillStyle = 1
                
            Case Tools.FilledRect:
                picMain.FillStyle = 0
                picMain.FillColor = FillCol
                picMain.Line (tX, tY)-(X, Y), FillCol, BF
                picMain.FillStyle = 1
                
            Case Tools.StraightLine:
                picMain.Line (tX, tY)-(X, Y), ForeCol
            
            Case Tools.Polygon:
                picMain.Line (sX, sY)-(wX, wY), ForeCol
                wX1 = X
                wY1 = Y
                
             Case Tools.btCopy:
                If Shift = vbAltMask Then
                    bCCT(0) = (OldY < StartY And OldX < StartX)
                    bCCT(1) = (OldY > StartY And OldX < StartX)
                    bCCT(2) = (OldY > StartY And OldX > StartX)
                    
                    If bCCT(0) Then
                        lCCHeight = StartY - OldY
                        lCCWidth = StartX - OldX
                    ElseIf bCCT(1) Then
                        lCCHeight = OldY - StartY
                        lCCWidth = StartX - OldX
                    ElseIf bCCT(2) Then
                        lCCHeight = StartY - OldY
                        lCCWidth = OldX - StartX
                    Else
                        lCCHeight = OldY - StartY
                        lCCWidth = OldX - StartX
                    End If
                    
                    lCCHeight = Abs(lCCHeight)
                    lCCWidth = Abs(lCCWidth)
                    picCutCopy.Height = lCCHeight
                    picCutCopy.Width = lCCWidth
                    picCutCopy.PaintPicture picMain.Image, 0, 0, picCutCopy.Width, picCutCopy.Height, StartX, StartY, picCutCopy.Width, picCutCopy.Height, vbNotSrcCopy
                    Clipboard.Clear
                    Clipboard.SetData picCutCopy.Image, vbCFBitmap
                    If Marked Then
                        MarkRect picMain, StartX, StartY, OldX, OldY
                        Marked = False
                        NowMove = False
                    End If
                    picCutCopy.Picture = LoadPicture
                End If
                
             Case Tools.btCut:
                If Shift = vbAltMask Then
                    bCCT(0) = (OldY < StartY And OldX < StartX)
                    bCCT(1) = (OldY > StartY And OldX < StartX)
                    bCCT(2) = (OldY > StartY And OldX > StartX)
                    
                    If bCCT(0) Then
                        lCCHeight = StartY - OldY
                        lCCWidth = StartX - OldX
                    ElseIf bCCT(1) Then
                        lCCHeight = OldY - StartY
                        lCCWidth = StartX - OldX
                    ElseIf bCCT(2) Then
                        lCCHeight = StartY - OldY
                        lCCWidth = OldX - StartX
                    Else
                        lCCHeight = OldY - StartY
                        lCCWidth = OldX - StartX
                    End If
                    
                    lCCHeight = Abs(lCCHeight)
                    lCCWidth = Abs(lCCWidth)
                    picCutCopy.Height = lCCHeight
                    picCutCopy.Width = lCCWidth
                    picCutCopy.PaintPicture picMain.Image, 0, 0, picCutCopy.Width, picCutCopy.Height, StartX, StartY, picCutCopy.Width, picCutCopy.Height, vbNotSrcCopy
                    Clipboard.Clear
                    Clipboard.SetData picCutCopy.Image, vbCFBitmap
                    If Marked Then
                        lFst = picMain.FillStyle
                        lCol = picMain.FillColor
                        picMain.FillStyle = 0
                        picMain.FillColor = vbWhite
                        picMain.Line (StartX, StartY)-(OldX, OldY), vbWhite, BF
                        picMain.FillColor = lCol
                        picMain.FillStyle = lFst
                        Marked = False
                        NowMove = False
                    End If
                    picCutCopy.Picture = LoadPicture
                End If
                
        End Select
    End If
End Sub

Private Sub ctlTools_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Marked Then
        MarkRect picMain, StartX, StartY, OldX, OldY
        Marked = False
        NowMove = False
    End If
    Select Case Button.Key
        Case "cmPencil":
            Call mnuPencil_Click
        Case "cmStar":
            Call mnuStar_Click
        Case "cmFillRgn":
            Call mnuFillRgn_Click
        Case "cmCircRect":
            PopupMenu mnuShapes
        Case "cmStLine":
            Call mnuStraightLine_Click
        Case "cmBrush":
            Call mnuBrush_Click
        Case "cmInsertText":
            Call mnuText_Click
        Case "cmHorzLine":
            Call mnuHorzLine_Click
        Case "cmVertLine"
            Call mnuVertLine_Click
        Case "cmDiagLineRL":
            Call mnuDiagLineRL_Click
        Case "cmDiagLineLR":
            Call mnuDiagLineLR_Click
        Case "cmCross":
            Call mnuCross_Click
        Case "cmDiagCross":
            Call mnuDiagCross_Click
        Case "cmPolygon":
            PopupMenu mnuPolGon
        Case "cmGetCol":
            PopupMenu mnuChooseCol2
        Case "cmErase":
            Call mnuErase_Click
    End Select
End Sub

Private Sub scrHorz_Change()
    picMain.Left = -scrHorz.Value
End Sub

Private Sub scrVert_Change()
    picMain.Top = -scrVert.Value
End Sub
