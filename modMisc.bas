Attribute VB_Name = "modMisc"
Option Explicit

Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Declare Function ExtFloodFill Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long
Declare Function PolyBezierTo Lib "gdi32" (ByVal hdc As Long, lppt As POINTAPI, ByVal cCount As Long) As Long
Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As Any) As Long
Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal Height As Long, _
                                                             ByVal Width As Long, _
                                                             ByVal Escapement As Long, _
                                                             ByVal Orientation As Long, _
                                                             ByVal fontwidth As Long, _
                                                             ByVal Italic As Long, _
                                                             ByVal Unerline As Long, _
                                                             ByVal StrikeOut As Long, _
                                                             ByVal CharSet As Long, _
                                                             ByVal OutputPrecision As Long, _
                                                             ByVal ClipPrecision As Long, _
                                                             ByVal Quality As Long, _
                                                             ByVal PitchAndFamily As Long, _
                                                             ByVal FontName As String) As Long
Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Global sinTab(360) As Double, cosTab(360) As Double
Global curTool As Integer
Global ForeCol As Long, FillCol As Long
Global i As Long, j As Long
Global curX As Long, curY As Long
Global r As Long, g As Long, b As Long
Global tCOl As Long
Global lArrCol() As Long
Global sX As Long, sY As Long
Global tX As Long, tY As Long
Global propFillStyle As Long
Global cuX As Long, cuY As Long
Global stat As Long, stat1 As Long
Global wX1 As Long, wY1 As Long
Global wX As Long, wY As Long
Global IsBitmap As Boolean
Global iIsBmp As Integer
Global sBmpPath As String
Global StartX, StartY
Global Marked, OldX, OldY
Global MoveX, MoveY, NowMove
Global IsCurOK As Boolean

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Enum Tools
    Pencil = 1
    Star = 2
    HorzLine = 3
    VertLine = 4
    Cross = 5
    DiagCross = 6
    DiagLineLR = 7
    DiagLineRL = 8
    UdefPoly = 9
    InsText = 10
    StraightLine = 11
    Brush = 12
    Polygon = 13
    Rect = 14
    tCircle = 15
    FilledRect = 16
    FilledCircle = 17
    FillRgn = 18
    tErase = 19
    EfHammer = 20
    EfHook = 21
    btCopy = 22
    btCut = 23
    pinp = 24
    RepCol = 25
End Enum

Public Const Pi = 3.14159265359
Public Const BRadius As Integer = 120
Public Const FW_BOLD = 700
Public Const FW_EXTRABOLD = 800
Public Const FW_EXTRALIGHT = 200
Public Const FW_HEAVY = 900
Public Const FW_LIGHT = 300
Public Const FW_MEDIUM = 500
Public Const FW_NORMAL = 400
Public Const FW_SEMIBOLD = 600
Public Const FW_THIN = 100
Public Const DEFAULT_CHARSET = 1
Public Const OUT_DEFAULT_PRECIS = 0
Public Const CLIP_DEFAULT_PRECIS = 0
Public Const PROOF_QUALITY = 2
Public Const FF_DONTCARE = 0

Public Function FileExist(sFileN As String) As Boolean
    Dim tmpRv As Long
    
    On Error Resume Next
    tmpRv = GetAttr(sFileN)
    FileExist = Not CBool(Err)
End Function

Public Sub RenPanel()
    Dim sText As String
    
    Select Case curTool
        Case 1: sText = "Pencil"
        Case 2: sText = "Star"
        Case 3: sText = "Horizontal Line"
        Case 4: sText = "Vertical Line"
        Case 5: sText = "Cross"
        Case 6: sText = "Diagonal Cross"
        Case 7: sText = "Diagonal Line (\)"
        Case 8: sText = "Diagonal Line (/)"
        Case 9: sText = "Userdefined Polygon"
        Case 10: sText = "Text"
        Case 11: sText = "Straight Line"
        Case 12: sText = "Brush"
        Case 13: sText = "Polygon"
        Case 14: sText = "Rect"
        Case 15: sText = "Circle"
        Case 16: sText = "Filled Rect"
        Case 17: sText = "Filled Circle"
        Case 18: sText = "Fill Region"
        Case 19: sText = "Erase"
        Case 20: sText = "Hammer"
        Case 21: sText = "Hook"
        Case 22: sText = "Copy"
        Case 23: sText = "Cut"
        Case 24: sText = "Choose Color"
        Case 25: sText = "Replace Color"
    End Select
    frmMain.ctlStatBar.Panels(2).Text = sText
End Sub

Public Sub Save()
    frmMain.picUndo.Picture = frmMain.picMain.Image
End Sub

Public Sub Undo()
    frmMain.picMain.Picture = frmMain.picUndo.Image
End Sub

Public Sub Filling(col As Long, ByVal FStyle As Long, X, Y)
    Dim rv As Long
    
    frmMain.picMain.FillStyle = FStyle
    frmMain.picMain.FillColor = FillCol
    rv = ExtFloodFill(frmMain.picMain.hdc, X, Y, col, 1)
    frmMain.picMain.FillStyle = 1
End Sub

Public Sub DrawBCircle(ByVal X As Long, ByVal Y As Long)
    Dim rv As Long, tPoints(0 To 2) As POINTAPI
    
    Call Save
    For i = 0 To 360
        tPoints(0).X = cosTab(i) * BRadius + X
        tPoints(0).Y = sinTab(i) * BRadius + Y
        tPoints(1).X = X
        tPoints(1).Y = Y
        tPoints(2).X = X
        tPoints(2).Y = Y
                
        frmMain.picMain.ForeColor = Abs(GetPixel(frmMain.picMain.hdc, (Cos(i) * (BRadius / 2.23) + X), (Sin(i) * (BRadius / 2.23) + Y)))
        rv = MoveToEx(frmMain.picMain.hdc, X, Y, 0&)
        rv = PolyBezierTo(frmMain.picMain.hdc, tPoints(0), 3)
    Next i
    frmMain.picMain.Refresh
End Sub

Public Sub PrepPic()
    frmMain.scrHorz.Value = 0
    frmMain.scrHorz.Max = frmMain.picMain.Width - 5
    frmMain.scrVert.Value = 0
    frmMain.scrVert.Max = frmMain.picMain.Height - 5
    curX = frmMain.picMain.Width
    curY = frmMain.picMain.Height
End Sub

Public Sub PrepareImg()
    ReDim lArrCol(2, curX, curY)
    For i = 0 To curX
        For j = 0 To curY
            tCOl = GetPixel(frmMain.picMain.hdc, i, j)
            r = tCOl Mod 256
            g = (tCOl / 256) Mod 256
            b = tCOl / 256 / 256
            lArrCol(0, i, j) = r
            lArrCol(1, i, j) = g
            lArrCol(2, i, j) = b
        Next j
        frmMain.ctlProgBar.Value = i * 100 \ (curX - 1)
    Next i
    frmMain.ctlProgBar.Value = 0
End Sub

Public Sub MarkRect(Pic As PictureBox, X1, Y1, X2, Y2)
    Dim RecMode As Long, RecStyle As Long
    
    RecMode = Pic.DrawMode
    RecStyle = Pic.DrawStyle
    
    Pic.DrawMode = 6
    Pic.DrawStyle = 0
    Pic.Line (X1, Y1)-(X2, Y2), , B
    
    Pic.DrawMode = RecMode
    Pic.DrawStyle = RecStyle
End Sub

Public Function PtInRegion(X, Y, X1, Y1, X2, Y2) As Boolean
    If X1 > X2 Then Swap X1, X2
    If Y1 > Y2 Then Swap Y1, Y2
    
    PtInRegion = (X > X1) And (X < X2) And (Y > Y1) And (Y < Y2)
End Function

Public Sub Swap(Val1 As Variant, Val2 As Variant)
    Dim Rec
    
    Rec = Val1
    Val1 = Val2
    Val2 = Rec
End Sub

Public Sub bmp_rotate(pic1 As PictureBox, pic2 As PictureBox, ByVal theta As Single)
    Dim c1x As Integer
    Dim c1y As Integer
    Dim c2x As Integer
    Dim c2y As Integer
    Dim a As Single
    Dim r As Integer
    Dim p1x As Long
    Dim p1y As Long
    Dim p2x As Long
    Dim p2y As Long
    Dim n As Integer, tVal As Integer
    Dim c0 As Long, c1 As Long, c2 As Long, c3 As Long
    
    c1x = pic1.ScaleWidth / 2
    c1y = pic1.ScaleHeight / 2
    c2x = pic2.ScaleWidth / 2
    c2y = pic2.ScaleHeight / 2

    n = pic2.ScaleWidth
    If n < pic2.ScaleHeight Then n = pic2.ScaleHeight
    n = n / 2 - 1
    For p2x = 0 To n
        For p2y = 0 To n
            If p2x = 0 Then
                a = Pi / 2
            Else
                a = Atn(p2y / p2x)
            End If
            r = Sqr(1 * p2x * p2x + 1 * p2y * p2y)
            p1x = r * Cos(a + theta)
            p1y = r * Sin(a + theta)
            c0 = pic1.Point(c1x + p1x, c1y + p1y)
            c1 = pic1.Point(c1x - p1x, c1y - p1y)
            c2 = pic1.Point(c1x + p1y, c1y - p1x)
            c3 = pic1.Point(c1x - p1y, c1y + p1x)
            If c0 <> -1 Then pic2.PSet (c2x + p2x, c2y + p2y), c0
            If c1 <> -1 Then pic2.PSet (c2x - p2x, c2y - p2y), c1
            If c2 <> -1 Then pic2.PSet (c2x + p2y, c2y - p2x), c2
            If c3 <> -1 Then pic2.PSet (c2x - p2y, c2y + p2x), c3
        Next
        If n = p2x Then
            tVal = 100
        Else
            tVal = (p2x * 100) \ (n - 1)
        End If
        frmRotate.ctlRotProgBar.Value = tVal
    Next
    frmRotate.ctlRotProgBar.Value = 0
End Sub

Public Function IsCurOKFirst() As Boolean
    IsCurOKFirst = FileExist(App.Path & "\Text.cur") And FileExist(App.Path & "\Brush.cur") And FileExist(App.Path & "\Copy.cur") And FileExist(App.Path & "\Cut.cur") And FileExist(App.Path & "\Hammer.cur") And FileExist(App.Path & "\Hook.cur") And FileExist(App.Path & "\NormDraw.cur") And FileExist(App.Path & "\FillRgn.cur") And FileExist(App.Path & "\Pinp.cur") And FileExist(App.Path & "\Erase.cur")
End Function

