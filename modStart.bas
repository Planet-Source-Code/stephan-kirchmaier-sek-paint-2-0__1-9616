Attribute VB_Name = "modStart"
Option Explicit
'******************************************************'
'*                 SEK - Paint 2.0                    *'
'*           Made by Stephan Kirchmaier               *'
'*    e-mail: VB_Empire@gmx.at | www.vb-empire.de.vu  *'
'******************************************************'
Sub Main()
    Dim iFreeNum As Integer
    
    On Error Resume Next
    
    If App.PrevInstance Then
        MsgBox "The App is already started!", vbCritical, "Error"
        Exit Sub
    End If
    
    IsBitmap = False
    If Command = vbNullString Then
        frmMain.Show
    Else
        If FileExist(Command) Then
            iFreeNum = FreeFile()
            Open Command For Binary Access Read Lock Write As iFreeNum
                Get iFreeNum, 1, iIsBmp
            Close iFreeNum
            IsBitmap = (iIsBmp = 19778)
            sBmpPath = Command
            frmMain.picMain.Picture = LoadPicture(Command)
            frmMain.Show
        End If
    End If
End Sub
