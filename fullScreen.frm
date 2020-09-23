VERSION 5.00
Begin VB.Form Form2 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   DrawStyle       =   2  'Dot
   ForeColor       =   &H000080FF&
   LinkTopic       =   "Form2"
   MousePointer    =   2  'Cross
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_DblClick()
End
End Sub

Private Sub Form_Load()

dskWnd& = GetDesktopWindow
dskDC& = GetDC(dskWnd&)

BitBlt Form2.hDC, 0&, 0&, Screen.Width, Screen.Height, dskDC&, 0&, 0&, SRCCOPY
Form2.Picture = Form2.Image
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
X_1 = X
Y_1 = Y
lastX = X
lastY = Y
Me.DrawMode = 7

YN = 1
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If YN <> 1 Then Exit Sub
Me.Line (X_1, Y_1)-(lastX, lastY), , B
Me.Line (X_1, Y_1)-(X, Y), , B
lastX = X
lastY = Y


End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
YN = 0
Form1.Visible = True
Form1.Show
Form1.Picture1.Cls
If X_1 = lastX Or Y_1 = lastY Then
Unload Me
Exit Sub
End If

If X_1 > lastX And Y_1 > lastY Then
Form1.Picture1.Width = X_1 - lastX
Form1.Picture1.Height = Y_1 - lastY
Form1.Picture1.PaintPicture Me.Picture, 0, 0, X_1 - lastX, Y_1 - lastY, lastX, lastY, X_1 - lastX, Y_1 - lastY
ElseIf X_1 > lastX Then
Form1.Picture1.Width = X_1 - lastX
Form1.Picture1.Height = lastY - Y_1
Form1.Picture1.PaintPicture Me.Picture, 0, 0, X_1 - lastX, lastY - Y_1, lastX, Y_1, X_1 - lastX, lastY - Y_1
ElseIf Y_1 > lastY Then
Form1.Picture1.Width = lastX - X_1
Form1.Picture1.Height = Y_1 - lastY
Form1.Picture1.PaintPicture Me.Picture, 0, 0, lastX - X_1, Y_1 - lastY, X_1, lastY, lastX - X_1, Y_1 - lastY
Else
Form1.Picture1.Width = lastX - X_1
Form1.Picture1.Height = lastY - Y_1
Form1.Picture1.PaintPicture Me.Picture, 0, 0, lastX - X_1, lastY - Y_1, X_1, Y_1, lastX - X_1, lastY - Y_1
End If
Form1.Picture1.Picture = Form1.Picture1.Image

Form1.Picture1.Height = Form1.ScaleHeight - 40
Form1.Picture1.Width = Form1.ScaleWidth - 4
Unload Me
End Sub
