VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Screen Shot Taker"
   ClientHeight    =   6675
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8085
   LinkTopic       =   "Form1"
   ScaleHeight     =   445
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   539
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   375
      Left            =   6480
      TabIndex        =   3
      Top             =   6240
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save to File..."
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   6240
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Take Screen Shot"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   6240
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   6135
      Left            =   0
      ScaleHeight     =   405
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   533
      TabIndex        =   0
      Top             =   0
      Width           =   8055
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   600
         Top             =   3360
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DefaultExt      =   ".bmp"
         Filter          =   "(*.bmp;*.jpg)|*.bmp;*.jpg"
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Note:  The whole picture may not be visible, but it will be saved."
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   6240
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Me.Visible = False
hold (0.5)
Form2.Show
End Sub

Private Sub Command2_Click()
CommonDialog1.ShowSave
If CommonDialog1.FileName = "" Then Exit Sub
SavePicture Picture1.Picture, CommonDialog1.FileName
End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Form_Resize()
If Form1.WindowState = 1 Then Exit Sub

Picture1.Width = Form1.ScaleWidth - 4
Picture1.Height = Form1.ScaleHeight - 40
Command1.Left = 4
Command1.Top = Form1.ScaleHeight - (Command1.Height + 4)
Command2.Left = Command1.Left + Command1.Width + 4
Command2.Top = Command1.Top
Command3.Left = Form1.ScaleWidth - Command3.Width - 4
Command3.Top = Command1.Top
Label1.Top = Command1.Top
Label1.Left = Command2.Left + Command2.Width + 10
Picture1.Refresh


End Sub
