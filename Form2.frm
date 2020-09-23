VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   4740
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6105
   LinkTopic       =   "Form2"
   ScaleHeight     =   316
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   407
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      Height          =   255
      Left            =   0
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   189
      TabIndex        =   1
      Top             =   0
      Width           =   2895
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Error"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   45
         TabIndex        =   2
         Top             =   -15
         Width           =   5655
      End
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   2895
      Left            =   1440
      ScaleHeight     =   193
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   257
      TabIndex        =   0
      Top             =   1320
      Width           =   3855
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim XX As Integer
Dim YY As Integer
Dim JB As Byte
Sub OpenPicture(Files As String)
On Error GoTo kakas
Picture1.Picture = LoadPicture(Files)
SizeIt
Label1.Caption = Files
Exit Sub
kakas:
Label1.Caption = "Error"
MsgBox "Could not open picture!", vbExclamation
Exit Sub
End Sub


Sub SizeIt()
Picture1.Left = (Form2.ScaleWidth / 2) - (Picture1.ScaleWidth / 2)
Picture1.Top = (Form2.ScaleHeight / 2) - (Picture1.ScaleHeight / 2)


End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
Unload Me
End If
End Sub


Private Sub Form_Resize()
Label1.Width = Form2.ScaleWidth
Picture2.Width = Form2.ScaleWidth
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
Unload Me
End If
If Button = vbRightButton Then
JB = 1
XX = X
YY = Y
End If
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton And JB = 1 Then
Picture1.Left = Picture1.Left + (X - XX)
Picture1.Top = Picture1.Top + (Y - YY)
Picture1.Refresh
End If
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

JB = 0

End Sub


