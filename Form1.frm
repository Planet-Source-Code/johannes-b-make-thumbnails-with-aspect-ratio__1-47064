VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Thumbnail viewer 1.0 by Johannes B 2003"
   ClientHeight    =   6465
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12015
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   431
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   801
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   0
      TabIndex        =   9
      Top             =   5400
      Width           =   2655
      Begin VB.CheckBox Check1 
         Caption         =   "XP fix"
         Height          =   255
         Left            =   1440
         TabIndex        =   14
         Top             =   360
         Width           =   735
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   360
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Make thumbnails"
         Height          =   255
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   2655
      End
      Begin VB.Label Label1 
         Caption         =   "Size"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.PictureBox Picture3 
      Height          =   255
      Left            =   0
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   173
      TabIndex        =   7
      Top             =   6120
      Width           =   2655
      Begin VB.PictureBox Picture4 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   -15
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   1
         TabIndex        =   8
         Top             =   0
         Width           =   15
      End
   End
   Begin VB.FileListBox File1 
      Height          =   1065
      Hidden          =   -1  'True
      Left            =   600
      Pattern         =   "*.jpg;*.jpeg;*.bmp;*.gif;*.emf;*.wmf;*.ico;*.cur"
      System          =   -1  'True
      TabIndex        =   6
      Top             =   3120
      Width           =   1455
      Visible         =   0   'False
   End
   Begin VB.DirListBox Dir1 
      Height          =   5040
      Left            =   0
      TabIndex        =   5
      Top             =   360
      Width           =   2655
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   2655
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   5655
      Left            =   11760
      TabIndex        =   3
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00404040&
      Height          =   6375
      Left            =   2760
      ScaleHeight     =   421
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   605
      TabIndex        =   1
      Top             =   0
      Width           =   9135
      Begin VB.PictureBox errorP 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   1455
         Left            =   6840
         ScaleHeight     =   93
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   101
         TabIndex        =   13
         Top             =   4200
         Width           =   1575
         Visible         =   0   'False
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   6255
         Left            =   0
         ScaleHeight     =   417
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   433
         TabIndex        =   2
         Top             =   0
         Width           =   6495
         Begin VB.Shape Shape1 
            BorderWidth     =   5
            DrawMode        =   6  'Mask Pen Not
            Height          =   1695
            Left            =   0
            Top             =   0
            Width           =   1815
            Visible         =   0   'False
         End
      End
   End
   Begin VB.PictureBox tmp 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   6360
      ScaleHeight     =   89
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   97
      TabIndex        =   0
      Top             =   4560
      Width           =   1455
      Visible         =   0   'False
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WW As Integer
Dim HH As Integer
Dim LL As Integer
Dim TT As Integer

Dim TW As Integer
Dim TH As Integer

Dim DW As Integer
Dim DH As Integer

Dim Ratio As Double

Dim FN As String
Dim Counter As Long
Dim Progress As Integer
Dim values As Long

Dim Size As Integer

Dim TempX As Integer
Dim TempY As Integer

Const STRETCHMODE = vbPaletteModeNone
Sub CalculateXcells()
'calculate maximum X cells
TW = Picture2.ScaleWidth - Size * 1.5
TW = TW / Val(Combo1.Text)

If TW < 0 Then TW = 0

Picture1.Width = (TW + 1) * Size
End Sub


Sub DrawTumbnail()
On Error GoTo kaka
If File1.ListCount = 0 Then
MsgBox "No images found in this directory!", vbExclamation
Exit Sub
End If
VScroll1.Value = 0

Picture1.Cls
Shape1.Visible = False
DW = 0
DH = 0
'set statusbar
values = Picture3.Width / File1.ListCount
'set height
TH = File1.ListCount / (TW + 1)
TH = (TH + 1) * Size
Picture1.Height = TH
'update scrollbars
SetScroll

For Counter = 0 To File1.ListCount - 1
File1.ListIndex = Counter

'See if we are in a sub dir
If Right(File1.Path, 1) <> "\" Then
        FN = File1.Path & "\" & File1.FileName
      Else
        FN = File1.Path & File1.FileName
End If
'Load the image so we can do a thumbnail of it
tmp.Picture = LoadPicture(FN)

'Fix aspect ratio
       If tmp.ScaleWidth > tmp.ScaleHeight Then
            Ratio = Abs(tmp.ScaleWidth / tmp.ScaleHeight)
            WW = Size
            HH = Size / Ratio
        Else
            Ratio = Abs(tmp.ScaleHeight / tmp.ScaleWidth)
            HH = Size
            WW = Size / Ratio
        End If
'Set position
LL = (Size - WW) / 2
TT = (Size - HH) / 2



'Draw it
If Check1.Value = 1 Then
SetStretchBltMode Picture1.hdc, STRETCHMODE
End If

StretchBlt Picture1.hdc, (DW * Size) + LL, (DH * Size) + TT, WW, HH, tmp.hdc, 0, 0, tmp.ScaleWidth, tmp.ScaleHeight, vbSrcCopy
Picture4.Width = Picture4.Width + values
Picture3.Refresh

DW = DW + 1

If DW > TW Then
DW = 0
DH = DH + 1
End If

Next
Picture1.Refresh
Picture4.Width = 1

Exit Sub
kaka:
Set tmp.Picture = errorP.Image
Resume Next
End Sub



Sub SetScroll()
If Picture1.Height <= Picture2.ScaleHeight Then
VScroll1.Enabled = False
Else
VScroll1.Enabled = True
VScroll1.Max = Picture1.ScaleHeight - Picture2.ScaleHeight
VScroll1.LargeChange = VScroll1.Max / 5
End If


End Sub

Private Sub Combo1_Click()
Size = Combo1.Text
CalculateXcells
Shape1.Width = Size
Shape1.Height = Size
Picture1.Cls
Shape1.Visible = False
End Sub


Private Sub Command1_Click()
DrawTumbnail
End Sub

Private Sub Dir1_Change()
 File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
On Error GoTo DriveHandler
Dir1.Path = Drive1.Drive
Exit Sub

DriveHandler:
    Drive1.Drive = Dir1.Path
    MsgBox "Drive not ready!", vbCritical
    Exit Sub
End Sub

Private Sub Form_Load()
Combo1.AddItem "32"
Combo1.AddItem "50"
Combo1.AddItem "75"
Combo1.AddItem "100"
Combo1.AddItem "150"
Combo1.AddItem "200"
Combo1.AddItem "250"
Combo1.AddItem "300"

Combo1 = "100"

Size = Combo1.Text
Shape1.Width = Size
Shape1.Height = Size


errorP.Print "ERROR!"
MsgBox "Thanks for downloading my code! Double-click on a thumbnail to view it in real size. When you are in the viewer: click with left button to close it, drag image with right button to move it (good if image is bigger then screen)", vbInformation
End Sub

Private Sub Form_Resize()
On Error Resume Next
If Form1.Width > 300 * 15 Then
Dir1.Height = Form1.ScaleHeight - Drive1.Height - Frame1.Height - Picture3.ScaleHeight - 5
Dir1.Top = Drive1.Height
Frame1.Top = Drive1.Height + Dir1.Height
Picture3.Top = Frame1.Height + Drive1.Height + Dir1.Height
Picture2.Height = Form1.ScaleHeight
VScroll1.Height = Form1.ScaleHeight
Picture2.Width = Form1.ScaleWidth - Dir1.Width - VScroll1.Width
VScroll1.Left = Form1.ScaleWidth - VScroll1.Width
CalculateXcells
SetScroll
Else
Form1.Width = 300 * 15
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
MsgBox "Please vote or leave feedback! =)", vbInformation
End
End Sub

Private Sub Picture1_DblClick()
If Not FN = "" Then
Form2.Show
Form2.OpenPicture FN
End If
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Calculate where to place marker
TempX = X
TempY = Y

TempX = TempX - (Size / 2)
TempY = TempY - (Size / 2)

TempX = TempX / Size
TempY = TempY / Size

Shape1.Left = TempX * Size + 1
Shape1.Top = TempY * Size + 1
Shape1.Visible = True

'Get the position of the selected thumbnail
TempX = (Shape1.Left / Size)
TempY = (Shape1.Top / Size) * (TW + 1)
If Val(TempX + TempY) <= File1.ListCount - 1 Then
File1.ListIndex = (TempX + TempY)
'See if we are in a sub dir
If Right(File1.Path, 1) <> "\" Then
    FN = File1.Path & "\" & File1.FileName
Else
    FN = File1.Path & File1.FileName
End If
Else
Shape1.Visible = False
FN = ""
End If
End Sub


Private Sub VScroll1_Change()
Picture1.Top = 0 - VScroll1.Value
End Sub


Private Sub VScroll1_Scroll()
VScroll1_Change
End Sub


