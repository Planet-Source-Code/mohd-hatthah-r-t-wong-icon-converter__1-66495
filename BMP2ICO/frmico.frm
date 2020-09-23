VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmico 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Icon Converter"
   ClientHeight    =   4125
   ClientLeft      =   3060
   ClientTop       =   2355
   ClientWidth     =   4740
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   4740
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   960
      ScaleHeight     =   161
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   177
      TabIndex        =   0
      Top             =   600
      Width           =   2655
   End
   Begin MSComctlLib.ImageList img 
      Left            =   1920
      Top             =   5400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   840
      Top             =   5520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu fail 
      Caption         =   "&Fail"
      Begin VB.Menu buka 
         Caption         =   "Buka/ Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu simpan 
         Caption         =   "&Simpan-Sebagai/ Save-as"
         Shortcut        =   ^S
      End
      Begin VB.Menu ap1 
         Caption         =   "-"
      End
      Begin VB.Menu keluar 
         Caption         =   "&Keluar/ Quit"
         Shortcut        =   ^Q
      End
   End
End
Attribute VB_Name = "frmico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lPic As Picture
Private Sub buka_Click()
On Error Resume Next
With cd
.FileName = ""
.Filter = "(Gambar) | *.bmp;*.jpg;*.gif;"
.ShowOpen
If .FileName <> "" Then
Set lPic = LoadPicture(.FileName)
Draw
End If
End With

End Sub
Sub Draw()
If Not lPic Is Nothing Then
Pic.Cls
Pic.PaintPicture lPic, 0, 0, Pic.ScaleWidth, Pic.ScaleHeight
Pic = Pic.Image
End If

End Sub
Sub save()
Dim p As ListImage
Dim j As Picture
On Error GoTo er
If lPic Is Nothing Then Exit Sub
img.MaskColor = Pic.Point(0, 0)
img.ImageHeight = 32
img.ImageWidth = 32
Set p = img.ListImages.Add(, "m", Pic.Image)
Set j = p.ExtractIcon
With cd
.FileName = ""
.Filter = "(Gambar) | *.ico;"
.ShowSave
If .FileName <> "" Then
SavePicture j, .FileName
End If
img.ListImages.Clear

Exit Sub
er:
MsgBox "tidak dapat di sempurnakan"

End With


End Sub

Private Sub keluar_Click()
End
End Sub

Private Sub simpan_Click()
save
End Sub
