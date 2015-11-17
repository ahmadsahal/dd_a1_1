VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Menampilkan isi disk"
   ClientHeight    =   2748
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   4512
   LinkTopic       =   "Form1"
   ScaleHeight     =   2748
   ScaleWidth      =   4512
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox File1 
      Height          =   1992
      Left            =   2280
      TabIndex        =   2
      Top             =   240
      Width           =   1812
   End
   Begin VB.DirListBox Dir1 
      Height          =   1584
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   1692
   End
   Begin VB.DriveListBox Drive1 
      Height          =   288
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1692
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Drive1_Change()
    Dir1.Path = UCase(Drive1.Drive)
End Sub

Private Sub Dir1_Change()
    File1.Pattern = "*.*"
    File1.Path = Dir1.Path
End Sub

Private Sub File1_Click()
    Namafile = Dir1.Path + "\" + File1.FileName
    If Mid$(Namafile, 4, 1) = "\" Then
        Namafile = Dir1.Path + File1.FileName
    End If
    Form1.Caption = Namafile
End Sub


