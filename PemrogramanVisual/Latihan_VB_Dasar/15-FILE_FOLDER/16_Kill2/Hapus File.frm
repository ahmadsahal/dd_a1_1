VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Menghapus file"
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8085
   LinkTopic       =   "Form1"
   ScaleHeight     =   3915
   ScaleWidth      =   8085
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdKeluar 
      Caption         =   "Keluar"
      Height          =   372
      Left            =   6720
      TabIndex        =   6
      Top             =   2520
      Width           =   1092
   End
   Begin VB.CommandButton CmdHapus 
      Caption         =   "Hapus"
      Height          =   372
      Left            =   6720
      TabIndex        =   5
      Top             =   1920
      Width           =   1092
   End
   Begin VB.Frame Frame1 
      Caption         =   "File akan dihapus"
      Height          =   1332
      Left            =   4680
      TabIndex        =   3
      Top             =   240
      Width           =   3132
      Begin VB.Label Label1 
         Height          =   972
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   2892
      End
   End
   Begin VB.FileListBox File1 
      Height          =   3210
      Left            =   2640
      TabIndex        =   2
      Top             =   240
      Width           =   1932
   End
   Begin VB.DirListBox Dir1 
      Height          =   2880
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   2052
   End
   Begin VB.DriveListBox Drive1 
      Height          =   288
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2052
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public NamaFile As String

Private Sub CmdKeluar_Click()
    End
End Sub

Private Sub Drive1_Change()
    Dir1.Path = UCase(Left$(Drive1.Drive, 1)) & ":\"
End Sub

Private Sub Dir1_Change()
    File1.Pattern = "*.*"
    File1.Path = Dir1.Path
End Sub

Private Sub File1_Click()
    NamaFile = File1.Path
    If Right$(NamaFile, 1) <> "\" Then NamaFile = NamaFile & "\"
    NamaFile = NamaFile & File1.FileName
    Label1.Caption = NamaFile
End Sub

Private Sub CmdHapus_Click()
    Dim Pesan%
    Pesan% = MsgBox("Apakah file ini: " & NamaFile & " akan dihapus?", vbExclamation + vbYesNo + vbDefaultButton2, "Menghapus File")
    If Pesan% = vbYes Then
        Kill NamaFile
        File1.Refresh
    End If
End Sub

