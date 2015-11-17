VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "MENGCOPY FILE"
   ClientHeight    =   4956
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   8124
   LinkTopic       =   "Form1"
   ScaleHeight     =   6720
   ScaleWidth      =   9504
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Caption         =   "TUJUAN"
      Height          =   1092
      Left            =   4440
      TabIndex        =   11
      Top             =   3480
      Width           =   3252
      Begin VB.Label Label2 
         Caption         =   "Belum diipilih"
         Height          =   612
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   2892
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "ASAL"
      Height          =   1092
      Left            =   240
      TabIndex        =   6
      Top             =   3480
      Width           =   3732
      Begin VB.Label Label1 
         Caption         =   "Belum dipilih"
         Height          =   612
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   3372
      End
   End
   Begin VB.CommandButton CmdKeluar 
      Caption         =   "Keluar"
      Height          =   372
      Left            =   6600
      TabIndex        =   5
      Top             =   1080
      Width           =   1092
   End
   Begin VB.CommandButton CmdCopy 
      Caption         =   "Copy"
      Height          =   372
      Left            =   6600
      TabIndex        =   4
      Top             =   480
      Width           =   1092
   End
   Begin VB.Frame Frame1 
      Caption         =   "TUJUAN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4572
      Left            =   4320
      TabIndex        =   3
      Top             =   120
      Width           =   3612
      Begin VB.DriveListBox Drive2 
         Height          =   288
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   1812
      End
      Begin VB.DirListBox Dir2 
         Height          =   2448
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   1812
      End
   End
   Begin VB.FileListBox File1 
      Height          =   2952
      Left            =   2040
      TabIndex        =   2
      Top             =   480
      Width           =   1932
   End
   Begin VB.DirListBox Dir1 
      Height          =   2448
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   1692
   End
   Begin VB.DriveListBox Drive1 
      Height          =   288
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   1812
   End
   Begin VB.Frame Frame3 
      Caption         =   "ASAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4572
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   4092
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public NamaFile, Tujuan As String

Private Sub CmdCopy_Click()
    If NamaFile = "" Or Tujuan = "" Then
        Beep
        MsgBox "Tidak dapat dicopy, file belum dipilih!", vbQuestion, "Mengcopy File"
        Exit Sub
    End If
    FileCopy NamaFile, Tujuan
    MsgBox "File sudah dicopy dengan sukses!", vbInformation, "Mengcopy File"
End Sub

Private Sub CmdKeluar_Click()
    End
End Sub

Private Sub Drive2_Change()
    Dir2.Path = UCase(Drive2.Drive)
End Sub

Private Sub Dir2_Change()
    Tujuan = Dir2.Path + "\" + File1.FileName
    If Mid$(Tujuan, 4, 1) = "\" Then
        Tujuan = Dir2.Path + File1.FileName
    End If
    Label2.Caption = Tujuan
End Sub

Private Sub Drive1_Change()
    Dir1.Path = UCase(Drive1.Drive)
End Sub

Private Sub Dir1_Change()
    File1.Pattern = "*.*"
    File1.Path = Dir1.Path
End Sub

Private Sub File1_Click()
    NamaFile = Dir1.Path + "\" + File1.FileName
    If Mid$(NamaFile, 4, 1) = "\" Then
        NamaFile = Dir1.Path + File1.FileName
    End If
    Label1.Caption = NamaFile
    Dir2_Change
End Sub

