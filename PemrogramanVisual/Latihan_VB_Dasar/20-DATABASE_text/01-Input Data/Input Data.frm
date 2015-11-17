VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "INPUT DATA"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4425
   LinkTopic       =   "Form1"
   ScaleHeight     =   2400
   ScaleWidth      =   4425
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   288
      Left            =   1560
      TabIndex        =   9
      Text            =   "1"
      Top             =   720
      Width           =   1572
   End
   Begin VB.CommandButton CmdSelesai 
      Caption         =   "Selesai"
      Height          =   372
      Left            =   360
      TabIndex        =   8
      Top             =   1680
      Width           =   1212
   End
   Begin VB.CommandButton CmdBatal 
      Caption         =   "Batal"
      Height          =   372
      Left            =   1680
      TabIndex        =   7
      Top             =   1680
      Width           =   1212
   End
   Begin VB.CommandButton CmdSimpan 
      Caption         =   "Simpan"
      Height          =   372
      Left            =   3000
      TabIndex        =   5
      Top             =   1680
      Width           =   1212
   End
   Begin VB.TextBox TxtGaji 
      Height          =   288
      Left            =   1560
      TabIndex        =   3
      Top             =   1200
      Width           =   1572
   End
   Begin VB.TextBox TxtNama 
      Height          =   288
      Left            =   1560
      TabIndex        =   0
      ToolTipText     =   "Ketik Nama Karyawan kemudian tekan TAB atau Enter."
      Top             =   240
      Width           =   2412
   End
   Begin VB.Label Label4 
      Caption         =   "#.###.###"
      Height          =   252
      Left            =   3360
      TabIndex        =   6
      Top             =   1200
      Width           =   972
   End
   Begin VB.Label Label3 
      Caption         =   "Gaji Pokok"
      Height          =   252
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   972
   End
   Begin VB.Label Label2 
      Caption         =   "Kode Golongan"
      Height          =   252
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   1212
   End
   Begin VB.Label Label1 
      Caption         =   "Nama Karyawan"
      Height          =   252
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1452
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Combo1.AddItem "1", 0
    Combo1.AddItem "2", 1
    Combo1.AddItem "3", 2
End Sub

Private Sub CmdBatal_Click()
    TxtNama = ""
    TxtGaji = ""
    TxtNama.SetFocus
End Sub

Private Sub CmdSelesai_Click()
    End
End Sub

Private Sub CmdSimpan_Click()
    Dim DirAktif As String
    If TxtNama = "" Or TxtGaji = "" Then GoTo AdaKosong
    DirAktif = Left(CurDir, 3)
    ChDir DirAktif
    Open App.Path & "\GAJI.DAT" For Append As #1
    Write #1, TxtNama, Combo1.Text, TxtGaji
    Close #1
    TxtNama = ""
    TxtGaji = ""
    TxtNama.SetFocus
    Exit Sub
AdaKosong:
    Beep
    If TxtNama = "" Then TxtNama.SetFocus
    If TxtGaji = "" Then TxtGaji.SetFocus
End Sub

