VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "HAPUS DATA"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5190
   LinkTopic       =   "Form1"
   ScaleHeight     =   3390
   ScaleWidth      =   5190
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdKeluar 
      Caption         =   "Keluar"
      Height          =   372
      Left            =   1440
      TabIndex        =   11
      Top             =   2640
      Width           =   972
   End
   Begin VB.CommandButton CmdBatal 
      Caption         =   "Batal"
      Height          =   372
      Left            =   2520
      TabIndex        =   10
      Top             =   2640
      Width           =   1092
   End
   Begin VB.CommandButton CmdHapus 
      Caption         =   "Hapus"
      Height          =   372
      Left            =   3720
      TabIndex        =   9
      Top             =   2640
      Width           =   1092
   End
   Begin VB.Frame Frame2 
      Caption         =   "DITEMUKAN"
      Height          =   1212
      Left            =   360
      TabIndex        =   2
      Top             =   1200
      Width           =   4452
      Begin VB.TextBox TxtGaji 
         Height          =   288
         Left            =   2880
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   480
         Width           =   1452
      End
      Begin VB.TextBox TxtKode 
         Height          =   288
         Left            =   1680
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   480
         Width           =   1092
      End
      Begin VB.TextBox TxtNama 
         Height          =   288
         Left            =   120
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   480
         Width           =   1452
      End
      Begin VB.Label Label3 
         Caption         =   "Gaji"
         Height          =   252
         Left            =   2880
         TabIndex        =   5
         Top             =   240
         Width           =   492
      End
      Begin VB.Label Label2 
         Caption         =   "Golongan"
         Height          =   252
         Left            =   1680
         TabIndex        =   4
         Top             =   240
         Width           =   852
      End
      Begin VB.Label Label1 
         Caption         =   "Nama"
         Height          =   252
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   732
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Nama akan dihapus"
      Height          =   732
      Left            =   2640
      TabIndex        =   0
      Top             =   240
      Width           =   2172
      Begin VB.TextBox TxtNamaCari 
         Height          =   288
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "Ketik nama yang akan dihapus kemudian Enter."
         Top             =   240
         Width           =   1932
      End
   End
   Begin VB.Label Label4 
      Caption         =   "HAPUS DATA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   360
      TabIndex        =   12
      Top             =   360
      Width           =   1812
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    CmdHapus.Enabled = False
End Sub

Private Sub CariData()
    Dim Nama, Kode, Gaji As String
    Dim Ada As Integer
    Open App.Path & "\GAJI.DAT" For Input As #1
    Do Until EOF(1)
        Input #1, Nama, Kode, Gaji
        If UCase(TxtNamaCari) = UCase(Nama) Then
            TxtNama = Nama
            TxtKode = Kode
            TxtGaji = Format(Gaji, "Currency")
            Ada = Ada + 1
            CmdHapus.Enabled = True 'Hidupkan tombol Hapus
        End If
    Loop
    Close
    If Ada <> 0 Then TxtNamaCari.Enabled = False
    If Ada = 0 Then
        MsgBox "Nama '" & TxtNamaCari & "' tidak ada dalam file!"
        Kosongkan
        TxtNamaCari.SetFocus
    End If
End Sub
  
Private Sub Kosongkan()
    TxtNamaCari = ""
    TxtNama = ""
    TxtKode = ""
    TxtGaji = ""
End Sub

Private Sub CmdHapus_Click()
    Dim Nama, Kode, Gaji As String
    Open "C:\VB6\GAJI.DAT" For Input As #1
    Open "C:\VB6\TEMPORER.DAT" For Append As #2

LEWATKAN:
    Do Until EOF(1)
        Input #1, Nama, Kode, Gaji
        If UCase(Nama) = UCase(TxtNamaCari) Then
            GoTo LEWATKAN 'Jangan ditulis ke dalam file
        End If
        Write #2, Nama, Kode, Gaji
    Loop
    Close
    Kill "C:\VB6\GAJI.DAT"
    Name "C:\VB6\TEMPORER.DAT" As "C:\VB6\GAJI.DAT"
    MsgBox "Data sudah dihapus, klik OK!"
    Kosongkan
    TxtNamaCari.Enabled = True
    TxtNamaCari.SetFocus
    CmdHapus.Enabled = False
End Sub

Private Sub CmdBatal_Click()
    Kosongkan
    TxtNamaCari.Enabled = True
    TxtNamaCari.SetFocus
    CmdHapus.Enabled = False
End Sub

Private Sub CmdKeluar_Click()
    End
End Sub

Private Sub TxtNamaCari_KeyDown(KeyCode As Integer, Shift As Integer)
    If TxtNamaCari <> "" And KeyCode = 13 Then CariData
    If KeyCode = 27 Then End
End Sub

