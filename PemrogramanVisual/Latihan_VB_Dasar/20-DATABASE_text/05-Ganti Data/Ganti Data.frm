VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "GANTI DATA"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5610
   LinkTopic       =   "Form1"
   ScaleHeight     =   3960
   ScaleWidth      =   5610
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdSelesai 
      Caption         =   "Selesai"
      Height          =   372
      Left            =   240
      TabIndex        =   12
      Top             =   2760
      Width           =   972
   End
   Begin VB.CommandButton CmdBatal 
      Caption         =   "Batal"
      Height          =   372
      Left            =   3240
      TabIndex        =   11
      Top             =   2760
      Width           =   972
   End
   Begin VB.CommandButton CmdGanti 
      Caption         =   "Ganti"
      Height          =   372
      Left            =   4320
      TabIndex        =   10
      Top             =   2760
      Width           =   972
   End
   Begin VB.Frame Frame2 
      Caption         =   "DITEMUKAN"
      Height          =   1092
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   5052
      Begin VB.TextBox TxtGaji 
         Height          =   288
         Left            =   3240
         TabIndex        =   9
         Top             =   480
         Width           =   1692
      End
      Begin VB.TextBox TxtKode 
         Height          =   288
         Left            =   2160
         TabIndex        =   8
         Top             =   480
         Width           =   972
      End
      Begin VB.TextBox TxtNama 
         Height          =   288
         Left            =   360
         TabIndex        =   7
         Top             =   480
         Width           =   1692
      End
      Begin VB.Label Label3 
         Caption         =   "Gaji"
         Height          =   252
         Left            =   3240
         TabIndex        =   4
         Top             =   240
         Width           =   492
      End
      Begin VB.Label Label2 
         Caption         =   "Golongan"
         Height          =   252
         Left            =   2160
         TabIndex        =   3
         Top             =   240
         Width           =   852
      End
      Begin VB.Label Label1 
         Caption         =   "Nama"
         Height          =   252
         Left            =   360
         TabIndex        =   2
         Top             =   240
         Width           =   732
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Nama akan diganti"
      Height          =   732
      Left            =   3120
      TabIndex        =   0
      Top             =   240
      Width           =   2172
      Begin VB.TextBox TxtNamaCari 
         Height          =   288
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1932
      End
   End
   Begin VB.Label Label7 
      Caption         =   "MENGGANTI DATA"
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
      Left            =   240
      TabIndex        =   5
      Top             =   360
      Width           =   2532
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CariData()
    Dim Nama, Kode, Gaji As String
    Dim Ada As Integer
     
    Open App.Path & "\GAJI.DAT" For Input As #1
    
    Do Until EOF(1)
        Input #1, Nama, Kode, Gaji
        If UCase(TxtNamaCari) = UCase(Nama) Then
            TxtNama = Nama
            TxtKode = Kode
            TxtGaji = Gaji
            Ada = Ada + 1
            CmdGanti.Enabled = True
        End If
    Loop
    Close
    TxtNama.SetFocus

    If Ada = 0 Then
        MsgBox "Nama '" & TxtNamaCari & "' tidak ada dalam file!"
        Kosongkan
    End If
End Sub
  
Private Sub Kosongkan()
    TxtNamaCari = ""
    TxtNama = ""
    TxtKode = ""
    TxtGaji = ""
    TxtNamaCari.SetFocus
    CmdGanti.Enabled = False
End Sub

Private Sub CmdBatal_Click()
    Kosongkan
End Sub

Private Sub CmdGanti_Click()
    Dim Nama, Kode, Gaji As String
    Open App.Path & "\GAJI.DAT" For Input As #1
    Open App.Path & "\TEMPORER.DAT" For Append As #2
    Do Until EOF(1)
        Input #1, Nama, Kode, Gaji
        If UCase(Nama) = UCase(TxtNamaCari) Then
            Nama = TxtNama
            Kode = TxtKode
            Gaji = TxtGaji
        End If
        Write #2, Nama, Kode, Gaji
    Loop
    Close
    Kill App.Path & "\GAJI.DAT"
    Name App.Path & "\TEMPORER.DAT" As App.Path & "\GAJI.DAT"
    MsgBox "Data sudah diganti, klik OK!"
    Kosongkan
End Sub

Private Sub CmdSelesai_Click()
    End
End Sub

Private Sub Form_Load()
    'Matikan tombol GANTI
    CmdGanti.Enabled = False
End Sub

Private Sub TxtNamaCari_KeyDown(KeyCode As Integer, Shift As Integer)
    If TxtNamaCari <> "" And KeyCode = 13 Then CariData
    If KeyCode = 27 Then End
End Sub

