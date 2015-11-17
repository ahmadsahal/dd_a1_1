VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7200
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7500
   LinkTopic       =   "Form1"
   ScaleHeight     =   7200
   ScaleWidth      =   7500
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   2595
      Left            =   960
      TabIndex        =   6
      Top             =   3600
      Width           =   4935
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2280
      TabIndex        =   5
      Top             =   1320
      Width           =   4215
   End
   Begin VB.TextBox txtTerbilang 
      Height          =   495
      Left            =   2280
      TabIndex        =   4
      Top             =   840
      Width           =   4215
   End
   Begin VB.CommandButton cmdProses 
      Caption         =   "Proses"
      Height          =   615
      Left            =   1440
      TabIndex        =   2
      Top             =   2280
      Width           =   2895
   End
   Begin VB.TextBox txtAngka 
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Top             =   240
      Width           =   4215
   End
   Begin VB.Label Label2 
      Caption         =   "Terbilang"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Angka"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim no_antri, no_antri_panggil As Integer
Dim Sounds(16) As String
Sub panggil()
Dim arrJumlahKarakterSpasi() As String
arrJumlahKarakterSpasi = Split(txtTerbilang.Text, " ")
    List1.Clear
    For i = LBound(arrJumlahKarakterSpasi) To UBound(arrJumlahKarakterSpasi)
        Call sndPlaySound(App.Path & "\Sounds\" & _
        arrJumlahKarakterSpasi(i) & ".wav", SND_NOSTOP)
        List1.AddItem arrJumlahKarakterSpasi(i)
    Next
    Call sndPlaySound(App.Path & "\Sounds\rupiah.wav", SND_NOSTOP)
End Sub
Private Sub cmdProses_Click()
txtTerbilang.Text = Trim(Bilang(txtAngka.Text))
    Text1.Text = Trim(Bilang(txtAngka.Text))
    Call panggil
End Sub

Private Sub Form_Load()
Me.KeyPreview = True
no_antri = 0
no_antri_panggil = 0
Sounds(1) = App.Path & "\Sounds\satu.wav"
   Sounds(2) = App.Path & "\Sounds\dua.wav"
   Sounds(3) = App.Path & "\Sounds\tiga.wav"
   Sounds(4) = App.Path & "\Sounds\empat.wav"
   Sounds(5) = App.Path & "\Sounds\lima.wav"
   Sounds(6) = App.Path & "\Sounds\enam.wav"
   Sounds(7) = App.Path & "\Sounds\tujuh.wav"
   Sounds(8) = App.Path & "\Sounds\delapan.wav"
   Sounds(9) = App.Path & "\Sounds\sembilan.wav"
   Sounds(10) = App.Path & "\Sounds\sepuluh.wav"
   Sounds(11) = App.Path & "\Sounds\sebelas.wav"
   Sounds(12) = App.Path & "\Sounds\puluh.wav"
   Sounds(13) = App.Path & "\Sounds\ratus.wav"
   Sounds(14) = App.Path & "\Sounds\belas.wav"
   Sounds(15) = App.Path & "\Sounds\nomor-urut.wav"
   Sounds(16) = App.Path & "\Sounds\loket.wav"
End Sub

