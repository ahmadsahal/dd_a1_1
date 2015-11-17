VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "SISTEM ANTRIAN"
   ClientHeight    =   9195
   ClientLeft      =   375
   ClientTop       =   -315
   ClientWidth     =   11085
   Icon            =   "Antri.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9195
   ScaleWidth      =   11085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   240
      TabIndex        =   16
      Top             =   4800
      Width           =   5295
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ANTRIAN NASABAH"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   22
         Top             =   240
         Width           =   4335
      End
      Begin VB.Line Line1 
         BorderWidth     =   5
         X1              =   360
         X2              =   4680
         Y1              =   2640
         Y2              =   2640
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "000"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   60
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1215
         Left            =   360
         TabIndex        =   19
         Top             =   2520
         Width           =   4335
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "START ANTRIAN"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   735
         Left            =   360
         TabIndex        =   18
         Top             =   1080
         Visible         =   0   'False
         Width           =   4215
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "000"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   80.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1695
         Left            =   360
         TabIndex        =   17
         Top             =   840
         Width           =   4335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Control Admin"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   5640
      TabIndex        =   6
      Top             =   4800
      Width           =   5295
      Begin VB.CommandButton Command1 
         Caption         =   "Panggil"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   13
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Ulang"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3480
         TabIndex        =   12
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Panggil"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   11
         Top             =   2520
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Ulang"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3480
         TabIndex        =   10
         Top             =   2520
         Width           =   1095
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Reset"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   9
         Top             =   3240
         Width           =   4095
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Start Antrian"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2520
         TabIndex        =   8
         Top             =   480
         Width           =   2415
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Tambah Antrian"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   7
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Teller 1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   480
         TabIndex        =   15
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Teller 2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   480
         TabIndex        =   14
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00FF80FF&
         FillStyle       =   0  'Solid
         Height          =   2775
         Left            =   240
         Top             =   1080
         Width           =   4695
      End
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   17040
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   9600
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   17040
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   9000
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Textantri 
      Height          =   735
      Left            =   17040
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   7920
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox txtTerbilang 
      Height          =   1455
      Left            =   17040
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   6240
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Image Image4 
      Height          =   450
      Left            =   7320
      Picture         =   "Antri.frx":08CA
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   660
   End
   Begin VB.Image Image3 
      Height          =   495
      Left            =   3480
      Picture         =   "Antri.frx":103E
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label8 
      Caption         =   "PROGRAMMER : IMAM SYAFI'I - FAKULTAS ILMU KOMPUTER UNSIKA"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   24
      Top             =   8880
      Width           =   6855
   End
   Begin VB.Image Image2 
      Height          =   3615
      Left            =   240
      Picture         =   "Antri.frx":2709
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   3135
   End
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   240
      Picture         =   "Antri.frx":3F29
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1440
   End
   Begin VB.Label Label7 
      BackColor       =   &H0080FFFF&
      Caption         =   "BANK FASILKOM UNSIKA"
      BeginProperty Font 
         Name            =   "Bernard MT Condensed"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1035
      Left            =   1680
      TabIndex        =   23
      Top             =   0
      Width           =   9270
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TELLER 2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   750
      Left            =   7080
      TabIndex        =   21
      Top             =   1080
      Width           =   3855
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TELLER 1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   750
      Left            =   3360
      TabIndex        =   20
      Top             =   1080
      Width           =   3615
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   99.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   7080
      TabIndex        =   1
      Top             =   1920
      Width           =   3855
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   99.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   3360
      TabIndex        =   0
      Top             =   1920
      Width           =   3615
   End
   Begin VB.Menu mnFile 
      Caption         =   "File"
      Begin VB.Menu mnExit 
         Caption         =   "Keluar"
      End
   End
   Begin VB.Menu mnHelp 
      Caption         =   "Bantuan"
      Begin VB.Menu mnAbout 
         Caption         =   "Tentang"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim no_antri, no_antri_panggil As Integer
Dim Sounds(16) As String
Sub panggil_L1()
Dim arrJumlahKarakterSpasi() As String
arrJumlahKarakterSpasi = Split(txtTerbilang.Text, " ")
    Call sndPlaySound(App.Path & "\Sounds\nomor-urut.wav", SND_NOSTOP)
    For i = LBound(arrJumlahKarakterSpasi) To UBound(arrJumlahKarakterSpasi)
        Call sndPlaySound(App.Path & "\Sounds\" & arrJumlahKarakterSpasi(i) & ".wav", SND_NOSTOP)
    Next
    Call sndPlaySound(App.Path & "\Sounds\loket.wav", SND_NOSTOP)
    Call sndPlaySound(App.Path & "\Sounds\satu.wav", SND_NOSTOP)
End Sub

Private Sub Command1_Click()
If no_antri_panggil < no_antri Then
    no_antri_panggil = no_antri_panggil + 1
    Label4.Caption = no_antri_panggil
    txtTerbilang.Text = Trim(Bilang(Label4.Caption))
    Text1.Text = Trim(Bilang(Label4.Caption))
    Call panggil_L1
    no_antri = no_antri - 1
    Label9.Caption = no_antri
End If
End Sub

Sub panggil_L2()
Dim arrJumlahKarakterSpasi() As String
arrJumlahKarakterSpasi = Split(txtTerbilang.Text, " ")
    Call sndPlaySound(App.Path & "\Sounds\nomor-urut.wav", SND_NOSTOP)
    For i = LBound(arrJumlahKarakterSpasi) To UBound(arrJumlahKarakterSpasi)
        Call sndPlaySound(App.Path & "\Sounds\" & arrJumlahKarakterSpasi(i) & ".wav", SND_NOSTOP)
    Next
    Call sndPlaySound(App.Path & "\Sounds\loket.wav", SND_NOSTOP)
    Call sndPlaySound(App.Path & "\Sounds\dua.wav", SND_NOSTOP)
End Sub

Private Sub Command2_Click()
Dim arrJumlahKarakterSpasi() As String
arrJumlahKarakterSpasi = Split(Text1.Text, " ")
    Call sndPlaySound(App.Path & "\Sounds\nomor-urut.wav", SND_NOSTOP)
    For i = 0 To UBound(arrJumlahKarakterSpasi)
        Call sndPlaySound(App.Path & "\Sounds\" & arrJumlahKarakterSpasi(i) & ".wav", SND_NOSTOP)
    Next
    Call sndPlaySound(App.Path & "\Sounds\loket.wav", SND_NOSTOP)
    Call sndPlaySound(App.Path & "\Sounds\satu.wav", SND_NOSTOP)
End Sub

Private Sub Command3_Click()
If no_antri_panggil < no_antri Then
    no_antri_panggil = no_antri_panggil + 1
    Label5.Caption = no_antri_panggil
    txtTerbilang.Text = Trim(Bilang(Label5.Caption))
    Text2.Text = Trim(Bilang(Label5.Caption))
    Call panggil_L2
    no_antri = no_antri - 1
    Label9.Caption = no_antri
End If
End Sub

Private Sub Command4_Click()
Dim arrJumlahKarakterSpasi() As String
arrJumlahKarakterSpasi = Split(Text2.Text, " ")
    Call sndPlaySound(App.Path & "\Sounds\nomor-urut.wav", SND_NOSTOP)
    For i = 0 To UBound(arrJumlahKarakterSpasi)
        Call sndPlaySound(App.Path & "\Sounds\" & arrJumlahKarakterSpasi(i) & ".wav", SND_NOSTOP)
    Next
    Call sndPlaySound(App.Path & "\Sounds\loket.wav", SND_NOSTOP)
    Call sndPlaySound(App.Path & "\Sounds\dua.wav", SND_NOSTOP)
End Sub

Private Sub Command5_Click()
no_antri = 0
no_antri_panggil = 0
Label4.Caption = no_antri_panggil
Label5.Caption = no_antri_panggil
Label3.Caption = no_antri
Label9.Caption = no_antri
Text1.Text = ""
Text2.Text = ""
Command6.Enabled = False
End Sub

Private Sub Command6_Click()
Label3.Visible = True
Label6.Visible = False
no_antri = no_antri + Val(Textantri.Text)
Label3.Caption = no_antri
Label9.Caption = no_antri
End Sub

Private Sub Command7_Click()
Form2.Show
Label3.Visible = False
Label6.Visible = True
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

Private Sub mnAbout_Click()
aboutcompany.Show
End Sub

Private Sub mnExit_Click()
End
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then Call Command6_Click
   If KeyAscii = 43 Then Call Command7_Click
   If KeyAscii = 49 Then Call Command1_Click
   If KeyAscii = 50 Then Call Command2_Click
   If KeyAscii = 51 Then Call Command3_Click
   If KeyAscii = 52 Then Call Command4_Click
   
End Sub

