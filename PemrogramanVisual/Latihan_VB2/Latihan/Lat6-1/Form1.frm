VERSION 5.00
Begin VB.Form ReportBtn 
   Caption         =   "Bank Ria Teller"
   ClientHeight    =   4875
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8160
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4875
   ScaleWidth      =   8160
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton ReportBtn 
      Caption         =   "Laporan Nasabah"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      MouseIcon       =   "Form1.frx":0442
      MousePointer    =   99  'Custom
      TabIndex        =   15
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton LogOutbtn 
      Caption         =   "LogOut"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      MouseIcon       =   "Form1.frx":074C
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton LoginBtn 
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      MouseIcon       =   "Form1.frx":0A56
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton NewBtn 
      Caption         =   "Nasabah Baru"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      MouseIcon       =   "Form1.frx":0D60
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   "Jasa Pembayaran"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   4320
      TabIndex        =   6
      Top             =   2520
      Width           =   3495
      Begin VB.CommandButton Voucherbtn 
         Caption         =   "Voucher"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1800
         TabIndex        =   10
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CommandButton HPPascaBtn 
         Caption         =   "HP Pasca Bayar"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CommandButton ListrikBtn 
         Caption         =   "Bayar Listrik"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1800
         TabIndex        =   8
         Top             =   480
         Width           =   1455
      End
      Begin VB.CommandButton AirBtn 
         Caption         =   "Bayar Air"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.CommandButton LogBtn 
      Caption         =   "Transaksi Log"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      MouseIcon       =   "Form1.frx":106A
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Transaksi Tunai"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   240
      TabIndex        =   2
      Top             =   2520
      Width           =   3855
      Begin VB.CommandButton TransferBtn 
         Caption         =   "Transfer"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1920
         TabIndex        =   14
         Top             =   480
         Width           =   1455
      End
      Begin VB.CommandButton TarikBtn 
         Caption         =   "Penarikan"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CommandButton SetorBtn 
         Caption         =   "Penyetoran"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.CommandButton TotalBtn 
      Caption         =   "Total Saldo"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      MouseIcon       =   "Form1.frx":1374
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   6720
      Top             =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Console Ria Teller"
      BeginProperty Font 
         Name            =   "Jokerman"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   975
      TabIndex        =   1
      Top             =   240
      Width           =   5400
   End
End
Attribute VB_Name = "ReportBtn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const Judul As String = "Bank Ria Teller - "



Private Sub AirBtn_Click()
JudulTarik = "Bayar Air"
TarikFrm.Show 1
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
Me.Show

Me.Caption = Judul + CStr(Now) + " - Locked! "

Stat = False

Kunci




End Sub

Private Sub HPPascaBtn_Click()
JudulTarik = "Bayar Rekening HP"
TarikFrm.Show 1
End Sub

Private Sub ListrikBtn_Click()
JudulTarik = "Bayar Rekening Listrik"
TarikFrm.Show 1
End Sub

Private Sub LogBtn_Click()
LogFrm.Show 1
End Sub

Private Sub LoginBtn_Click()

User = InputBox("Masukkan user anda", "Input User")
Pass = InputBox("Masukkan password " + User, "Input Password " + User)

If User = "" Then Exit Sub
If Pass = "" Then Exit Sub

Stat = BukaDB(User, Pass)

If Stat = False Then
    Kunci
Else
    Lepas
End If

End Sub

Private Sub Kunci()
TotalBtn.Enabled = False
LogBtn.Enabled = False
NewBtn.Enabled = False
LogOutbtn.Enabled = False
LoginBtn.Enabled = True

SetorBtn.Enabled = False
TarikBtn.Enabled = False
AirBtn.Enabled = False
ListrikBtn.Enabled = False
HPPascaBtn.Enabled = False
Voucherbtn.Enabled = False

End Sub

Private Sub Lepas()
TotalBtn.Enabled = True
LogBtn.Enabled = True
NewBtn.Enabled = True
LogOutbtn.Enabled = True
LoginBtn.Enabled = False

SetorBtn.Enabled = True
TarikBtn.Enabled = True
AirBtn.Enabled = True
ListrikBtn.Enabled = True
HPPascaBtn.Enabled = True
Voucherbtn.Enabled = True

End Sub
Private Sub LogOutbtn_Click()
Stat = False

'tutup database
MyDB.Close

Kunci

End Sub

Private Sub NewBtn_Click()
NasabahFrm.Show 1
End Sub

Private Sub ReportBtn_Click()
Report1.Show 1
End Sub

Private Sub SetorBtn_Click()
SetorFrm.Show 1
End Sub

Private Sub TarikBtn_Click()
JudulTarik = "Penarikan"
TarikFrm.Show 1
End Sub

Private Sub Timer1_Timer()

If Stat = False Then
    Me.Caption = Judul + CStr(Now) + " - Locked!"
Else
    Me.Caption = Judul + CStr(Now) + " - " + User
End If

End Sub

Private Sub TotalBtn_Click()
TotalFrm.Show 1
End Sub

Private Sub TransferBtn_Click()
TransferFrm.Show 1
End Sub

Private Sub Voucherbtn_Click()
JudulTarik = "Voucher Pulsa Ria GSM"
TarikFrm.Show 1
End Sub
