VERSION 5.00
Begin VB.Form NasabahFrm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tambah Nasabah Baru"
   ClientHeight    =   5460
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8970
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   8970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CancelBtn 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5018
      TabIndex        =   19
      Top             =   4800
      Width           =   2415
   End
   Begin VB.CommandButton TambahBtn 
      Caption         =   "Tambah"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   18
      Top             =   4800
      Width           =   2415
   End
   Begin VB.Frame Frame2 
      Caption         =   "Jenis Tabungan"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   6240
      TabIndex        =   13
      Top             =   720
      Width           =   2415
      Begin VB.OptionButton Option4 
         Caption         =   "Ria Exim"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   2040
         Width           =   1935
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Ria Dollar"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   1560
         Width           =   1935
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Super Saving"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   1080
         Width           =   1935
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Tahapan"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   600
         Value           =   -1  'True
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Data Nasabah"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   5775
      Begin VB.TextBox TxtSetor 
         Height          =   375
         Left            =   1320
         TabIndex        =   12
         Top             =   3000
         Width           =   4215
      End
      Begin VB.TextBox TxtPhone 
         Height          =   375
         Left            =   1320
         MaxLength       =   8
         TabIndex        =   10
         Top             =   2400
         Width           =   4215
      End
      Begin VB.TextBox TxtKota 
         Height          =   375
         Left            =   1320
         TabIndex        =   9
         Top             =   1800
         Width           =   4215
      End
      Begin VB.TextBox TxtAlamat 
         Height          =   375
         Left            =   1320
         TabIndex        =   6
         Top             =   1200
         Width           =   4215
      End
      Begin VB.TextBox TxtNama 
         Height          =   375
         Left            =   1320
         TabIndex        =   5
         Top             =   600
         Width           =   4215
      End
      Begin VB.Label Label6 
         Caption         =   "Setoran awal"
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
         TabIndex        =   11
         Top             =   3000
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Phone"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Kota"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Alamat"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Nama"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   975
      End
   End
   Begin VB.Label LblNoRek 
      Caption         =   "Not assigned!! (yet)"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label Label3 
      Caption         =   "Nomor Rekening"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "NasabahFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NoRekTemp As String
Dim Suffix As String
Dim Prefix As String

Private Sub CancelBtn_Click()
Unload Me
End Sub

Private Sub Form_Load()

'buka recordset tabel Profile, Saving, dan transaksi
'eksekusi perintah SQL
Cmd.CommandText = "SELECT * FROM profile"
Set Cmd.ActiveConnection = MyDB
'Set Rc = Cmd.Execute
RcProfile.Open Cmd, , adOpenDynamic, adLockOptimistic
    
Cmd.CommandText = "SELECT*FROM saving"
RcSaving.Open Cmd, , adOpenDynamic, adLockOptimistic
    
Cmd.CommandText = "SELECT*FROM transaksi"
RcTransaksi.Open Cmd, , adOpenDynamic, adLockOptimistic

Prefix = "740-10-"
Suffix = "-1"
Option1.Value = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
'tutup semua recordset
RcProfile.Close
RcSaving.Close
RcTransaksi.Close

End Sub

Private Sub Option1_Click()
Suffix = "-1"
End Sub

Private Sub Option2_Click()
Suffix = "-2"
End Sub

Private Sub Option3_Click()
Suffix = "-3"
End Sub

Private Sub Option4_Click()
Suffix = "-4"
End Sub

Private Sub TambahBtn_Click()
Dim x As Long
Dim Jawab As VbMsgBoxResult

'cek apakah ada textbox yang kosong
If TxtNama.Text = "" Then Exit Sub
If TxtAlamat.Text = "" Then Exit Sub
If TxtKota.Text = "" Then Exit Sub
If TxtPhone.Text = "" Then Exit Sub
If TxtSetor.Text = "" Then Exit Sub

'ambil tgl sekarang
Waktu = Date

'bawa pointer ke record terakhir
RcProfile.MoveLast
RcSaving.MoveLast
RcTransaksi.MoveLast

x = Val(Mid(RcProfile!NoRek, 8, 5))
NoRekTemp = Trim(Str(x + 1))
NoRekTemp = Trim(Prefix + NoRekTemp + Suffix)

'konfirmasi akhir
Jawab = MsgBox("Nomor rekening: " + NoRekTemp + vbCrLf + "Nama: " + TxtNama.Text + vbCrLf + "dibuat sekarang?", vbYesNo, "Konfirmasi akhir")

If Not Jawab = vbYes Then Exit Sub


'rutin tambah record ke tiga tabel
'assign nilai ke masing-masing record
RcProfile.AddNew

With RcProfile
    !NoRek = NoRekTemp
    !nama = TxtNama.Text
    !alamat = TxtAlamat.Text
    !kota = TxtKota.Text
    !phone = TxtPhone.Text
    .Update
End With

x = Val(RcSaving!no)

RcSaving.AddNew

With RcSaving
    !no = Str(x + 1)
    !NoRek = NoRekTemp
    !saldo = Val(TxtSetor.Text)
    .Update
End With

x = Val(RcTransaksi!tid)

RcTransaksi.AddNew

With RcTransaksi
    !tid = Str(x + 1)
    !tgl = Waktu
    !NoRek = NoRekTemp
    !jenis = "kredit"
    !Jumlah = Val(TxtSetor.Text)
    .Update
End With



Unload Me
End Sub
