VERSION 5.00
Begin VB.Form TarikFrm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Penarikan"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5595
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   5595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CancelBtn 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3570
      TabIndex        =   5
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton OKBtn 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   930
      TabIndex        =   4
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox TxtSetor 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   1200
      Width           =   3015
   End
   Begin VB.TextBox TxtNoRek 
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      ToolTipText     =   "Input nomor rekening lalu tekan ENTER!"
      Top             =   360
      Width           =   3015
   End
   Begin VB.Label Label2 
      Caption         =   "Jumlah"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Nomor Rekening"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1935
   End
End
Attribute VB_Name = "TarikFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CancelBtn_Click()
Unload Me
End Sub

Private Sub Form_Load()

'buka recordset saving dan transaksi
Cmd.CommandText = "SELECT*FROM saving"
Cmd.ActiveConnection = MyDB
RcSaving.Open Cmd, , adOpenDynamic, adLockOptimistic

Cmd.CommandText = "SELECT*FROM transaksi"
RcTransaksi.Open Cmd, , adOpenDynamic, adLockOptimistic

'gerak maju ke depan (saving) dan ke belakang (transaksi)
RcSaving.MoveFirst
RcTransaksi.MoveLast

Me.Caption = JudulTarik
End Sub

Private Sub Form_Unload(Cancel As Integer)
RcSaving.Close
RcTransaksi.Close
End Sub

Private Sub OKBtn_Click()
Dim NorekCrit As String
Dim Ammount As Long
Dim Current_Saldo As Single
Dim x As Long

If TxtSetor.Text = "" Then Exit Sub
If TxtNoRek.Text = "" Then Exit Sub

NorekCrit = TxtNoRek.Text
Ammount = Val(TxtSetor.Text)

RcSaving.Find "norek='" + Trim(NorekCrit) + "'"

'kalo tidak ketemu
If RcSaving.EOF Then
    MsgBox "Tidak ada nomor rekening: " + NorekCrit, vbCritical, "Tidak ketemu!"
    Exit Sub
End If

'kalo ketemu... dilanjut!!
Waktu = Now
x = Val(RcTransaksi!tid)
x = x + 1

Current_Saldo = RcSaving!saldo

If Ammount > (Current_Saldo - 25000) Then
    MsgBox "Penarikan dana terlalu besar!" + vbCrLf + "Penarikan maks. sebesar " + Format(Current_Saldo - 25000, "###,###,###,###.00"), vbCritical, "Duit Tong-PeZZ!"
    Exit Sub
End If


'masukkan ke database
With RcTransaksi
    .AddNew
    !tid = Trim(Str(x))
    !NoRek = NorekCrit
    !tgl = Waktu
    !jenis = "debet"
    !Jumlah = Ammount
    .Update
End With


With RcSaving
    !NoRek = NorekCrit
    !saldo = Current_Saldo - Ammount
    .Update
End With


Unload Me
End Sub



