VERSION 5.00
Begin VB.Form SetorFrm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Penyetoran"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5700
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   5700
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
      Left            =   3563
      TabIndex        =   5
      Top             =   2040
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
      Left            =   923
      TabIndex        =   4
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox TxtSetor 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   1320
      Width           =   3015
   End
   Begin VB.TextBox TxtNoRek 
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      ToolTipText     =   "Input nomor rekening lalu tekan ENTER!"
      Top             =   480
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
      Top             =   1320
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
      Top             =   480
      Width           =   1935
   End
End
Attribute VB_Name = "SetorFrm"
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

End Sub

Private Sub Form_Unload(Cancel As Integer)

RcTransaksi.Close
RcSaving.Close
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

RcSaving.MoveFirst
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

'masukkan ke database
With RcTransaksi
    .AddNew
    !tid = Trim(Str(x))
    !NoRek = NorekCrit
    !tgl = Waktu
    !jenis = "kredit"
    !Jumlah = Ammount
    .Update
End With



With RcSaving
    !NoRek = NorekCrit
    !saldo = Current_Saldo + Ammount
    .Update
End With


Unload Me
End Sub


