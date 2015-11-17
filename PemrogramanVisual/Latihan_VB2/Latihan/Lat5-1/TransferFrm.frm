VERSION 5.00
Begin VB.Form TransferFrm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Transfer Internal"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5745
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   5745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text3 
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
      Left            =   360
      TabIndex        =   9
      Top             =   3240
      Width           =   5055
   End
   Begin VB.Frame Frame1 
      Caption         =   "Perhatian!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   6
      Top             =   4440
      Width           =   5415
      Begin VB.Label Label3 
         Caption         =   "Pihak pengirim dan penerima harus merupakan nasabah pada Bank Ria."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   5175
      End
   End
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
      Left            =   4200
      TabIndex        =   5
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton KirimBtn 
      Caption         =   "Kirim"
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
      Left            =   360
      TabIndex        =   4
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox Text2 
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
      Left            =   360
      TabIndex        =   3
      Top             =   2040
      Width           =   5055
   End
   Begin VB.TextBox Text1 
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
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   5055
   End
   Begin VB.Label Label4 
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
      Left            =   360
      TabIndex        =   8
      Top             =   2640
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "No Rekening Penerima"
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
      Left            =   360
      TabIndex        =   1
      Top             =   1440
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "No Rekening Pengirim"
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
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   2535
   End
End
Attribute VB_Name = "TransferFrm"
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
RcSaving.Close
RcTransaksi.Close
End Sub

Private Sub KirimBtn_Click()
Dim NorekCrit1 As String
Dim NorekCrit2 As String
Dim Ammount As Long
Dim Current_Saldo As Single
Dim x As Long

'cek textbox
If Text1.Text = "" Then Exit Sub
If Text2.Text = "" Then Exit Sub
If Val(Text3.Text) < 10000 Then
    MsgBox "Transfer minimal sebesar 10.000", vbCritical, "Minimum Transfer"
    Exit Sub
End If

'cek account pengirim
NorekCrit1 = Text1.Text
Ammount = Val(Text3.Text)

RcSaving.MoveFirst
RcSaving.Find "norek='" + Trim(NorekCrit1) + "'"

'kalo tidak ketemu
If RcSaving.EOF Then
    MsgBox "Tidak ada nomor rekening: " + NorekCrit1, vbCritical, "Tidak ketemu!"
    Exit Sub
End If

'ketemu, cek penerima
NorekCrit2 = Text2.Text

RcSaving.MoveFirst
RcSaving.Find "norek='" + Trim(NorekCrit2) + "'"

'kalo tidak ketemu
If RcSaving.EOF Then
    MsgBox "Tidak ada nomor rekening: " + NorekCrit2, vbCritical, "Tidak ketemu!"
    Exit Sub
End If

'ketemu semuanya, langsung proses transfer
'kurangi saldo pengirim
RcSaving.MoveFirst
RcSaving.Find "norek='" + Trim(NorekCrit1) + "'"

'kalo ketemu... dilanjut!!
Waktu = Now
x = Val(RcTransaksi!tid)
x = x + 1

Current_Saldo = RcSaving!saldo

'masukkan ke database
With RcTransaksi
    .AddNew
    !tid = Trim(Str(x))
    !NoRek = NorekCrit1
    !tgl = Waktu
    !jenis = "debet"
    !Jumlah = Ammount
    .Update
End With



With RcSaving
    !NoRek = NorekCrit1
    !saldo = Current_Saldo - Ammount
    .Update
End With

'tambah saldo pengirim
RcSaving.MoveFirst
RcSaving.Find "norek='" + Trim(NorekCrit2) + "'"

Waktu = Now
x = Val(RcTransaksi!tid)
x = x + 1

Current_Saldo = RcSaving!saldo

'masukkan ke database
With RcTransaksi
    .AddNew
    !tid = Trim(Str(x))
    !NoRek = NorekCrit2
    !tgl = Waktu
    !jenis = "kredit"
    !Jumlah = Ammount
    .Update
End With



With RcSaving
    !NoRek = NorekCrit2
    !saldo = Current_Saldo + Ammount
    .Update
End With

MsgBox "Transfer sukses!", vbInformation, "Done!"

Unload Me
End Sub
