VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Begin VB.Form LogFrm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Transaksi Log"
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11910
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   5295
      Left            =   5160
      TabIndex        =   14
      Top             =   240
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   9340
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      ScrollBars      =   2
   End
   Begin VB.TextBox Text1 
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
      Left            =   240
      TabIndex        =   13
      Text            =   "NoRek"
      ToolTipText     =   "Ketik secara manual no rekening nasabah, lalu tekan ENTER!"
      Top             =   840
      Width           =   4695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Data Nasabah"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   240
      TabIndex        =   2
      Top             =   1920
      Width           =   4695
      Begin VB.Label LblSaldo 
         Caption         =   "Saldo akhir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1800
         TabIndex        =   12
         Top             =   3120
         Width           =   2775
      End
      Begin VB.Label Label10 
         Caption         =   "Saldo akhir"
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
         Left            =   120
         TabIndex        =   11
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label LblPhone 
         Caption         =   "Telepon"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1800
         TabIndex        =   10
         Top             =   2640
         Width           =   2655
      End
      Begin VB.Label LblKota 
         Caption         =   "Kota"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1800
         TabIndex        =   9
         Top             =   2160
         Width           =   2775
      End
      Begin VB.Label LblAlamat 
         Caption         =   "Alamat"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   975
         Left            =   1800
         TabIndex        =   8
         Top             =   960
         Width           =   2655
      End
      Begin VB.Label LblNama 
         Caption         =   "Nama"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1800
         TabIndex        =   7
         Top             =   480
         Width           =   2655
      End
      Begin VB.Label Label5 
         Caption         =   "Telepon"
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
         Left            =   120
         TabIndex        =   6
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Kota"
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
         Left            =   120
         TabIndex        =   5
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Alamat"
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
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Nama"
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
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   0
      ToolTipText     =   "Pilih no rekening nasabah"
      Top             =   1320
      Width           =   4695
   End
   Begin VB.Label Label1 
      Caption         =   "No Rekening"
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
      TabIndex        =   1
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "LogFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cari(Nomor As String)
Dim NoRek As String
Dim x As Single

NoRek = Nomor

RcProfile.MoveFirst
RcProfile.Find "norek ='" + NoRek + "'"

If RcProfile.EOF Then
    MsgBox "Tidak ada nomor rekening tersebut!", vbCritical, "Kosong!"
    Exit Sub
End If

RcSaving.MoveFirst
RcSaving.Find "norek ='" + NoRek + "'"




'bila suatu no rekening ada, maka minimal pasti ada
'1 x transaksi, setoran awal
Cmd.CommandText = "SELECT * FROM transaksi WHERE norek ='" + Nomor + "'"
Set Cmd.ActiveConnection = MyDB
'Set Rc = Cmd.Execute
RcTransaksi.Open Cmd, , adOpenDynamic, adLockOptimistic

RcTransaksi.MoveFirst
'gunakan trim untuk menghilangkan spasi!!
If Trim(RcTransaksi!jenis) = "kredit" Then
    Grid1.AddItem Str(RcTransaksi!tgl) + Chr(9) + Chr(9) + Format(RcTransaksi!Jumlah, "###,###,###,###.00")
    x = x + RcTransaksi!Jumlah
Else
    Grid1.AddItem Str(RcTransaksi!tgl) + Chr(9) + Format(RcTransaksi!Jumlah, "###,###,###,###.00")
    x = x + RcTransaksi!Jumlah
End If
RcTransaksi.MoveNext

While Not RcTransaksi.EOF
    If Trim(RcTransaksi!jenis) = "kredit" Then
        Grid1.AddItem Str(RcTransaksi!tgl) + Chr(9) + Chr(9) + Format(RcTransaksi!Jumlah, "###,###,###,###.00")
        x = x + RcTransaksi!Jumlah
    Else
        Grid1.AddItem Str(RcTransaksi!tgl) + Chr(9) + Format(RcTransaksi!Jumlah, "###,###,###,###.00")
        x = x - RcTransaksi!Jumlah
    End If
    RcTransaksi.MoveNext
Wend

RcTransaksi.Close

RcSaving!saldo = x
RcSaving.Update

LblNama.Caption = RcProfile!nama
LblAlamat.Caption = RcProfile!alamat
LblKota.Caption = RcProfile!kota
LblPhone.Caption = RcProfile!phone
LblSaldo.Caption = Format(RcSaving!saldo, "###,###,###,###.00")
End Sub
Private Sub Combo1_Click()
Dim x As Long

Grid1.Clear
Grid1.Rows = 2
Grid1.Cols = 3

'set Flexi Grid
Grid1.ColWidth(0) = 1500
Grid1.ColWidth(1) = 2300
Grid1.ColWidth(2) = 2300

'buat text header
Grid1.Row = 0
Grid1.Col = 0
Grid1.Text = "Tgl"

Grid1.Row = 0
Grid1.Col = 1
Grid1.Text = "Debet"

Grid1.Row = 0
Grid1.Col = 2
Grid1.Text = "Kredit"

Cari (Combo1.Text)
End Sub

Private Sub Form_Load()
'bersihkan tampilan
Text1.Text = ""
LblNama.Caption = ""
LblAlamat.Caption = ""
LblKota.Caption = ""
LblPhone.Caption = ""
LblSaldo.Caption = ""
Combo1.Clear
Grid1.Clear

'set Flexi Grid
Grid1.ColWidth(0) = 1500
Grid1.ColWidth(1) = 2300
Grid1.ColWidth(2) = 2300

'buat text header
Grid1.Row = 0
Grid1.Col = 0
Grid1.Text = "Tgl"

Grid1.Row = 0
Grid1.Col = 1
Grid1.Text = "Debet"

Grid1.Row = 0
Grid1.Col = 2
Grid1.Text = "Kredit"


'isi combo 1
'buka recordset tabel Profile, Saving
'eksekusi perintah SQL
    Cmd.CommandText = "SELECT * FROM profile"
    Set Cmd.ActiveConnection = MyDB
    'Set Rc = Cmd.Execute
    RcProfile.Open Cmd, , adOpenDynamic, adLockOptimistic
    
    Cmd.CommandText = "SELECT*FROM saving"
    RcSaving.Open Cmd, , adOpenDynamic, adLockOptimistic
    
RcSaving.MoveFirst
RcProfile.MoveFirst
Combo1.AddItem RcProfile!NoRek
RcProfile.MoveNext

While Not RcProfile.EOF
    Combo1.AddItem RcProfile!NoRek
    RcProfile.MoveNext
Wend

End Sub

Private Sub Form_Unload(Cancel As Integer)
RcProfile.Close
RcSaving.Close
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
'bila ditekan enter maka langsung mencari
If KeyAscii = 13 Then
    Cari (Text1.Text)
End If

End Sub
