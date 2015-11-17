VERSION 5.00
Begin VB.Form frmEditPenjualan 
   Caption         =   "Edit Data Penjualan"
   ClientHeight    =   3690
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5490
   LinkTopic       =   "Form1"
   ScaleHeight     =   3690
   ScaleWidth      =   5490
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtHargaBarang 
      DataField       =   "Harga Barang"
      DataMember      =   "Penjualan"
      DataSource      =   "DE"
      Height          =   285
      Left            =   2040
      TabIndex        =   18
      Top             =   2020
      Width           =   1320
   End
   Begin VB.TextBox txtBanyaknyabarang 
      DataField       =   "Banyaknya barang"
      DataMember      =   "Penjualan"
      DataSource      =   "DE"
      Height          =   285
      Left            =   2040
      TabIndex        =   16
      Top             =   1640
      Width           =   660
   End
   Begin VB.TextBox txtKodeBarang 
      DataField       =   "Kode Barang"
      DataMember      =   "Penjualan"
      DataSource      =   "DE"
      Height          =   285
      Left            =   2040
      TabIndex        =   14
      Top             =   1260
      Width           =   990
   End
   Begin VB.TextBox txtKodePelanggan 
      DataField       =   "Kode Pelanggan"
      DataMember      =   "Penjualan"
      DataSource      =   "DE"
      Height          =   285
      Left            =   2040
      TabIndex        =   12
      Top             =   880
      Width           =   990
   End
   Begin VB.TextBox txtTanggalBon 
      DataField       =   "Tanggal Bon"
      DataMember      =   "Penjualan"
      DataSource      =   "DE"
      Height          =   285
      Left            =   2040
      TabIndex        =   10
      Top             =   500
      Width           =   1650
   End
   Begin VB.TextBox txtNoBon 
      DataField       =   "No Bon"
      DataMember      =   "Penjualan"
      DataSource      =   "DE"
      Height          =   285
      Left            =   2040
      TabIndex        =   8
      Top             =   120
      Width           =   1650
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Height          =   375
      Left            =   840
      TabIndex        =   6
      Top             =   3240
      Width           =   1695
   End
   Begin VB.CommandButton cmdTutup 
      Caption         =   "Tutup"
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   2400
      Width           =   4815
      Begin VB.CommandButton cmdNext 
         Caption         =   ">"
         Default         =   -1  'True
         Height          =   372
         Left            =   2520
         TabIndex        =   4
         Top             =   240
         Width           =   972
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "<"
         Height          =   372
         Left            =   1320
         TabIndex        =   3
         Top             =   240
         Width           =   972
      End
      Begin VB.CommandButton cmdFirst 
         Caption         =   "I<"
         Height          =   372
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   972
      End
      Begin VB.CommandButton cmdLast 
         Caption         =   ">I"
         Height          =   372
         Left            =   3720
         TabIndex        =   1
         Top             =   240
         Width           =   972
      End
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Harga Barang:"
      Height          =   255
      Index           =   5
      Left            =   435
      TabIndex        =   17
      Top             =   2070
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Banyaknya barang:"
      Height          =   255
      Index           =   4
      Left            =   435
      TabIndex        =   15
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Kode Barang:"
      Height          =   255
      Index           =   3
      Left            =   435
      TabIndex        =   13
      Top             =   1305
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Kode Pelanggan:"
      Height          =   255
      Index           =   2
      Left            =   435
      TabIndex        =   11
      Top             =   930
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Tanggal Bon:"
      Height          =   255
      Index           =   1
      Left            =   435
      TabIndex        =   9
      Top             =   540
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "No Bon:"
      Height          =   255
      Index           =   0
      Left            =   435
      TabIndex        =   7
      Top             =   165
      Width           =   1815
   End
End
Attribute VB_Name = "frmEditPenjualan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Proyek     :Menginput data Penjualan
'tanggal    :22 November 2001
'Programmer :Kok Yung
'Descripsi  :Menampilakan, menambah, dan menghapus data pada database AVB2,
'            menggunakan ADO dan DE
                                   
Option Explicit

Private Sub cmdEdit_Click()
     'Edit record yang anda sedang ditampilkan
    On Error GoTo HandleErrors 'jalankan penyaringan kesalahan / error utk penanggannan kesalahan
    
    DE.rsPenjualan.Update             'Edit record
   
cmdEdit_Click_Exit:
    Exit Sub

HandleErrors: 'penyaringan kesalahan / error
    Dim strMessage As String
    Dim errDBError As ADODB.Error
    
    For Each errDBError In DE.conAVB.Errors
        strMessage = strMessage & errDBError.Description & vbCrLf
    Next
    MsgBox strMessage, vbExclamation, " Data Kembar"
    On Error GoTo 0     'matikan fungsi penyaringan kesalahan / error
End Sub
Private Sub cmdFirst_Click()
    'Move to first record
    On Error Resume Next
    DE.rsPenjualan.MoveFirst
End Sub

Private Sub cmdLast_Click()
    'Move to last record
    
    On Error Resume Next
    DE.rsPenjualan.MoveLast
End Sub

Private Sub cmdNext_Click()
    'Move to next record
    
    On Error Resume Next
    With DE.rsPenjualan
        .MoveNext
        If .EOF Then
            .MoveLast
        End If
    End With
End Sub

Private Sub cmdPrevious_Click()
    'Move to previous record
    
    On Error Resume Next
    With DE.rsPenjualan
        .MovePrevious
        If .BOF Then
            .MoveFirst
        End If
    End With
End Sub

Private Sub cmdTutup_Click()
'Keluar dari proyek
    Unload Me
End Sub


