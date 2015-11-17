VERSION 5.00
Begin VB.Form frmEditPembelian 
   Caption         =   "Edit Data Pembelian"
   ClientHeight    =   3765
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4725
   LinkTopic       =   "Form1"
   ScaleHeight     =   3765
   ScaleWidth      =   4725
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtBanyaknyaBarang 
      DataField       =   "Banyaknya Barang"
      DataMember      =   "Pembelian"
      DataSource      =   "DE"
      Height          =   285
      Left            =   2445
      TabIndex        =   18
      Top             =   2040
      Width           =   330
   End
   Begin VB.TextBox txtHargaSatuan 
      DataField       =   "Harga Satuan"
      DataMember      =   "Pembelian"
      DataSource      =   "DE"
      Height          =   285
      Left            =   2445
      TabIndex        =   16
      Top             =   1650
      Width           =   1320
   End
   Begin VB.TextBox txtKodeBarang 
      DataField       =   "Kode Barang"
      DataMember      =   "Pembelian"
      DataSource      =   "DE"
      Height          =   285
      Left            =   2445
      TabIndex        =   14
      Top             =   1275
      Width           =   990
   End
   Begin VB.TextBox txtKodePemasok 
      DataField       =   "Kode Pemasok"
      DataMember      =   "Pembelian"
      DataSource      =   "DE"
      Height          =   285
      Left            =   2445
      TabIndex        =   12
      Top             =   900
      Width           =   990
   End
   Begin VB.TextBox txtTanggalFaktur 
      DataField       =   "Tanggal Faktur"
      DataMember      =   "Pembelian"
      DataSource      =   "DE"
      Height          =   285
      Left            =   2445
      TabIndex        =   10
      Top             =   510
      Width           =   1650
   End
   Begin VB.TextBox txtNoFaktur 
      DataField       =   "No Faktur"
      DataMember      =   "Pembelian"
      DataSource      =   "DE"
      Height          =   285
      Left            =   2445
      TabIndex        =   8
      Top             =   135
      Width           =   1650
   End
   Begin VB.CommandButton cmdTutup 
      Caption         =   "Tutup"
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   3120
      Width           =   1575
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   2400
      Width           =   3615
      Begin VB.CommandButton cmdLast 
         Caption         =   ">I"
         Height          =   255
         Left            =   2760
         TabIndex        =   4
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdFirst 
         Caption         =   "I<"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "<"
         Height          =   255
         Left            =   1080
         TabIndex        =   2
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   ">"
         Default         =   -1  'True
         Height          =   255
         Left            =   1920
         TabIndex        =   1
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Banyaknya Barang:"
      Height          =   255
      Index           =   5
      Left            =   600
      TabIndex        =   17
      Top             =   2085
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Harga Satuan:"
      Height          =   255
      Index           =   4
      Left            =   600
      TabIndex        =   15
      Top             =   1695
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Kode Barang:"
      Height          =   255
      Index           =   3
      Left            =   600
      TabIndex        =   13
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Kode Pemasok:"
      Height          =   255
      Index           =   2
      Left            =   600
      TabIndex        =   11
      Top             =   945
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Tanggal Faktur:"
      Height          =   255
      Index           =   1
      Left            =   600
      TabIndex        =   9
      Top             =   555
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "No Faktur:"
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   7
      Top             =   180
      Width           =   1815
   End
End
Attribute VB_Name = "frmEditPembelian"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Proyek     :Menginput data Pembelian
'tanggal    :22 November 2001
'Programmer :Kok Yung
'Descripsi  :Menampilakan, menambah, dan menghapus data pada database AVB2,
'            menggunakan ADO dan DE
                                   
Option Explicit

Private Sub cmdEdit_Click()
     'Edit record yang anda sedang ditampilkan
    On Error GoTo HandleErrors 'jalankan penyaringan kesalahan / error utk penanggannan kesalahan
    
    DE.rsPembelian.Update             'Edit record
   
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
    DE.rsPembelian.MoveFirst
End Sub

Private Sub cmdLast_Click()
    'Move to last record
    
    On Error Resume Next
    DE.rsPembelian.MoveLast
End Sub

Private Sub cmdNext_Click()
    'Move to next record
    
    On Error Resume Next
    With DE.rsPembelian
        .MoveNext
        If .EOF Then
            .MoveLast
        End If
    End With
End Sub

Private Sub cmdPrevious_Click()
    'Move to previous record
    
    On Error Resume Next
    With DE.rsPembelian
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

