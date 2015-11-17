VERSION 5.00
Begin VB.Form Form28 
   Caption         =   "Mengganti dan membetulkan Data"
   ClientHeight    =   3735
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4455
   LinkTopic       =   "Form1"
   ScaleHeight     =   3735
   ScaleWidth      =   4455
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTutup 
      Caption         =   "Tutup"
      Height          =   615
      Left            =   2280
      TabIndex        =   14
      Top             =   3000
      Width           =   1815
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Enabled         =   0   'False
      Height          =   615
      Left            =   120
      TabIndex        =   13
      Top             =   3000
      Width           =   1815
   End
   Begin VB.TextBox txtBanyaknyaBarang 
      DataField       =   "Banyaknya Barang"
      DataMember      =   "Pembelian"
      DataSource      =   "DE"
      Enabled         =   0   'False
      Height          =   285
      Left            =   1935
      TabIndex        =   12
      Top             =   2565
      Width           =   1290
   End
   Begin VB.TextBox txtHargaSatuan 
      DataField       =   "Harga Satuan"
      DataMember      =   "Pembelian"
      DataSource      =   "DE"
      Enabled         =   0   'False
      Height          =   285
      Left            =   1935
      TabIndex        =   10
      Top             =   2175
      Width           =   1320
   End
   Begin VB.TextBox txtKodeBarang 
      DataField       =   "Kode Barang"
      DataMember      =   "Pembelian"
      DataSource      =   "DE"
      Enabled         =   0   'False
      Height          =   285
      Left            =   1935
      TabIndex        =   8
      Top             =   1800
      Width           =   1950
   End
   Begin VB.TextBox txtKodePemasok 
      DataField       =   "Kode Pemasok"
      DataMember      =   "Pembelian"
      DataSource      =   "DE"
      Enabled         =   0   'False
      Height          =   285
      Left            =   1935
      TabIndex        =   6
      Top             =   1425
      Width           =   1950
   End
   Begin VB.TextBox txtTanggalFaktur 
      DataField       =   "Tanggal Faktur"
      DataMember      =   "Pembelian"
      DataSource      =   "DE"
      Enabled         =   0   'False
      Height          =   285
      Left            =   1935
      TabIndex        =   4
      Top             =   1035
      Width           =   1980
   End
   Begin VB.TextBox txtNoFaktur 
      DataField       =   "No Faktur"
      DataMember      =   "Pembelian"
      DataSource      =   "DE"
      Enabled         =   0   'False
      Height          =   285
      Left            =   1935
      TabIndex        =   2
      Top             =   660
      Width           =   1980
   End
   Begin VB.ComboBox cboNoFaktur 
      Height          =   315
      Left            =   840
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Banyaknya Barang:"
      Height          =   255
      Index           =   5
      Left            =   210
      TabIndex        =   11
      Top             =   2610
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Harga Satuan:"
      Height          =   255
      Index           =   4
      Left            =   210
      TabIndex        =   9
      Top             =   2220
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Kode Barang:"
      Height          =   255
      Index           =   3
      Left            =   210
      TabIndex        =   7
      Top             =   1845
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Kode Pemasok:"
      Height          =   255
      Index           =   2
      Left            =   210
      TabIndex        =   5
      Top             =   1470
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Tanggal Faktur:"
      Height          =   255
      Index           =   1
      Left            =   210
      TabIndex        =   3
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "No Faktur:"
      Height          =   255
      Index           =   0
      Left            =   210
      TabIndex        =   1
      Top             =   705
      Width           =   1815
   End
   Begin VB.Menu mnuKeluar 
      Caption         =   "Keluar"
   End
End
Attribute VB_Name = "Form28"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Nama Program   : Ado Database
'Programmer     : Kok Yung
'Tanggal        : 11/2001
'Purpose        : mengedit data menggunakan objek visual Data Environment
'Folder         : ADO/BAB2/2_Meng-editData

Option Explicit

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub Form_Load()
'mengisi combo box

    On Error GoTo HandleError
    With DE
        Do Until .rsPembelian.EOF 'mengisi data ke combo
            If .rsPembelian![No Faktur] <> "" Then
                cboNoFaktur.AddItem .rsPembelian![No Faktur]
            End If
            .rsPembelian.MoveNext
        Loop
        .rsPembelian.MoveFirst
    End With
    
Form_Load_Exit:
    Exit Sub
    
HandleError:
    MsgBox "Data No Faktur tidak dapat ditampilkan", vbInformation, "Perhatian"
    On Error GoTo 0
End Sub

Private Sub cboNoFaktur_Click()
    'cari record yang dipilih kemudian tampilkan recordnya
    Dim strSearch As String
    
    strSearch = "[No Faktur] ='" & cboNoFaktur.Text & "'"
    With DE.rsPembelian
        .MoveFirst      'mulai dari recordset awal
        .Find strSearch
        If .EOF Then    'tampilkan pesan jika data tidak ditemukan
            MsgBox "No Faktur tidak ada", vbExclamation, "Perhatian"
        End If
    End With
    aktifkantombol
End Sub
Private Sub aktifkantombol()
    txtTanggalFaktur.Enabled = True
    txtKodePemasok.Enabled = True
    txtKodeBarang.Enabled = True
    txtHargaSatuan.Enabled = True
    txtBanyaknyaBarang.Enabled = True
    cmdEdit.Enabled = True
End Sub
Private Sub cmdEdit_Click()
    'Simpan record yang telah ditulis
    On Error GoTo HandleErrors
    
    DE.rsPembelian.Update
    'txtNoFaktur.Locked = True       'kunci text no faktur, kosongkan semua field
    nonaktifkan
        
cmdSimpan_Click_Exit:
    Exit Sub

HandleErrors:
    Dim strMessage As String
    Dim errDBError As ADODB.Error
    
    For Each errDBError In DE.conAVB.Errors
        strMessage = strMessage & errDBError.Description & vbCrLf
    Next
    MsgBox strMessage, vbExclamation, "Data yang ditambahkan sudah ada"
    On Error GoTo 0
End Sub
Private Sub nonaktifkan()
    txtNoFaktur.Enabled = False
    txtTanggalFaktur.Enabled = False
    txtKodePemasok.Enabled = False
    txtKodeBarang.Enabled = False
    txtHargaSatuan.Enabled = False
    txtBanyaknyaBarang.Enabled = False
    cmdEdit.Enabled = False
End Sub

Private Sub mnuKeluar_Click()
    'tutup program
    Unload Me
End Sub


