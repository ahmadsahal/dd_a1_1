VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmInJual 
   Caption         =   "Penjualan"
   ClientHeight    =   4470
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   4470
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtJumlah 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Left            =   1440
      MultiLine       =   -1  'True
      TabIndex        =   16
      Top             =   2640
      Width           =   2028
   End
   Begin VB.Frame fraNavigation 
      Height          =   735
      Left            =   480
      TabIndex        =   11
      Top             =   3120
      Width           =   3495
      Begin VB.CommandButton cmdLast 
         Caption         =   ">I"
         Height          =   375
         Left            =   2640
         TabIndex        =   15
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   ">"
         Default         =   -1  'True
         Height          =   375
         Left            =   1800
         TabIndex        =   14
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "<"
         Height          =   375
         Left            =   960
         TabIndex        =   13
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdFirst 
         Caption         =   "I<"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdHapus 
      Caption         =   "&Hapus"
      Height          =   372
      Left            =   2280
      TabIndex        =   10
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton cmdTambah 
      Caption         =   "&Tambah"
      CausesValidation=   0   'False
      Height          =   372
      Left            =   360
      TabIndex        =   9
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "&Simpan"
      Enabled         =   0   'False
      Height          =   372
      Left            =   1320
      TabIndex        =   8
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton cmdTutup 
      Caption         =   "Tutup"
      Height          =   375
      Left            =   3240
      TabIndex        =   7
      Top             =   3960
      Width           =   855
   End
   Begin VB.TextBox txtNoBon 
      DataField       =   "No Bon"
      DataMember      =   "Penjualan"
      DataSource      =   "DE"
      Height          =   285
      Left            =   1440
      TabIndex        =   6
      Top             =   120
      Width           =   1650
   End
   Begin VB.TextBox txtTanggalBon 
      DataField       =   "Tanggal Bon"
      DataMember      =   "Penjualan"
      DataSource      =   "DE"
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Top             =   480
      Width           =   1170
   End
   Begin VB.TextBox txtBanyaknyaBarang 
      DataField       =   "Banyaknya barang"
      DataMember      =   "Penjualan"
      DataSource      =   "DE"
      Height          =   285
      Left            =   1440
      TabIndex        =   4
      Top             =   2280
      Width           =   330
   End
   Begin VB.TextBox txtNamaBarang 
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Top             =   1560
      Width           =   3015
   End
   Begin VB.TextBox txtHargaBarang 
      DataField       =   "Harga Barang"
      DataMember      =   "Penjualan"
      DataSource      =   "DE"
      Height          =   285
      Left            =   1440
      TabIndex        =   2
      Top             =   1920
      Width           =   1200
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Bindings        =   "frmInJual.frx":0000
      DataField       =   "Kode Pelanggan"
      DataMember      =   "Penjualan"
      DataSource      =   "DE"
      Height          =   315
      Left            =   1440
      TabIndex        =   0
      Top             =   840
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Kode Pelanggan"
      BoundColumn     =   "Nama Pemasok"
      Text            =   "DataCombo2"
      Object.DataMember      =   "Pelanggan"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "frmInJual.frx":001B
      DataField       =   "Kode Barang"
      DataMember      =   "Penjualan"
      DataSource      =   "DE"
      Height          =   315
      Left            =   1440
      TabIndex        =   1
      Top             =   1200
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Kode Barang"
      Text            =   "DataCombo1"
      Object.DataMember      =   "Barang"
   End
   Begin VB.Label Label1 
      Caption         =   "Kode Pelanggan"
      Height          =   210
      Left            =   120
      TabIndex        =   25
      Top             =   840
      Width           =   1545
   End
   Begin VB.Label Label3 
      Caption         =   "Kode Barang"
      Height          =   210
      Left            =   120
      TabIndex        =   24
      Top             =   1200
      Width           =   1065
   End
   Begin VB.Label Label4 
      Caption         =   "Nama Barang"
      Height          =   210
      Left            =   120
      TabIndex        =   23
      Top             =   1560
      Width           =   1590
   End
   Begin VB.Label Label5 
      Caption         =   "Harga Satuan"
      Height          =   210
      Left            =   120
      TabIndex        =   22
      Top             =   1920
      Width           =   1065
   End
   Begin VB.Label Label6 
      Caption         =   "Banyak"
      Height          =   210
      Left            =   120
      TabIndex        =   21
      Top             =   2280
      Width           =   1065
   End
   Begin VB.Label Label7 
      Caption         =   "No Bon"
      Height          =   210
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   1065
   End
   Begin VB.Label Label8 
      Caption         =   "Tanggal  Bon"
      Height          =   210
      Left            =   120
      TabIndex        =   19
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label9 
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
      Height          =   300
      Left            =   120
      TabIndex        =   18
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label10 
      Caption         =   "dd-mm-yyyy"
      Height          =   255
      Left            =   2640
      TabIndex        =   17
      Top             =   480
      Width           =   1335
   End
End
Attribute VB_Name = "frmInJual"
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
Dim nilai As Integer
Private Sub cmdTambah_Click()
 'Add a new record
    
    On Error GoTo HandleError        'matikan penyaringan kesalahan / error utk penanganan kesalahan
    If cmdTambah.Caption = "&Tambah" Then
        DE.rsPenjualan.AddNew          'kosongkan field untuk record data baru
        NonaktifkanKontrol           'tombol-tombol pengerak Record dimatikan
        cmdTambah.Caption = "&Batal" 'mengganti tulisan tambah menjadi batal
        cmdSimpan.Enabled = True     'aktifkan tombol simpan
        txtNoBon.Locked = False 'buka text kode Penjualan
        txtNamaBarang.Text = ""
    Else
        DE.rsPenjualan.CancelUpdate     'Batalkan proses penambahan data
        txtNoBon.Locked = True  'kunci text kode Penjualan
        AktifkanTombol                'semua tombol penggerak record diaktifkan
        cmdTambah.Caption = "&Tambah"    'ubah tulisan batal menjadi Tambah
        cmdSimpan.Enabled = False       'Disable the Save button
        DE.rsPenjualan.MoveLast         'Pindah ke record data terakhir
    End If
    
cmdTambah_Click_Exit:
    Exit Sub
    
HandleError:
    MsgBox "Proses tidak bisa dikerjakan.", vbInformation, "Perhatian"
    On Error GoTo 0 'matikan fungsi penyaringan kesalahan / error
End Sub
Private Sub cmdHapus_Click()
 'hapus record yang sedang ditampilkan /aktif
    On Error GoTo HandleError 'jalankan perangkap kesalahan jika terjadi kesalahan
    Dim NILAI1 As Integer
    NILAI1 = nilai + Val(txtBanyaknyaBarang.Text)
    DE.rsBarang![Jumlah Barang] = NILAI1
    DE.rsBarang.Update
    
    With DE.rsPenjualan
        .Delete                             'menghapus record yang aktif
        .MoveNext                           'pindah ke record selanjutnya
        If .EOF Then                        'jika data terakhir yang dihapus
            .MovePrevious                   'maju  1 record ke arah record pertama
            If .BOF Then            'jika nilai BOF dan EOF benar,recordset kosong
                MsgBox "Data Kosong.", vbInformation, "Perhatian" 'tampilkan pesan data kosong
                NonaktifkanKontrol 'tombol tombol pengerak record dinonaktifkan
            End If
        End If
    End With
    
cmdHapus_Click_Exit:
    Exit Sub
    
HandleError:
    MsgBox "Data tidak dapat diproses.", vbInformation, "Perhatian"
    On Error GoTo 0 'matikan fungsi penyaringan kesalahan / error
End Sub

Private Sub cmdSimpan_Click()
     'Simpan record yang anda sedang ditampilkan
    On Error GoTo HandleErrors 'jalankan penyaringan kesalahan / error utk penanggannan kesalahan
    Dim NILAI1 As Integer
    NILAI1 = nilai - Val(txtBanyaknyaBarang.Text)
    DE.rsBarang![Jumlah Barang] = NILAI1
    DE.rsBarang.Update
    DE.rsPenjualan.Update             'simpan record
    txtNoBon.Locked = True    'kunci text kode Penjualan
    AktifkanTombol      'semua tombol penggerak record diaktifkan
    cmdSimpan.Enabled = False 'tombol simpan dinonAktifkan
    cmdTambah.Caption = "&Tambah"  ' ubah kembali tulisan menjadi tambah dari batal
    
cmdSimpan_Click_Exit:
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

Private Sub NonaktifkanKontrol()
    'mengunci / mematikan tombol-tombol pengerak record
    cmdNext.Enabled = False
    cmdPrevious.Enabled = False
    cmdFirst.Enabled = False
    cmdLast.Enabled = False
    cmdHapus.Enabled = False
End Sub
Private Sub AktifkanTombol()
    'membuka / mengaktifkan tombol-tombol pengerak record
    cmdNext.Enabled = True
    cmdPrevious.Enabled = True
    cmdFirst.Enabled = True
    cmdLast.Enabled = True
    cmdHapus.Enabled = True
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


Private Sub DataCombo1_Change()
 Dim strSearch   As String
    Dim vntBookMark As Variant
    
        strSearch = "[Kode Barang] = '" & DataCombo1 & "'"
        With DE.rsBarang
            vntBookMark = .Bookmark
            .MoveFirst
            .Find strSearch
            If .EOF Then
               
                .Bookmark = vntBookMark
            End If
            txtNamaBarang.Text = ![Nama Barang]
    End With
End Sub


Private Sub DataCombo1_Click(Area As Integer)
    
    Dim strSearch   As String
    Dim vntBookMark As Variant
    
        strSearch = "[Kode Barang] = '" & DataCombo1 & "'"
        With DE.rsBarang
            vntBookMark = .Bookmark  'simpan record yang aktif serta dipilih
            .MoveFirst
            .Find strSearch
            If .EOF Then

                .Bookmark = vntBookMark 'Kembali ke record sebelumnya
            End If
            txtHargaBarang.Text = ![Harga Barang]
            txtNamaBarang.Text = ![Nama Barang]
            nilai = ![Jumlah Barang]
        End With
End Sub





Private Sub txtBanyaknyaBarang_KeyPress(KeyAscii As Integer)
    'hanya boleh diisi angka atau backspace
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack) Then
       Beep
       KeyAscii = 0
    End If
End Sub

Private Sub txtHargaBarang_Change()
    Dim Jumlah As Single
    Jumlah = Val(txtHargaBarang.Text) * Val(txtBanyaknyaBarang.Text)
    On Error GoTo Salah
    txtJumlah.Text = Format(Jumlah, "Rp ###,###,###") & ",-"
    On Error GoTo 0
    Exit Sub
Salah:
End Sub

Private Sub txtBanyaknyaBarang_Change()
    Dim Jumlah As Single
    Jumlah = Val(txtHargaBarang.Text) * Val(txtBanyaknyaBarang.Text)
    On Error GoTo Salah
    txtJumlah.Text = Format(Jumlah, "Rp ###,###,###") & ",-"
    On Error GoTo 0
    Exit Sub
Salah:
End Sub

Private Sub txtHargaBarang_KeyPress(KeyAscii As Integer)
    'hanya boleh diisi angka atau backspace
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack) Then
       Beep
       KeyAscii = 0
    End If
End Sub
Private Sub txtNoBon_KeyPress(KeyAscii As Integer)
    'mengubah huruf kecil jadi huruf besar
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub



