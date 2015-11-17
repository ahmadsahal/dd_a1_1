VERSION 5.00
Begin VB.Form frmBarang 
   Caption         =   "Data Barang"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5580
   LinkTopic       =   "Form1"
   ScaleHeight     =   2955
   ScaleWidth      =   5580
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtJumlahBarang 
      DataField       =   "Jumlah Barang"
      DataMember      =   "Barang"
      DataSource      =   "DE"
      Height          =   285
      Left            =   1440
      TabIndex        =   16
      Top             =   1200
      Width           =   660
   End
   Begin VB.TextBox txtHargaBarang 
      DataField       =   "Harga Barang"
      DataMember      =   "Barang"
      DataSource      =   "DE"
      Height          =   285
      Left            =   1410
      TabIndex        =   14
      Top             =   840
      Width           =   1320
   End
   Begin VB.TextBox txtNamaBarang 
      DataField       =   "Nama Barang"
      DataMember      =   "Barang"
      DataSource      =   "DE"
      Height          =   285
      Left            =   1440
      TabIndex        =   12
      Top             =   450
      Width           =   3375
   End
   Begin VB.TextBox txtKodeBarang 
      DataField       =   "Kode Barang"
      DataMember      =   "Barang"
      DataSource      =   "DE"
      Height          =   285
      Left            =   1410
      TabIndex        =   10
      Top             =   75
      Width           =   990
   End
   Begin VB.Frame fraNavigation 
      Height          =   735
      Left            =   840
      TabIndex        =   4
      Top             =   1560
      Width           =   3495
      Begin VB.CommandButton cmdLast 
         Caption         =   ">I"
         Height          =   375
         Left            =   2640
         TabIndex        =   8
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   ">"
         Default         =   -1  'True
         Height          =   375
         Left            =   1800
         TabIndex        =   7
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "<"
         Height          =   375
         Left            =   960
         TabIndex        =   6
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdFirst 
         Caption         =   "I<"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdHapus 
      Caption         =   "&Hapus"
      Height          =   372
      Left            =   2640
      TabIndex        =   3
      Top             =   2400
      Width           =   972
   End
   Begin VB.CommandButton cmdTambah 
      Caption         =   "&Tambah"
      CausesValidation=   0   'False
      Height          =   372
      Left            =   240
      TabIndex        =   2
      Top             =   2400
      Width           =   972
   End
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "&Simpan"
      Enabled         =   0   'False
      Height          =   372
      Left            =   1440
      TabIndex        =   1
      Top             =   2400
      Width           =   972
   End
   Begin VB.CommandButton cmdTutup 
      Caption         =   "Tutup"
      Height          =   375
      Left            =   3840
      TabIndex        =   0
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Jumlah Barang:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   15
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Harga Barang:"
      Height          =   255
      Index           =   2
      Left            =   165
      TabIndex        =   13
      Top             =   885
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Nama Barang:"
      Height          =   255
      Index           =   1
      Left            =   165
      TabIndex        =   11
      Top             =   495
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Kode Barang:"
      Height          =   255
      Index           =   0
      Left            =   165
      TabIndex        =   9
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmBarang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Proyek     :Menginput data Barang
'tanggal    :22 November 2001
'Programmer :Kok Yung
'Descripsi  :Menampilakan, menambah, dan menghapus data pada database AVB2,
'            menggunakan ADO dan DE
                                   
Option Explicit
Private Sub cmdTambah_Click()
 'Add a new record
    
    On Error GoTo HandleError        'matikan penyaringan kesalahan / error utk penanganan kesalahan
    If cmdTambah.Caption = "&Tambah" Then
        DE.rsBarang.AddNew          'kosongkan field untuk record data baru
        NonaktifkanKontrol           'tombol-tombol pengerak Record dimatikan
        cmdTambah.Caption = "&Batal" 'mengganti tulisan tambah menjadi batal
        cmdSimpan.Enabled = True     'aktifkan tombol simpan
        txtKodeBarang.Locked = False 'buka text kode Barang
        
    Else
        DE.rsBarang.CancelUpdate     'Batalkan proses penambahan data
        txtKodeBarang.Locked = True  'kunci text kode Barang
        AktifkanTombol                'semua tombol penggerak record diaktifkan
        cmdTambah.Caption = "&Tambah"    'ubah tulisan batal menjadi Tambah
        cmdSimpan.Enabled = False       'Disable the Save button
        DE.rsBarang.MoveLast         'Pindah ke record data terakhir
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
    With DE.rsBarang
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
    DE.rsBarang.Update             'simpan record
    txtKodeBarang.Locked = True    'kunci text kode Barang
    AktifkanTombol      'semua tombol penggerak record diaktifkan
    cmdSimpan.Enabled = False 'tombol simpan dinonAktifkan
    cmdTambah.Caption = "&Tambah"  ' ubah kembali tulisan menjadi tambah dari batal
    DE.rsBarang.MoveFirst 'dg DE field ini tidak boleh kosong karena nantinya data tidak dapat di delete
    DE.rsBarang.MoveLast 'dg DE field ini tidak boleh kosong karena nantinya data tidak dapat di delete
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
    DE.rsBarang.MoveFirst
End Sub

Private Sub cmdLast_Click()
    'Move to last record
    
    On Error Resume Next
    DE.rsBarang.MoveLast
End Sub

Private Sub cmdNext_Click()
    'Move to next record
    
    On Error Resume Next
    With DE.rsBarang
        .MoveNext
        If .EOF Then
            .MoveLast
        End If
    End With
End Sub

Private Sub cmdPrevious_Click()
    'Move to previous record
    
    On Error Resume Next
    With DE.rsBarang
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


Private Sub txtHargaBarang_KeyPress(KeyAscii As Integer)
    'hanya boleh diisi angka atau backspace
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack) Then
       Beep
       KeyAscii = 0
    End If
End Sub
Private Sub txtKodeBarang_KeyPress(KeyAscii As Integer)
    'mengubah huruf kecil jadi huruf besar
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub txtNamaBarang_KeyPress(KeyAscii As Integer)
    'mengubah huruf kecil jadi huruf besar
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

