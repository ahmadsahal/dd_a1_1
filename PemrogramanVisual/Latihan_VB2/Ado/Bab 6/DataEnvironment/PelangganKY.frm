VERSION 5.00
Begin VB.Form frmPelanggan 
   Caption         =   "Data Customer"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5730
   LinkTopic       =   "Form1"
   ScaleHeight     =   3225
   ScaleWidth      =   5730
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtTeleponpelanggan 
      DataField       =   "Telepon pelanggan"
      DataMember      =   "Pelanggan"
      DataSource      =   "DE"
      Height          =   285
      Left            =   2160
      TabIndex        =   16
      Top             =   1260
      Width           =   1980
   End
   Begin VB.TextBox txtAlamatpelanggan 
      DataField       =   "Alamat pelanggan"
      DataMember      =   "Pelanggan"
      DataSource      =   "DE"
      Height          =   285
      Left            =   2160
      TabIndex        =   14
      Top             =   880
      Width           =   3375
   End
   Begin VB.TextBox txtNamaPelanggan 
      DataField       =   "Nama Pelanggan"
      DataMember      =   "Pelanggan"
      DataSource      =   "DE"
      Height          =   285
      Left            =   2160
      TabIndex        =   12
      Top             =   500
      Width           =   3375
   End
   Begin VB.TextBox txtKodePelanggan 
      DataField       =   "Kode Pelanggan"
      DataMember      =   "Pelanggan"
      DataSource      =   "DE"
      Height          =   285
      Left            =   2160
      TabIndex        =   10
      Top             =   120
      Width           =   990
   End
   Begin VB.Frame fraNavigation 
      Height          =   735
      Left            =   1080
      TabIndex        =   4
      Top             =   1800
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
      Left            =   2880
      TabIndex        =   3
      Top             =   2640
      Width           =   972
   End
   Begin VB.CommandButton cmdTambah 
      Caption         =   "&Tambah"
      CausesValidation=   0   'False
      Height          =   372
      Left            =   480
      TabIndex        =   2
      Top             =   2640
      Width           =   972
   End
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "&Simpan"
      Enabled         =   0   'False
      Height          =   372
      Left            =   1680
      TabIndex        =   1
      Top             =   2640
      Width           =   972
   End
   Begin VB.CommandButton cmdTutup 
      Caption         =   "Tutup"
      Height          =   375
      Left            =   4080
      TabIndex        =   0
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Telepon pelanggan:"
      Height          =   255
      Index           =   3
      Left            =   315
      TabIndex        =   15
      Top             =   1305
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Alamat pelanggan:"
      Height          =   255
      Index           =   2
      Left            =   315
      TabIndex        =   13
      Top             =   930
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Nama Pelanggan:"
      Height          =   255
      Index           =   1
      Left            =   315
      TabIndex        =   11
      Top             =   540
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Kode Pelanggan:"
      Height          =   255
      Index           =   0
      Left            =   315
      TabIndex        =   9
      Top             =   165
      Width           =   1815
   End
End
Attribute VB_Name = "frmPelanggan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Proyek     :Menginput data Pelanggan
'tanggal    :22 November 2001
'Programmer :Kok Yung
'Descripsi  :Menampilakan, menambah, dan menghapus data pada database AVB2,
'            menggunakan ADO dan DE
                                   
Option Explicit
Private Sub cmdTambah_Click()
 'Add a new record
    
    On Error GoTo HandleError        'matikan penyaringan kesalahan / error utk penanganan kesalahan
    If cmdTambah.Caption = "&Tambah" Then
        DE.rsPelanggan.AddNew          'kosongkan field untuk record data baru
        NonaktifkanKontrol           'tombol-tombol pengerak Record dimatikan
        cmdTambah.Caption = "&Batal" 'mengganti tulisan tambah menjadi batal
        cmdSimpan.Enabled = True     'aktifkan tombol simpan
        txtKodePelanggan.Locked = False 'buka text kode Pelanggan
        'SetUpAdd
    Else
        DE.rsPelanggan.CancelUpdate     'Batalkan proses penambahan data
        txtKodePelanggan.Locked = True  'kunci text kode Pelanggan
        AktifkanTombol                'semua tombol penggerak record diaktifkan
        cmdTambah.Caption = "&Tambah"    'ubah tulisan batal menjadi Tambah
        cmdSimpan.Enabled = False       'Disable the Save button
        DE.rsPelanggan.MoveLast         'Pindah ke record data terakhir
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
    With DE.rsPelanggan
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
    
    DE.rsPelanggan.Update             'simpan record
    txtKodePelanggan.Locked = True    'kunci text kode Pelanggan
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
    DE.rsPelanggan.MoveFirst
End Sub

Private Sub cmdLast_Click()
    'Move to last record
    
    On Error Resume Next
    DE.rsPelanggan.MoveLast
End Sub

Private Sub cmdNext_Click()
    'Move to next record
    
    On Error Resume Next
    With DE.rsPelanggan
        .MoveNext
        If .EOF Then
            .MoveLast
        End If
    End With
End Sub

Private Sub cmdPrevious_Click()
    'Move to previous record
    
    On Error Resume Next
    With DE.rsPelanggan
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




Private Sub txtAlamatpelanggan_KeyPress(KeyAscii As Integer)
    'mengubah huruf kecil jadi huruf besar
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtKodePelanggan_KeyPress(KeyAscii As Integer)
    'mengubah huruf kecil jadi huruf besar
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub txtNamaPelanggan_KeyPress(KeyAscii As Integer)
    'mengubah huruf kecil jadi huruf besar
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtTeleponpelanggan_KeyPress(KeyAscii As Integer)
 'hanya boleh diisi angka atau backspace
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii <= Asc("-") Or KeyAscii = vbKeyBack) Then
       Beep
       KeyAscii = 0
    End If
End Sub
