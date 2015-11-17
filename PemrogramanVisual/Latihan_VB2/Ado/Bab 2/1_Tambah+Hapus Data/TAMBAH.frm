VERSION 5.00
Begin VB.Form Form28 
   Caption         =   "Menambah, membatalkan, dan menghapus Data"
   ClientHeight    =   2730
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7095
   LinkTopic       =   "Form1"
   ScaleHeight     =   2730
   ScaleWidth      =   7095
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTutup 
      Caption         =   "Tutup"
      Height          =   375
      Left            =   5880
      TabIndex        =   16
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "&Simpan"
      Enabled         =   0   'False
      Height          =   372
      Left            =   5880
      TabIndex        =   15
      Top             =   960
      Width           =   972
   End
   Begin VB.CommandButton cmdTambah 
      Caption         =   "&Tambah"
      CausesValidation=   0   'False
      Height          =   372
      Left            =   5880
      TabIndex        =   14
      Top             =   480
      Width           =   972
   End
   Begin VB.CommandButton cmdHapus 
      Caption         =   "&Hapus"
      Height          =   372
      Left            =   5880
      TabIndex        =   13
      Top             =   1440
      Width           =   972
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   ">"
      Default         =   -1  'True
      Height          =   372
      Left            =   3000
      TabIndex        =   12
      Top             =   1800
      Width           =   972
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "<"
      Height          =   372
      Left            =   1920
      TabIndex        =   11
      Top             =   1800
      Width           =   972
   End
   Begin VB.CommandButton cmdFirst 
      Caption         =   "I<"
      Height          =   372
      Left            =   840
      TabIndex        =   10
      Top             =   1800
      Width           =   972
   End
   Begin VB.CommandButton cmdLast 
      Caption         =   ">I"
      Height          =   372
      Left            =   4080
      TabIndex        =   9
      Top             =   1800
      Width           =   972
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   480
      TabIndex        =   8
      Top             =   1560
      Width           =   5055
   End
   Begin VB.TextBox txtNoTelepon 
      DataField       =   "No Telepon"
      DataMember      =   "Pemasok"
      DataSource      =   "DE"
      Height          =   285
      Left            =   2280
      TabIndex        =   7
      Top             =   1185
      Width           =   1980
   End
   Begin VB.TextBox txtAlamatPemasok 
      DataField       =   "Alamat Pemasok"
      DataMember      =   "Pemasok"
      DataSource      =   "DE"
      Height          =   285
      Left            =   2280
      TabIndex        =   5
      Top             =   810
      Width           =   3375
   End
   Begin VB.TextBox txtNamaPemasok 
      DataField       =   "Nama Pemasok"
      DataMember      =   "Pemasok"
      DataSource      =   "DE"
      Height          =   285
      Left            =   2280
      TabIndex        =   3
      Top             =   420
      Width           =   3375
   End
   Begin VB.TextBox txtKodePemasok 
      DataField       =   "Kode Pemasok"
      DataMember      =   "Pemasok"
      DataSource      =   "DE"
      Height          =   285
      Left            =   2280
      TabIndex        =   1
      Top             =   45
      Width           =   990
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "No Telepon:"
      Height          =   255
      Index           =   3
      Left            =   435
      TabIndex        =   6
      Top             =   1230
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Alamat Pemasok:"
      Height          =   255
      Index           =   2
      Left            =   435
      TabIndex        =   4
      Top             =   855
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Nama Pemasok:"
      Height          =   255
      Index           =   1
      Left            =   435
      TabIndex        =   2
      Top             =   465
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Kode Pemasok:"
      Height          =   255
      Index           =   0
      Left            =   435
      TabIndex        =   0
      Top             =   90
      Width           =   1815
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
'Purpose        : Menambah dan menghapus data menggunakan objek visual Data Environment
'Folder         : ADO/BAB2/1_Tambah+HapusData

Private Sub cmdTambah_Click()
 'Add a new record
    
    On Error GoTo HandleError        'matikan penyaringan kesalahan / error utk penanganan kesalahan
    If cmdTambah.Caption = "&Tambah" Then
        DE.rsPemasok.AddNew          'kosongkan field untuk record data baru
        NonaktifkanKontrol           'tombol-tombol pengerak Record dimatikan
        cmdTambah.Caption = "&Batal" 'mengganti tulisan tambah menjadi batal
        cmdSimpan.Enabled = True     'aktifkan tombol simpan
        txtKodePemasok.Locked = False 'buka text kode Pemasok
        'SetUpAdd
    Else
        DE.rsPemasok.CancelUpdate     'Batalkan proses penambahan data
        txtKodePemasok.Locked = True  'kunci text kode pemasok
        AktifkanTombol                'semua tombol penggerak record diaktifkan
        cmdTambah.Caption = "&Tambah"    'ubah tulisan batal menjadi Tambah
        cmdSimpan.Enabled = False       'Disable the Save button
        DE.rsPemasok.MoveLast         'Pindah ke record data terakhir
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
    With DE.rsPemasok
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
    
    DE.rsPemasok.Update             'simpan record
    txtKodePemasok.Locked = True    'kunci text kode pemasok
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
    DE.rsPemasok.MoveFirst
End Sub

Private Sub cmdLast_Click()
    'Move to last record
    
    On Error Resume Next
    DE.rsPemasok.MoveLast
End Sub

Private Sub cmdNext_Click()
    'Move to next record
    
    On Error Resume Next
    With DE.rsPemasok
        .MoveNext
        If .EOF Then
            .MoveLast
        End If
    End With
End Sub

Private Sub cmdPrevious_Click()
    'Move to previous record
    
    On Error Resume Next
    With DE.rsPemasok
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

