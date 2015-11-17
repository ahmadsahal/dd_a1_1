VERSION 5.00
Begin VB.Form frmPelanggan 
   Caption         =   "Data Pelanggan"
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
'Descripsi  :Menampilakan, menambah, dan menghapus data Pelanggan pada database AVB8,
'            menggunakan Perintah SQL

Option Explicit
Dim conAVB As ADODB.Connection
Dim rsPelanggan As ADODB.Recordset

Private Sub Form_Load()
    'Create the connection
    Dim strsql As String
    Set conAVB = New ADODB.Connection
    
    conAVB.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "Persist Security Info=False;Data Source=" & App.Path & _
        "\AVB8.mdb;Mode = readwrite"
    conAVB.Open
    
    'Buat recordset
    Set rsPelanggan = New ADODB.Recordset
    strsql = "Select * from Pelanggan"
    
    rsPelanggan.Open strsql, conAVB, adOpenDynamic, adLockOptimistic, adCmdText
    If rsPelanggan.RecordCount > 0 Then ' jika ada record data (Data tidak kosong)
        TampilkanData
    End If
End Sub
Private Sub TampilkanData()
    'Transfer dari database
        With rsPelanggan
            txtKodePelanggan.Text = ![Kode Pelanggan]
            txtNamaPelanggan.Text = ![Nama Pelanggan]
            txtAlamatpelanggan.Text = ![ALamat Pelanggan]
            txtTeleponpelanggan.Text = ![Telepon Pelanggan]
         End With
End Sub

Private Sub cmdHapus_Click()
'Hapus Record yang aktif
    Dim strsql As String
    strsql = "Delete From Pelanggan " _
           & "Where [Kode Pelanggan] = '" & txtKodePelanggan.Text & "'"
    conAVB.Execute strsql, , adCmdText
    With rsPelanggan
        BersihkanFieldText
        .MoveNext                   'Move to the following record
        If .EOF Then                'If last record deleted
            .MovePrevious
            If .BOF Then            'If BOF and EOF true, no records remain
                MsgBox "Data kosong.", vbInformation, "Tidak ada record data"
                nonAktifkanTombol
            End If
        End If
    End With
End Sub

Private Sub cmdFirst_Click()
'Pindah ke record pertama
    On Error Resume Next
    rsPelanggan.MoveFirst
    TampilkanData
End Sub

Private Sub cmdLast_Click()
'Pindah Ke record terakhir
    On Error Resume Next
    rsPelanggan.MoveLast
    TampilkanData
End Sub

Private Sub cmdNext_Click()

    On Error Resume Next
    With rsPelanggan
        .MoveNext
        If .EOF Then
            .MoveFirst
        End If
        TampilkanData
    End With

End Sub

Private Sub cmdPrevious_Click()

    On Error Resume Next
    With rsPelanggan
        .MovePrevious
        If .BOF Then
            .MoveLast
        End If
        TampilkanData
    End With

End Sub

Private Sub cmdSimpan_Click()
 'Simpan record yang sedang aktif
    On Error GoTo HandleErrors
    
        Dim strsql As String
        strsql = "Insert Into Pelanggan " _
               & "([Kode Pelanggan],[Nama Pelanggan],[Alamat Pelanggan],[Telepon pelanggan]) " _
               & "VALUES ('" & txtKodePelanggan & "', '" & txtNamaPelanggan & "','" & txtAlamatpelanggan & "','" & txtTeleponpelanggan & "')"
        conAVB.Execute strsql, , adCmdText
        
        rsPelanggan.Requery   'Tambah record baru pada  recordset
        txtKodePelanggan.Locked = True       'Hapus semuanya
        AktifkanTombol
        cmdSimpan.Enabled = False
        cmdTambah.Caption = "&Tambah"
    
cmdSimpan_Click_Exit:
Exit Sub

HandleErrors:
    Dim strMessage As String
    Dim errDBError As ADODB.Error
    
    For Each errDBError In conAVB.Errors
        strMessage = strMessage & Err.Description & vbCrLf
    Next
    MsgBox strMessage, vbExclamation, "Kesalahan Provider"
    SetUpTambahRecord  'Simpan data pemakai
    On Error GoTo 0 'matikan perangkan kesalahan
End Sub
Private Sub SetUpTambahRecord()
    'Set up a new Tambah to give user another try
    Dim StrKodePelanggan As String, strNamaPelanggan As String
    Dim strAlamatPelanggan As String, StrTeleponpelanggan As String
    
    On Error Resume Next
    'Save contents of form controls
    StrKodePelanggan = txtKodePelanggan.Text
    strNamaPelanggan = txtNamaPelanggan.Text
    strAlamatPelanggan = txtAlamatpelanggan.Text
    StrTeleponpelanggan = txtTeleponpelanggan.Text
    'Mulai penambahan sebuah data baru
    rsPelanggan.AddNew
    'Pangil kembali data yang telah disimpan ke form
    With txtKodePelanggan
        .Text = StrKodePelanggan
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With
    txtNamaPelanggan.Text = strNamaPelanggan
    txtAlamatpelanggan.Text = strAlamatPelanggan
    txtTeleponpelanggan.Text = StrTeleponpelanggan
End Sub

Private Sub cmdTambah_Click()
    'Tambah data record baru
    If cmdTambah.Caption = "&Tambah" Then
        BersihkanFieldText
        nonAktifkanTombol                'Disable navigation
        cmdSimpan.Enabled = True
        cmdTambah.Caption = "&Batal"
    Else
        rsPelanggan.CancelUpdate      'Cancel the Tambah
        txtKodePelanggan.Locked = True
        AktifkanTombol           'Enable navigation
        cmdSimpan.Enabled = False     'Disable the Save button
        cmdTambah.Caption = "&Tambah"     'Reset the Tambah button
        rsPelanggan.MoveLast
        TampilkanData
    End If
End Sub
Private Sub nonAktifkanTombol()
    'Tombol-tombol navigasi dinonaktifkan
    cmdNext.Enabled = False
    cmdPrevious.Enabled = False
    cmdFirst.Enabled = False
    cmdLast.Enabled = False
    cmdHapus.Enabled = False
End Sub
Private Sub AktifkanTombol()
    'tombol-tombol navigasi diaktifkan
    cmdNext.Enabled = True
    cmdPrevious.Enabled = True
    cmdFirst.Enabled = True
    cmdLast.Enabled = True
    cmdHapus.Enabled = True
End Sub
Private Sub BersihkanFieldText()
    'Bersihkan semua text boxes untuk sebuah proses penambahan data
    With txtKodePelanggan
        .Text = ""
        .Locked = False
        .SetFocus
    End With
    txtNamaPelanggan.Text = ""
    txtAlamatpelanggan.Text = ""
    txtTeleponpelanggan.Text = ""
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub


Private Sub txtTeleponpelanggan_KeyPress(KeyAscii As Integer)
    'hanya boleh diisi angka atau backspace
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii <= Asc("-") Or KeyAscii = vbKeyBack) Then
       Beep
       KeyAscii = 0
    End If
End Sub

Private Sub txtKodePelanggan_KeyPress(KeyAscii As Integer)
    'mengubah huruf kecil jadi huruf besar
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub txtNamaPelanggan_KeyPress(KeyAscii As Integer)
    'mengubah huruf kecil jadi huruf besar
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub





