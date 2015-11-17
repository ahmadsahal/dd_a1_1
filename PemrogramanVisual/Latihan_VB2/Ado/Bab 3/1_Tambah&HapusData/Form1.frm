VERSION 5.00
Begin VB.Form Form28 
   Caption         =   "Menambah, menghapus, dan Menambah Data"
   ClientHeight    =   2595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5970
   LinkTopic       =   "Form1"
   ScaleHeight     =   2595
   ScaleWidth      =   5970
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   12
      Top             =   1680
      Width           =   4695
      Begin VB.CommandButton cmdLast 
         Caption         =   ">I"
         Height          =   372
         Left            =   3480
         TabIndex        =   16
         Top             =   240
         Width           =   972
      End
      Begin VB.CommandButton cmdFirst 
         Caption         =   "I<"
         Height          =   372
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Width           =   972
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "<"
         Height          =   372
         Left            =   1320
         TabIndex        =   14
         Top             =   240
         Width           =   972
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   ">"
         Default         =   -1  'True
         Height          =   372
         Left            =   2400
         TabIndex        =   13
         Top             =   240
         Width           =   972
      End
   End
   Begin VB.CommandButton cmdHapus 
      Caption         =   "&Hapus"
      Height          =   372
      Left            =   4920
      TabIndex        =   11
      Top             =   1080
      Width           =   972
   End
   Begin VB.CommandButton cmdTambah 
      Caption         =   "&Tambah"
      CausesValidation=   0   'False
      Height          =   372
      Left            =   4920
      TabIndex        =   10
      Top             =   120
      Width           =   972
   End
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "&Simpan"
      Enabled         =   0   'False
      Height          =   372
      Left            =   4920
      TabIndex        =   9
      Top             =   600
      Width           =   972
   End
   Begin VB.CommandButton cmdTutup 
      Caption         =   "Tutup"
      Height          =   375
      Left            =   4920
      TabIndex        =   8
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox txtNoTelepon 
      Height          =   285
      Left            =   1425
      TabIndex        =   7
      Top             =   1305
      Width           =   1980
   End
   Begin VB.TextBox txtAlamatPemasok 
      Height          =   285
      Left            =   1425
      TabIndex        =   5
      Top             =   930
      Width           =   3375
   End
   Begin VB.TextBox txtNamaPemasok 
      Height          =   285
      Left            =   1425
      TabIndex        =   3
      Top             =   540
      Width           =   3375
   End
   Begin VB.TextBox txtKodePemasok 
      Height          =   285
      Left            =   1425
      TabIndex        =   1
      Top             =   165
      Width           =   990
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "No Telepon:"
      Height          =   255
      Index           =   3
      Left            =   60
      TabIndex        =   6
      Top             =   1350
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Alamat Pemasok:"
      Height          =   255
      Index           =   2
      Left            =   60
      TabIndex        =   4
      Top             =   975
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Nama Pemasok:"
      Height          =   255
      Index           =   1
      Left            =   60
      TabIndex        =   2
      Top             =   585
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Kode Pemasok:"
      Height          =   255
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   210
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
'Purpose        : menambah dan menghapus data tanpa menggunakan
'                 objek visual ( dengan menggunakan perintah SQL)
'Folder         : ADO/BAB2/1_Menambah+menghapusData


Option Explicit
Dim conAVB As ADODB.Connection
Dim rsPemasok As ADODB.Recordset


Private Sub Form_Load()
    'membuat connection
    Dim strSQL As String
    Set conAVB = New ADODB.Connection
    
    conAVB.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "Persist Security Info=False;Data Source=" & App.Path & _
        "\AVB.mdb;Mode = readwrite"
    conAVB.Open
    
    'membuat recordset
    Set rsPemasok = New ADODB.Recordset
    strSQL = "Select * from Pemasok"
    
    rsPemasok.Open strSQL, conAVB, adOpenDynamic, adLockOptimistic, adCmdText
    If rsPemasok.RecordCount > 0 Then
        TampilkanData
    End If
End Sub
Private Sub cmdTambah_Click()
    'menambah sebuah record baru
    If cmdTambah.Caption = "&Tambah" Then
        BersihkanText
        NonAktifkanTombol       'menonaktifkan tombol navigasi
        cmdSimpan.Enabled = True
        cmdTambah.Caption = "&Cancel"
    Else
        rsPemasok.CancelUpdate      'membatalkan penambahan data
        txtKodePemasok.Locked = True
        AktifkanTombol       'mengaktifkan tombol navigasi
        cmdSimpan.Enabled = False     'mengunci tombol simpan
        cmdTambah.Caption = "&Tambah"     'mengembalikan ke tombol tambah
        rsPemasok.MoveLast
        
    End If
   
End Sub

Private Sub cmdHapus_Click()
    'Menghapus record yang sedang aktif / bekerja
    Dim strSQL As String
    
    strSQL = "Delete From Pemasok " _
           & "Where [Kode Pemasok] = '" & txtKodePemasok.Text & "'"
    conAVB.Execute strSQL, , adCmdText

    
    With rsPemasok
        .MoveNext           'pindah ke record sebelumnya
        If .EOF Then            'jika itu record terakhir hapus
        .MovePrevious
            If .BOF Then        'jika BOF dan EOF benar, record kosong
                MsgBox "Tidak ada data.", vbInformation, "tidak ada record"
                NonAktifkanTombol
            End If
        End If
        BersihkanText
    End With
End Sub

Private Sub cmdFirst_Click()
    On Error Resume Next
    rsPemasok.MoveFirst
    TampilkanData
End Sub

Private Sub cmdLast_Click()
    On Error Resume Next
    rsPemasok.MoveLast
    TampilkanData
End Sub

Private Sub cmdNext_Click()
    On Error Resume Next
    With rsPemasok
        .MoveNext
        If .EOF Then
            .MoveFirst
        End If
        TampilkanData
    End With
End Sub

Private Sub cmdPrevious_Click()
    On Error Resume Next
    With rsPemasok
        .MovePrevious
        If .BOF Then
            .MoveLast
        End If
        TampilkanData
    End With
End Sub

Private Sub cmdSimpan_Click()
    
    On Error GoTo HandleErrors
    Dim strSQL As String
    
    strSQL = "Insert Into Pemasok " _
                 & "([Kode Pemasok],[Nama Pemasok], [Alamat Pemasok], [No Telepon]) " _
                 & "VALUES ('" & txtKodePemasok & "', '" & txtNamaPemasok & "', '" _
                 & txtAlamatPemasok & " ', '" & txtNoTelepon & "')"
    
    conAVB.Execute strSQL, , adCmdText
    rsPemasok.Requery  'Tambah record baru ke recordset
    txtKodePemasok.Locked = True       'hapus semuanya
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
    MsgBox strMessage, vbExclamation, "Salah Provider"
    SetUpTambahRecord  'simpan data pemakai
    On Error GoTo 0 'matikan perangkap kesalahan
End Sub
Private Sub cmdTutup_Click()
    Unload Me
End Sub
Private Sub NonAktifkanTombol()
    
    'non aktifkan tombol navigasi
    cmdNext.Enabled = False
    cmdPrevious.Enabled = False
    cmdFirst.Enabled = False
    cmdLast.Enabled = False
    cmdHapus.Enabled = False
End Sub

Private Sub AktifkanTombol()
    
    'aktifkan tombol navigasi
    cmdNext.Enabled = True
    cmdPrevious.Enabled = True
    cmdFirst.Enabled = True
    cmdLast.Enabled = True
    cmdHapus.Enabled = True
End Sub
Private Sub SetUpTambahRecord()
    'bentuk sebuah record untuk dicoba kembali oleh pemakai(user)
    Dim strKodePemasok As String, strNamaPemasok As String
    Dim strAlamatPemasok As String
    Dim strNoTelepon As String
    
    On Error Resume Next
    
    'simpan isi ke dalam form kontrol
    strKodePemasok = txtKodePemasok.Text
    strNamaPemasok = txtNamaPemasok.Text
    strAlamatPemasok = txtAlamatPemasok.Text
    strNoTelepon = txtNoTelepon.Text
    
    'mulai dengan menambah data baru
    rsPemasok.AddNew
    
    'panggil data yang disiman ke dalam form
    With txtKodePemasok
        .Text = strKodePemasok
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With
    txtNamaPemasok.Text = strNamaPemasok
    txtAlamatPemasok.Text = strAlamatPemasok
    txtNoTelepon.Text = strNoTelepon
End Sub

Private Sub BersihkanText()
    
    ' bersihkan objek tex box untuk penambahan data
    With txtKodePemasok
        .Text = ""
        .Locked = False
        .SetFocus
    End With
    txtNamaPemasok.Text = ""
    txtAlamatPemasok.Text = ""
    txtNoTelepon.Text = ""
End Sub

Private Sub TampilkanData()
    'Transfer dari database
            
        With rsPemasok
            txtKodePemasok.Text = ![Kode Pemasok]
            txtNamaPemasok.Text = ![Nama Pemasok]
            txtAlamatPemasok.Text = ![Alamat Pemasok]
            txtNoTelepon.Text = ![No Telepon]
        End With
    
End Sub


