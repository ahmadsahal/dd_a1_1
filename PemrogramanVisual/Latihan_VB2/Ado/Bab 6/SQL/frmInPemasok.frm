VERSION 5.00
Begin VB.Form frmPemasok 
   Caption         =   "Data Pemasok"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5385
   LinkTopic       =   "Form1"
   ScaleHeight     =   3075
   ScaleWidth      =   5385
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtNoTelepon 
      DataField       =   "No Telepon"
      DataMember      =   "Pemasok"
      DataSource      =   "DE"
      Height          =   285
      Left            =   1530
      TabIndex        =   16
      Top             =   1380
      Width           =   1980
   End
   Begin VB.TextBox txtAlamatPemasok 
      DataField       =   "Alamat Pemasok"
      DataMember      =   "Pemasok"
      DataSource      =   "DE"
      Height          =   285
      Left            =   1530
      TabIndex        =   14
      Top             =   1005
      Width           =   3375
   End
   Begin VB.TextBox txtNamaPemasok 
      DataField       =   "Nama Pemasok"
      DataMember      =   "Pemasok"
      DataSource      =   "DE"
      Height          =   285
      Left            =   1530
      TabIndex        =   12
      Top             =   615
      Width           =   3375
   End
   Begin VB.TextBox txtKodePemasok 
      DataField       =   "Kode Pemasok"
      DataMember      =   "Pemasok"
      DataSource      =   "DE"
      Height          =   285
      Left            =   1530
      TabIndex        =   10
      Top             =   240
      Width           =   990
   End
   Begin VB.Frame fraNavigation 
      Height          =   735
      Left            =   960
      TabIndex        =   4
      Top             =   1680
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
      Left            =   2760
      TabIndex        =   3
      Top             =   2520
      Width           =   972
   End
   Begin VB.CommandButton cmdTambah 
      Caption         =   "&Tambah"
      CausesValidation=   0   'False
      Height          =   372
      Left            =   360
      TabIndex        =   2
      Top             =   2520
      Width           =   972
   End
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "&Simpan"
      Enabled         =   0   'False
      Height          =   372
      Left            =   1560
      TabIndex        =   1
      Top             =   2520
      Width           =   972
   End
   Begin VB.CommandButton cmdTutup 
      Caption         =   "Tutup"
      Height          =   375
      Left            =   3960
      TabIndex        =   0
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "No Telepon:"
      Height          =   255
      Index           =   3
      Left            =   165
      TabIndex        =   15
      Top             =   1425
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Alamat Pemasok:"
      Height          =   255
      Index           =   2
      Left            =   165
      TabIndex        =   13
      Top             =   1050
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Nama Pemasok:"
      Height          =   255
      Index           =   1
      Left            =   165
      TabIndex        =   11
      Top             =   660
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Kode Pemasok:"
      Height          =   255
      Index           =   0
      Left            =   165
      TabIndex        =   9
      Top             =   285
      Width           =   1815
   End
End
Attribute VB_Name = "frmPemasok"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Proyek     :Menginput data Pemasok
'tanggal    :22 November 2001
'Programmer :Kok Yung
'Descripsi  :Menampilakan, menambah, dan menghapus data Pemasok pada database AVB8,
'            menggunakan Perintah SQL

Option Explicit
Dim conAVB As ADODB.Connection
Dim rsPemasok As ADODB.Recordset

Private Sub Form_Load()
    'Create the connection
    Dim strsql As String
    Set conAVB = New ADODB.Connection
    
    conAVB.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "Persist Security Info=False;Data Source=" & App.Path & _
        "\AVB8.mdb;Mode = readwrite"
    conAVB.Open
    
    'Buat recordset
    Set rsPemasok = New ADODB.Recordset
    strsql = "Select * from Pemasok"
    
    rsPemasok.Open strsql, conAVB, adOpenDynamic, adLockOptimistic, adCmdText
    If rsPemasok.RecordCount > 0 Then ' jika ada record data (Data tidak kosong)
        TampilkanData
    End If
End Sub
Private Sub TampilkanData()
    'Transfer dari database
        With rsPemasok
            txtKodePemasok.Text = ![Kode Pemasok]
            txtNamaPemasok.Text = ![Nama Pemasok]
            txtAlamatPemasok.Text = ![ALamat Pemasok]
            txtNoTelepon.Text = ![no telepon]
         End With
End Sub

Private Sub cmdHapus_Click()
'Hapus Record yang aktif
    Dim strsql As String
    strsql = "Delete From Pemasok " _
           & "Where [Kode Pemasok] = '" & txtKodePemasok.Text & "'"
    conAVB.Execute strsql, , adCmdText
    With rsPemasok
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
    rsPemasok.MoveFirst
    TampilkanData
End Sub

Private Sub cmdLast_Click()
'Pindah Ke record terakhir
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
 'Simpan record yang sedang aktif
    On Error GoTo HandleErrors
    
        Dim strsql As String
        strsql = "Insert Into Pemasok " _
               & "([Kode Pemasok],[Nama Pemasok],[Alamat Pemasok],[No Telepon]) " _
               & "VALUES ('" & txtKodePemasok & "', '" & txtNamaPemasok & "','" & txtAlamatPemasok & "','" & txtNoTelepon & "')"
        conAVB.Execute strsql, , adCmdText
        
        rsPemasok.Requery   'Tambah record baru pada  recordset
        txtKodePemasok.Locked = True       'Hapus semuanya
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
    Dim StrKodePemasok As String, strNamaPemasok As String
    Dim strAlamatPemasok As String, strNoTelepon As String
    
    On Error Resume Next
    'Save contents of form controls
    StrKodePemasok = txtKodePemasok.Text
    strNamaPemasok = txtNamaPemasok.Text
    strAlamatPemasok = txtAlamatPemasok.Text
    strNoTelepon = txtNoTelepon.Text
    'Mulai penambahan sebuah data baru
    rsPemasok.AddNew
    'Pangil kembali data yang telah disimpan ke form
    With txtKodePemasok
        .Text = StrKodePemasok
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With
    txtNamaPemasok.Text = strNamaPemasok
    txtAlamatPemasok.Text = strAlamatPemasok
    txtNoTelepon.Text = strNoTelepon
End Sub

Private Sub cmdTambah_Click()
    'Tambah data record baru
    If cmdTambah.Caption = "&Tambah" Then
        BersihkanFieldText
        nonAktifkanTombol                'Disable navigation
        cmdSimpan.Enabled = True
        cmdTambah.Caption = "&Batal"
    Else
        rsPemasok.CancelUpdate      'Cancel the Tambah
        txtKodePemasok.Locked = True
        AktifkanTombol           'Enable navigation
        cmdSimpan.Enabled = False     'Disable the Save button
        cmdTambah.Caption = "&Tambah"     'Reset the Tambah button
        rsPemasok.MoveLast
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
    With txtKodePemasok
        .Text = ""
        .Locked = False
        .SetFocus
    End With
    txtNamaPemasok.Text = ""
    txtAlamatPemasok.Text = ""
    txtNoTelepon.Text = ""
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub


Private Sub txtNoTelepon_KeyPress(KeyAscii As Integer)
    'hanya boleh diisi angka atau backspace
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii <= Asc("-") Or KeyAscii = vbKeyBack) Then
       Beep
       KeyAscii = 0
    End If
End Sub

Private Sub txtKodePemasok_KeyPress(KeyAscii As Integer)
    'mengubah huruf kecil jadi huruf besar
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub txtNamaPemasok_KeyPress(KeyAscii As Integer)
    'mengubah huruf kecil jadi huruf besar
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub txtAlamatPemasok_KeyPress(KeyAscii As Integer)
    'mengubah huruf kecil jadi huruf besar
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub



