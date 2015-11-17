VERSION 5.00
Begin VB.Form frmBarang 
   Caption         =   "Data Barang"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4995
   LinkTopic       =   "Form1"
   ScaleHeight     =   2505
   ScaleWidth      =   4995
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtHargaBarang 
      Height          =   285
      Left            =   1560
      TabIndex        =   14
      Top             =   840
      Width           =   1320
   End
   Begin VB.TextBox txtNamaBarang 
      Height          =   285
      Left            =   1560
      TabIndex        =   12
      Top             =   480
      Width           =   3375
   End
   Begin VB.TextBox txtKodeBarang 
      Height          =   285
      Left            =   1560
      TabIndex        =   10
      Top             =   120
      Width           =   990
   End
   Begin VB.Frame fraNavigation 
      Height          =   735
      Left            =   720
      TabIndex        =   4
      Top             =   1200
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
      Left            =   2520
      TabIndex        =   3
      Top             =   2040
      Width           =   972
   End
   Begin VB.CommandButton cmdTambah 
      Caption         =   "&Tambah"
      CausesValidation=   0   'False
      Height          =   372
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   972
   End
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "&Simpan"
      Enabled         =   0   'False
      Height          =   372
      Left            =   1320
      TabIndex        =   1
      Top             =   2040
      Width           =   972
   End
   Begin VB.CommandButton cmdTutup 
      Caption         =   "Tutup"
      Height          =   375
      Left            =   3720
      TabIndex        =   0
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Harga Barang:"
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   13
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Nama Barang:"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   11
      Top             =   495
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Kode Barang:"
      Height          =   255
      Index           =   0
      Left            =   360
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
'Descripsi  :Menampilakan, menambah, dan menghapus data barang pada database AVB8,
'            menggunakan Perintah SQL

Option Explicit
Dim conAVB As ADODB.Connection
Dim rsBarang As ADODB.Recordset

Private Sub cmdHapus_Click()
'Hapus Record yang aktif
    Dim strsql As String
    strsql = "Delete From Barang " _
           & "Where [Kode Barang] = '" & txtKodeBarang.Text & "'"
    conAVB.Execute strsql, , adCmdText
    With rsBarang
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

Private Sub Form_Load()
    'Create the connection
    Dim strsql As String
    Set conAVB = New ADODB.Connection
    
    conAVB.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "Persist Security Info=False;Data Source=" & App.Path & _
        "\AVB8.mdb;Mode = readwrite"
    conAVB.Open
    
    'Buat recordset
    Set rsBarang = New ADODB.Recordset
    strsql = "Select * from Barang"
    
    rsBarang.Open strsql, conAVB, adOpenDynamic, adLockOptimistic, adCmdText
    If rsBarang.RecordCount > 0 Then ' jika ada record data (Data tidak kosong)
        TampilkanData
    End If
End Sub
Private Sub TampilkanData()
    'Transfer dari database
        With rsBarang
            txtKodeBarang.Text = ![Kode Barang]
            txtNamaBarang.Text = ![Nama barang]
            txtHargaBarang.Text = ![Harga Barang]
         End With
End Sub

Private Sub cmdFirst_Click()
'Pindah ke record pertama
    On Error Resume Next
    rsBarang.MoveFirst
    TampilkanData
End Sub

Private Sub cmdLast_Click()
'Pindah Ke record terakhir
    On Error Resume Next
    rsBarang.MoveLast
    TampilkanData
End Sub

Private Sub cmdNext_Click()

    On Error Resume Next
    With rsBarang
        .MoveNext
        If .EOF Then
            .MoveFirst
        End If
        TampilkanData
    End With

End Sub

Private Sub cmdPrevious_Click()

    On Error Resume Next
    With rsBarang
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
        strsql = "Insert Into Barang " _
               & "([Kode Barang],[Nama Barang],[Harga Barang]) " _
               & "VALUES ('" & txtKodeBarang & "', '" & txtNamaBarang & "','" & txtHargaBarang & "')"
        conAVB.Execute strsql, , adCmdText
        
        rsBarang.Requery   'Tambah the new record to the recordset
        txtKodeBarang.Locked = True       'Reset all
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
    Dim StrKodeBarang As String
    Dim strNamaBarang As String
    Dim strHargaBarang As String
    
    On Error Resume Next
    'Save contents of form controls
    StrKodeBarang = txtKodeBarang.Text
    strNamaBarang = txtNamaBarang.Text
    strHargaBarang = txtHargaBarang.Text
    'Mulai penambahan sebuah data baru
    rsBarang.AddNew
    'Pangil kembali data yang telah disimpan ke form
    With txtKodeBarang
        .Text = StrKodeBarang
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With
    txtNamaBarang.Text = strNamaBarang
    txtHargaBarang.Text = strHargaBarang
End Sub

Private Sub cmdTambah_Click()
    'Tambah data record baru
    If cmdTambah.Caption = "&Tambah" Then
        BersihkanFieldText
        nonAktifkanTombol                'Disable navigation
        cmdSimpan.Enabled = True
        cmdTambah.Caption = "&Batal"
    Else
        rsBarang.CancelUpdate      'Cancel the Tambah
        txtKodeBarang.Locked = True
        AktifkanTombol           'Enable navigation
        cmdSimpan.Enabled = False     'Disable the Save button
        cmdTambah.Caption = "&Tambah"     'Reset the Tambah button
        rsBarang.MoveLast
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
    With txtKodeBarang
        .Text = ""
        .Locked = False
        .SetFocus
    End With
    txtNamaBarang.Text = ""
    txtHargaBarang.Text = ""
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub txtHargaBarang_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack) Then
       Beep
       KeyAscii = 0
    End If
End Sub

Private Sub txtKodeBarang_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtNamaBarang_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
