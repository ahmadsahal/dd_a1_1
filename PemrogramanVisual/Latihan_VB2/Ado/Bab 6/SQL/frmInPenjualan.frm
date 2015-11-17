VERSION 5.00
Begin VB.Form frmInJual 
   Caption         =   "Data Penjualan"
   ClientHeight    =   4800
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4755
   LinkTopic       =   "Form1"
   ScaleHeight     =   4800
   ScaleWidth      =   4755
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraNavigation 
      Height          =   735
      Left            =   600
      TabIndex        =   22
      Top             =   3480
      Width           =   3495
      Begin VB.CommandButton cmdLast 
         Caption         =   ">I"
         Height          =   375
         Left            =   2640
         TabIndex        =   26
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   ">"
         Default         =   -1  'True
         Height          =   375
         Left            =   1800
         TabIndex        =   25
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "<"
         Height          =   375
         Left            =   960
         TabIndex        =   24
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdFirst 
         Caption         =   "I<"
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdHapus 
      Caption         =   "&Hapus"
      Height          =   372
      Left            =   2400
      TabIndex        =   21
      Top             =   4320
      Width           =   855
   End
   Begin VB.CommandButton cmdTambah 
      Caption         =   "&Tambah"
      CausesValidation=   0   'False
      Height          =   372
      Left            =   480
      TabIndex        =   20
      Top             =   4320
      Width           =   855
   End
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "&Simpan"
      Enabled         =   0   'False
      Height          =   372
      Left            =   1440
      TabIndex        =   19
      Top             =   4320
      Width           =   855
   End
   Begin VB.CommandButton cmdTutup 
      Caption         =   "Tutup"
      Height          =   375
      Left            =   3360
      TabIndex        =   18
      Top             =   4320
      Width           =   855
   End
   Begin VB.TextBox txtNamaBrg 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1560
      TabIndex        =   12
      Top             =   1920
      Width           =   2985
   End
   Begin VB.TextBox txtHargaSatuan 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   300
      Left            =   1560
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   2280
      Width           =   1164
   End
   Begin VB.TextBox txtBanyaknyaBarang 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   1560
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   2640
      Width           =   1164
   End
   Begin VB.TextBox txtJumlah 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
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
      Left            =   1560
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   3000
      Width           =   2028
   End
   Begin VB.ComboBox dbcbarang 
      Height          =   315
      Left            =   1560
      TabIndex        =   8
      Top             =   1560
      Width           =   1695
   End
   Begin VB.ComboBox dbcPelanggan 
      Height          =   315
      Left            =   1560
      TabIndex        =   7
      Top             =   840
      Width           =   1695
   End
   Begin VB.TextBox txtTanggalBon 
      Height          =   300
      Left            =   1560
      TabIndex        =   2
      Top             =   480
      Width           =   1065
   End
   Begin VB.TextBox txtNamaPelanggan 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1560
      TabIndex        =   1
      Top             =   1200
      Width           =   2028
   End
   Begin VB.TextBox txtNoBon 
      Height          =   300
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Kode Barang"
      Height          =   210
      Left            =   240
      TabIndex        =   17
      Top             =   1560
      Width           =   1065
   End
   Begin VB.Label Label4 
      Caption         =   "Nama Barang"
      Height          =   210
      Left            =   240
      TabIndex        =   16
      Top             =   1920
      Width           =   1590
   End
   Begin VB.Label Label5 
      Caption         =   "Harga Satuan"
      Height          =   210
      Left            =   240
      TabIndex        =   15
      Top             =   2280
      Width           =   1065
   End
   Begin VB.Label Label6 
      Caption         =   "Banyak"
      Height          =   210
      Left            =   240
      TabIndex        =   14
      Top             =   2640
      Width           =   1065
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
      Left            =   240
      TabIndex        =   13
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Kode Pelanggan"
      Height          =   210
      Left            =   240
      TabIndex        =   6
      Top             =   840
      Width           =   1890
   End
   Begin VB.Label Label2 
      Caption         =   "Nama"
      Height          =   210
      Left            =   240
      TabIndex        =   5
      Top             =   1200
      Width           =   870
   End
   Begin VB.Label Label7 
      Caption         =   "Nomor Bon"
      Height          =   210
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   1065
   End
   Begin VB.Label Label8 
      Caption         =   "Tanggal Bon"
      Height          =   210
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Width           =   870
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
'Descripsi  :Menampilakan, menambah, dan menghapus penjualan barang pada database AVB8,
'            menggunakan perintah SQL
                                   
Option Explicit
Dim conAVB As ADODB.Connection
Dim rsPenjualan As ADODB.Recordset
Dim rsBarang As ADODB.Recordset
Dim rsPelanggan As ADODB.Recordset
Dim nilai As Integer

Private Sub dbcbarang_Change()
'dapatkan record yang dipilih
    Dim strsqlBarang  As String
    
    strsqlBarang = "Select * from Barang Where [Kode Barang] = '" & dbcbarang.Text & "'"
    Set rsBarang = conAVB.Execute(strsqlBarang, , adCmdText)
    'kirim dari database
    With rsBarang
        If .BOF And .EOF Then
            Exit Sub
        Else
            txtNamaBrg.Text = ![Nama barang]
            txtHargaSatuan.Text = "" & ![Harga Barang]
        End If
    End With
End Sub





Private Sub dbcbarang_Click()

    'dapatakan record ang dipilih
    Dim strsql  As String
    
    strsql = "Select * from Barang Where [Kode Barang] = '" & dbcbarang.Text & "'"
    Set rsBarang = conAVB.Execute(strsql, , adCmdText)
    'kirim dari database
    
    With rsBarang
        txtNamaBrg.Text = ![Nama barang]
        txtHargaSatuan.Text = "" & ![Harga Barang]
        nilai = ![Jumlah Barang]
    End With
    txtBanyaknyaBarang.SetFocus
End Sub

Private Sub dbcPelanggan_Change()
'pilih record dari sesuai dengan kode pelanggan yang diganti
    Dim strsql  As String
    
    strsql = "Select [Kode Pelanggan],[Nama Pelanggan] from Pelanggan Where [Kode Pelanggan] = '" & dbcPelanggan.Text & "'"
    Set rsPelanggan = conAVB.Execute(strsql, , adCmdText)
    'kirim data dari database
    
    With rsPelanggan
        If .BOF And .EOF Then
            Exit Sub
        Else
            txtNamaPelanggan.Text = ![Nama Pelanggan]
        End If
    End With
End Sub

Private Sub dbcPelanggan_CLick()
    'Dapatkan record yang dipilih
    Dim strsql  As String
    
    strsql = "Select [Kode Pelanggan],[Nama Pelanggan] from Pelanggan Where [Kode Pelanggan] = '" & dbcPelanggan.Text & "'"
    Set rsPelanggan = conAVB.Execute(strsql, , adCmdText)
    'kirim dari database
    
    With rsPelanggan
        txtNamaPelanggan.Text = ![Nama Pelanggan]
    End With
End Sub

Private Sub Form_Load()
    'buat connection
    Dim strsql As String
    Set conAVB = New ADODB.Connection
    
    conAVB.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "Persist Security Info=False;Data Source=" & App.Path & _
        "\AVB8.mdb;Mode = readwrite"
    conAVB.Open
    
    'Buat recordset
    Set rsPenjualan = New ADODB.Recordset
    strsql = "Select * from Penjualan"
    
    rsPenjualan.Open strsql, conAVB, adOpenDynamic, adLockOptimistic, adCmdText
    If rsPenjualan.RecordCount > 0 Then ' jika ada record data (Data tidak kosong)
        TampilkanData
    End If
End Sub
Private Sub cmdTambah_Click()
    'Tambah data record baru
    If cmdTambah.Caption = "&Tambah" Then
        BersihkanFieldText
        nonAktifkanTombol
        cmdSimpan.Enabled = True
        cmdTambah.Caption = "&Batal"
    Else
        rsPenjualan.CancelUpdate 'batakan penambahan data
        txtNoBon.Locked = True
        AktifkanTombol
        cmdSimpan.Enabled = False 'tombol simpan dinonaktifkan
        cmdTambah.Caption = "&Tambah" 'kembalikan tulisan menjadi tambah
        rsPenjualan.MoveLast
        TampilkanData
    End If
End Sub
Private Sub cmdHapus_Click()
    'Tambah jumlah stock barang yang dijual
    Dim nilai1 As Integer
    nilai1 = nilai + Val(txtBanyaknyaBarang.Text)
        
    Dim strSQL1 As String
    strSQL1 = "Update Barang " & _
              "Set [Jumlah barang] = '" & nilai1 & "' " & _
              "Where [Kode Barang] = '" & dbcbarang.Text & "'"
    conAVB.Execute strSQL1, , adCmdText
    
    'Hapus Record yang aktif
    Dim strsql As String
    strsql = "Delete From Penjualan " _
           & "Where [No Bon] = '" & txtNoBon.Text & "' and  [kode barang] = '" & dbcbarang.Text & "' "
    conAVB.Execute strsql, , adCmdText
    With rsPenjualan
        BersihkanFieldText
        .MoveNext                   'pindah ke record selanjutnya
        If .EOF Then                'hapus jika sampai pada record terakhir
            .MovePrevious
            If .BOF Then
                MsgBox "Data kosong.", vbInformation, "Tidak ada record data"
                nonAktifkanTombol
            End If
        End If
    End With
End Sub
Private Sub BersihkanFieldText()
    'Bersihkan semua text boxes untuk sebuah proses penambahan data
    With txtNoBon
        .Text = ""
        .Locked = False
        .SetFocus
    End With
    txtTanggalBon.Text = ""
    dbcPelanggan.Clear
    txtNamaPelanggan.Text = ""
    dbcbarang.Clear
    txtNamaBrg.Text = ""
    txtHargaSatuan.Text = ""
    txtBanyaknyaBarang.Text = ""
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
Private Sub cmdFirst_Click()
    'Pindah ke record pertama
    On Error Resume Next
    rsPenjualan.MoveFirst
    TampilkanData
End Sub
Private Sub cmdLast_Click()
    'Pindah Ke record terakhir
    On Error Resume Next
    rsPenjualan.MoveLast
    TampilkanData
End Sub
Private Sub cmdNext_Click()

    On Error Resume Next
    With rsPenjualan
        .MoveNext
        If .EOF Then
            .MoveFirst
        End If
        TampilkanData
    End With
End Sub
Private Sub cmdPrevious_Click()

    On Error Resume Next
    With rsPenjualan
        .MovePrevious
        If .BOF Then
            .MoveLast
        End If
        TampilkanData
    End With
End Sub
Private Sub TampilkanData()
    'Transfer dari database
        With rsPenjualan
            txtBanyaknyaBarang.Text = ![Banyaknya Barang]
            txtNoBon.Text = ![No Bon]
            txtTanggalBon.Text = ![Tanggal Bon]
            dbcPelanggan.Text = ![Kode Pelanggan]
            dbcbarang.Text = ![Kode Barang]
            txtHargaSatuan.Text = ![Harga Satuan]
            
        End With
End Sub
Private Sub cmdSimpan_Click()
    'Simpan record yang sedang aktif
    On Error GoTo HandleErrors
    If txtNoBon.Text <> "" Then
        Dim strsql As String
        strsql = "Insert Into Penjualan " _
               & "([No Bon],[Tanggal Bon], [Kode Pelanggan],[Kode Barang],[Harga Barang],[Banyaknya Barang]) " _
               & "VALUES ('" & txtNoBon & "', '" & txtTanggalBon & "','" & dbcPelanggan & "', '" _
               & dbcbarang & "','" & txtHargaSatuan & "','" & txtBanyaknyaBarang & "')"
        conAVB.Execute strsql, , adCmdText
        'Kurangi jumlah stock barang yang dijual
        Dim nilai1 As Integer
        nilai1 = nilai - Val(txtBanyaknyaBarang.Text)
        
        Dim strSQL1 As String
        strSQL1 = "Update Barang " & _
                 "Set [Jumlah barang] = '" & nilai1 & "' " & _
                 "Where [Kode Barang] = '" & dbcbarang.Text & "'"
        conAVB.Execute strSQL1, , adCmdText
        
        rsPenjualan.Requery   'Tambah the new record to the recordset
        txtNoBon.Locked = True     'hapus seluruhnya
        AktifkanTombol
        cmdSimpan.Enabled = False
        cmdTambah.Caption = "&Tambah"
    Else
        Dim x As String
        x = MsgBox("Harus ada data minimal No Bon", vbOKOnly, "Keterangan")
        Exit Sub
    End If

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
    
    Dim StrNoBon As String, strTanggalBon As String
    Dim StrKodePelanggan As String, StrKodeBarang As String
    Dim strHargaSatuan As String, strBanyaknyaBarang As String
    On Error Resume Next
    
    StrNoBon = txtNoBon.Text
    strTanggalBon = txtTanggalBon.Text
    StrKodePelanggan = dbcPelanggan.Text
    StrKodeBarang = dbcbarang.Text
    strHargaSatuan = txtHargaSatuan.Text
    strBanyaknyaBarang = txtBanyaknyaBarang.Text
    'mulai tambah record baru
    rsPenjualan.AddNew
    'letakkan kembali data yang disimpan pada form
    With txtNoBon
        .Text = StrNoBon
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With
    txtTanggalBon.Text = strTanggalBon
    dbcPelanggan.Text = StrKodePelanggan
    dbcbarang.Text = StrKodeBarang
    txtHargaSatuan.Text = strHargaSatuan
    txtBanyaknyaBarang.Text = strBanyaknyaBarang
End Sub
Private Sub cmdTutup_Click()
    Unload Me 'keluar dari program
End Sub
Private Sub txtHargaSatuan_Change()
    Dim Jumlah As Single
    Jumlah = Val(txtHargaSatuan.Text) * Val(txtBanyaknyaBarang.Text)
    On Error GoTo Salah
    txtJumlah.Text = Format(Jumlah, "Rp ###,###,###") & ",-"
    On Error GoTo 0
    Exit Sub
Salah:
End Sub
Private Sub txtBanyaknyaBarang_Change()
    Dim Jumlah As Single
    Jumlah = Val(txtHargaSatuan.Text) * Val(txtBanyaknyaBarang.Text)
    On Error GoTo Salah
    txtJumlah.Text = Format(Jumlah, "Rp ###,###,###") & ",-"
    On Error GoTo 0
    Exit Sub
Salah:
End Sub
Private Sub txtHargaSatuan_KeyPress(KeyAscii As Integer)
    'hanya boleh diisi angka atau backspace
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack) Then
       Beep
       KeyAscii = 0
    End If
End Sub
Private Sub txtTanggalBon_LostFocus()
    Dim cTanggal As String
    Dim cBulan As String
    Dim ctahun As String
    Dim CekTanggal As Date
    cTanggal = Mid(txtTanggalBon.Text, 1, 2)
    cBulan = Mid(txtTanggalBon.Text, 4, 2)
    ctahun = Mid(txtTanggalBon.Text, 7)
    If Not (Val(cTanggal) >= 1 And Val(cTanggal) <= 31) Then
       GoTo SalahTanggal
    End If
    If Not (Val(cBulan) >= 1 And Val(cBulan) <= 12) Then
       GoTo SalahTanggal
    End If
    If Not (Val(ctahun) >= 1900 And Val(ctahun) <= 2200) Then
       GoTo SalahTanggal
    End If
    On Error GoTo SalahTanggal
    CekTanggal = CDate(cTanggal & "-" & cBulan & "-" & ctahun)
    CekTanggal = CDate(txtTanggalBon.Text)
    On Error GoTo 0
    Exit Sub
SalahTanggal:
    If Len(Trim(txtTanggalBon.Text)) = 0 Then
       Exit Sub
    End If
    Dim x As String
    x = MsgBox("Format yang benar: hari-bulan-tahun" & Chr(13) & "Misalnya: 28-12-1974", vbOKOnly, "Penulisan Tanggal Salah")
    txtTanggalBon.SetFocus
End Sub
Private Sub txtNoBon_KeyPress(KeyAscii As Integer)
    'mengubah huruf kecil jadi huruf besar
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub txtNoBon_gotFocus()
    isiCombo
End Sub
Private Sub isiCombo()
    Dim strSQL1 As String
    Dim strSQL2 As String
    Dim StrKodePelanggan As String
    Dim StrKodeBarang As String
    'membuat recorset untuk combo box kode Pelanggan
    Set rsPelanggan = New ADODB.Recordset
    strSQL1 = "Select * from Pelanggan"
    Set rsPelanggan = conAVB.Execute(strSQL1, , adCmdText)
    
    'Isi combo box dengan kode pelanggan
    With rsPelanggan
        Do Until .EOF
            StrKodePelanggan = ![Kode Pelanggan]
            dbcPelanggan.AddItem StrKodePelanggan
            .MoveNext
        Loop
        .MoveFirst
    End With
    rsPelanggan.Close
    'membuat recorset untuk combo box kode Pelanggan
    Set rsBarang = New ADODB.Recordset
    strSQL2 = "Select * from Barang"
    Set rsBarang = conAVB.Execute(strSQL2, , adCmdText)
    'isi combo box dengan kode barang
    With rsBarang
        Do Until .EOF
            StrKodeBarang = ![Kode Barang]
            dbcbarang.AddItem StrKodeBarang
            .MoveNext
        Loop
        .MoveFirst
    End With
    rsBarang.Close
End Sub


