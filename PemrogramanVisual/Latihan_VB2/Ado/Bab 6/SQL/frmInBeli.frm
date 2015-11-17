VERSION 5.00
Begin VB.Form frmInBeli 
   Caption         =   "Pemasukan Data Pembelian"
   ClientHeight    =   4440
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   ScaleHeight     =   4440
   ScaleWidth      =   4800
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox dbcBarang 
      Height          =   315
      Left            =   1560
      TabIndex        =   25
      Top             =   1200
      Width           =   1695
   End
   Begin VB.ComboBox dbcPemasok 
      Height          =   315
      Left            =   1560
      TabIndex        =   24
      Top             =   840
      Width           =   1695
   End
   Begin VB.CommandButton cmdTutup 
      Caption         =   "Tutup"
      Height          =   375
      Left            =   3480
      TabIndex        =   22
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "&Simpan"
      Enabled         =   0   'False
      Height          =   372
      Left            =   1560
      TabIndex        =   21
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton cmdTambah 
      Caption         =   "&Tambah"
      CausesValidation=   0   'False
      Height          =   372
      Left            =   600
      TabIndex        =   20
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton cmdHapus 
      Caption         =   "&Hapus"
      Height          =   372
      Left            =   2520
      TabIndex        =   19
      Top             =   3960
      Width           =   855
   End
   Begin VB.Frame fraNavigation 
      Height          =   735
      Left            =   720
      TabIndex        =   14
      Top             =   3120
      Width           =   3495
      Begin VB.CommandButton cmdFirst 
         Caption         =   "I<"
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "<"
         Height          =   375
         Left            =   960
         TabIndex        =   17
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   ">"
         Default         =   -1  'True
         Height          =   375
         Left            =   1800
         TabIndex        =   16
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdLast 
         Caption         =   ">I"
         Height          =   375
         Left            =   2640
         TabIndex        =   15
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.TextBox txtNoFaktur 
      Height          =   300
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   972
   End
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
      Left            =   1560
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   2640
      Width           =   2028
   End
   Begin VB.TextBox txtBanyaknyaBarang 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   1560
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   2280
      Width           =   1164
   End
   Begin VB.TextBox txtHargaSatuan 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   1560
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   1920
      Width           =   1164
   End
   Begin VB.TextBox txtNamaBrg 
      Height          =   300
      Left            =   1560
      TabIndex        =   3
      Top             =   1560
      Width           =   2985
   End
   Begin VB.TextBox txtTglFaktur 
      Height          =   300
      Left            =   1560
      TabIndex        =   2
      Top             =   480
      Width           =   1068
   End
   Begin VB.Label Label10 
      Caption         =   "dd-mm-yyyy"
      Height          =   255
      Left            =   2760
      TabIndex        =   23
      Top             =   480
      Width           =   1335
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
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label8 
      Caption         =   "Tanggal  Faktur"
      Height          =   210
      Left            =   240
      TabIndex        =   12
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "No Faktur"
      Height          =   210
      Left            =   240
      TabIndex        =   11
      Top             =   120
      Width           =   1065
   End
   Begin VB.Label Label6 
      Caption         =   "Banyak"
      Height          =   210
      Left            =   240
      TabIndex        =   10
      Top             =   2280
      Width           =   1065
   End
   Begin VB.Label Label5 
      Caption         =   "Harga Satuan"
      Height          =   210
      Left            =   240
      TabIndex        =   9
      Top             =   1920
      Width           =   1065
   End
   Begin VB.Label Label4 
      Caption         =   "Nama Barang"
      Height          =   210
      Left            =   240
      TabIndex        =   8
      Top             =   1560
      Width           =   1590
   End
   Begin VB.Label Label3 
      Caption         =   "Kode Barang"
      Height          =   210
      Left            =   240
      TabIndex        =   7
      Top             =   1200
      Width           =   1065
   End
   Begin VB.Label Label1 
      Caption         =   "Kode Pemasok"
      Height          =   210
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   1185
   End
End
Attribute VB_Name = "frmInBeli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Proyek     :Menginput data Pembelian
'tanggal    :22 November 2001
'Programmer :Kok Yung
'Descripsi  :Menampilakan, menambah, dan menghapus Pembelian barang pada database AVB8,
'            menggunakan perintah SQL
                                   
Option Explicit
Dim conAVB As ADODB.Connection
Dim rsPembelian As ADODB.Recordset
Dim rsBarang As ADODB.Recordset
Dim rsPemasok As ADODB.Recordset
Dim nilai As Integer





Private Sub Form_Load()
    'Membuat koneksi
    Dim strsql As String
    Set conAVB = New ADODB.Connection
    
    conAVB.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "Persist Security Info=False;Data Source=" & App.Path & _
        "\AVB8.mdb;Mode = readwrite"
    conAVB.Open
    
    'Buat recordset
    Set rsPembelian = New ADODB.Recordset
    strsql = "Select * from Pembelian"
    
    rsPembelian.Open strsql, conAVB, adOpenDynamic, adLockOptimistic, adCmdText
    If rsPembelian.RecordCount > 0 Then ' jika ada record data (Data tidak kosong)
        TampilkanData
    End If
End Sub
Private Sub dbcBarang_Click()
    'dapatkan record dari data yang dipilih
    Dim strsql  As String
    
    strsql = "Select * from Barang Where [Kode Barang] = '" & dbcBarang.Text & "'"
    Set rsBarang = conAVB.Execute(strsql, , adCmdText)
    'Transfer dari database
    
    With rsBarang
        txtNamaBrg.Text = ![Nama barang]
        txtHargaSatuan.Text = "" & ![Harga Barang]
        nilai = ![Jumlah Barang]
    End With
    txtBanyaknyaBarang.SetFocus
End Sub
Private Sub dbcBarang_Change()

    'dapatkan record dari data yang dipilih
    Dim strsql  As String
    
    strsql = "Select * from Barang Where [Kode Barang] = '" & dbcBarang.Text & "'"
    Set rsBarang = conAVB.Execute(strsql, , adCmdText)
    'Transfer dari database
    
    With rsBarang
        If .BOF And .EOF Then
            Exit Sub
        Else
            txtNamaBrg.Text = ![Nama barang]
            txtHargaSatuan.Text = "" & ![Harga Barang]
            nilai = ![Jumlah Barang]
        End If
    End With
    txtBanyaknyaBarang.SetFocus
End Sub

Private Sub cmdTambah_Click()
    'Tambah data record baru
    If cmdTambah.Caption = "&Tambah" Then
        BersihkanFieldText
        nonAktifkanTombol
        cmdSimpan.Enabled = True
        cmdTambah.Caption = "&Batal"
    Else
        rsPembelian.CancelUpdate
        txtNoFaktur.Locked = True
        AktifkanTombol
        cmdSimpan.Enabled = False
        cmdTambah.Caption = "&Tambah"
        rsPembelian.MoveLast
        TampilkanData
    End If
End Sub
Private Sub cmdHapus_Click()
    'Kurangi jumlah stock
    Dim nilai1 As Integer
    nilai1 = nilai - Val(txtBanyaknyaBarang.Text)
      
    Dim strSQL1 As String
    strSQL1 = "Update Barang " & _
              "Set [Jumlah barang] = '" & nilai1 & "' " & _
              "Where [Kode Barang] = '" & dbcBarang.Text & "'"
    conAVB.Execute strSQL1, , adCmdText
        
    
    'Reset the Tambah button
    Dim strsql As String
    strsql = "Delete From Pembelian " _
           & "Where [No Faktur] = '" & txtNoFaktur.Text & "' and [kode barang] = '" & dbcBarang.Text & "'"
    conAVB.Execute strsql, , adCmdText
    With rsPembelian
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
Private Sub BersihkanFieldText()
    'Bersihkan semua text boxes untuk sebuah proses penambahan data
    With txtNoFaktur
        .Text = ""
        .Locked = False
        .SetFocus
    End With
    txtTglFaktur.Text = ""
    dbcPemasok.Clear
    dbcBarang.Clear
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
    rsPembelian.MoveFirst
    TampilkanData
End Sub
Private Sub cmdLast_Click()
    'Pindah Ke record terakhir
    On Error Resume Next
    rsPembelian.MoveLast
    TampilkanData
End Sub
Private Sub cmdNext_Click()

    On Error Resume Next
    With rsPembelian
        .MoveNext
        If .EOF Then
            .MoveFirst
        End If
        TampilkanData
    End With
End Sub
Private Sub cmdPrevious_Click()

    On Error Resume Next
    With rsPembelian
        .MovePrevious
        If .BOF Then
            .MoveLast
        End If
        TampilkanData
    End With
End Sub
Private Sub TampilkanData()
    'Transfer from database
        With rsPembelian
            txtNoFaktur.Text = ![No Faktur]
            txtTglFaktur.Text = ![Tanggal Faktur]
        dbcPemasok.Text = ![Kode Pemasok]
            dbcBarang.Text = ![Kode Barang]
            txtHargaSatuan.Text = ![Harga Satuan]
            txtBanyaknyaBarang.Text = ![Banyaknya Barang]
        End With
End Sub
Private Sub cmdSimpan_Click()
    'Simpan record yang sedang aktif
    On Error GoTo HandleErrors
    If txtNoFaktur.Text <> "" Then
        Dim strsql As String
        strsql = "Insert Into Pembelian " _
               & "([No Faktur],[Tanggal Faktur], [Kode Pemasok],[Kode Barang],[Harga Satuan],[Banyaknya Barang]) " _
               & "VALUES ('" & txtNoFaktur & "', '" & txtTglFaktur & "','" & dbcPemasok & "','" _
               & dbcBarang & "','" & txtHargaSatuan & "','" & txtBanyaknyaBarang & "')"
        conAVB.Execute strsql, , adCmdText
        
        'Tambah jumlah Stock
        Dim nilai1 As Integer
        nilai1 = nilai + Val(txtBanyaknyaBarang.Text)
        
        Dim strSQL1 As String
        strSQL1 = "Update Barang " & _
                 "Set [Jumlah barang] = '" & nilai1 & "' " & _
                 "Where [Kode Barang] = '" & dbcBarang.Text & "'"
        conAVB.Execute strSQL1, , adCmdText
        
        rsPembelian.Requery   'Tambah the new record to the recordset
        txtNoFaktur.Locked = True       'Reset all
        AktifkanTombol
        cmdSimpan.Enabled = False
        cmdTambah.Caption = "&Tambah"
    Else
        Dim x As String
        x = MsgBox("Harus ada data minimal No Faktur", vbOKOnly, "Keterangan")
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
    'Set up a new Tambah to give user another try
    Dim StrNoFaktur As String, strTanggalFaktur As String
    Dim StrKodePemasok As String, StrKodeBarang As String
    Dim strHargaSatuan As String, strBanyaknyaBarang As String
    On Error Resume Next
    'Save contents of form controls
    StrNoFaktur = txtNoFaktur.Text
    strTanggalFaktur = txtTglFaktur.Text
    StrKodePemasok = dbcPemasok.Text
    StrKodeBarang = dbcBarang.Text
    strHargaSatuan = txtHargaSatuan.Text
    strBanyaknyaBarang = txtBanyaknyaBarang.Text
    'Start a new Tambah
    rsPembelian.AddNew
    'Place saved data back on form
    With txtNoFaktur
        .Text = StrNoFaktur
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With
    txtTglFaktur.Text = strTanggalFaktur
   dbcPemasok.Text = StrKodePemasok
    dbcBarang.Text = StrKodeBarang
    txtHargaSatuan.Text = strHargaSatuan
    txtBanyaknyaBarang.Text = strBanyaknyaBarang
End Sub
Private Sub cmdTutup_Click()
    Unload Me 'keluar dari program
End Sub

Private Sub txtBanyaknyaBarang_KeyPress(KeyAscii As Integer)
    'hanya boleh diisi angka atau backspace
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack) Then
       Beep
       KeyAscii = 0
    End If

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
Private Sub txtNoFaktur_KeyPress(KeyAscii As Integer)
    'mengubah huruf kecil jadi huruf besar
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub txtNoFaktur_gotFocus()
    isiCombo
End Sub
Private Sub isiCombo()
    Dim strSQL1 As String
    Dim strSQL2 As String
    Dim StrKodePemasok As String
    Dim StrKodeBarang As String
'Create the recordset for the combo box kode barang
    Set rsPemasok = New ADODB.Recordset
    strSQL1 = "Select * from Pemasok"
    Set rsPemasok = conAVB.Execute(strSQL1, , adCmdText)
    
    'Fill the combo box with names and patient numbers
    With rsPemasok
        Do Until .EOF
            StrKodePemasok = ![Kode Pemasok]
            dbcPemasok.AddItem (StrKodePemasok)
            .MoveNext
        Loop
        .MoveFirst
    End With
    rsPemasok.Close
'Create the recordset for the combo box kode barang
    Set rsBarang = New ADODB.Recordset
    strSQL2 = "Select * from Barang"
    Set rsBarang = conAVB.Execute(strSQL2, , adCmdText)
    'Fill the combo box with names and patient numbers
    With rsBarang
        Do Until .EOF
            StrKodeBarang = ![Kode Barang]
            dbcBarang.AddItem StrKodeBarang
            .MoveNext
        Loop
        .MoveFirst
    End With
    rsBarang.Close
End Sub

Private Sub txtTglFaktur_LostFocus()
    Dim cTanggal As String
    Dim cBulan As String
    Dim ctahun As String
    Dim CekTanggal As Date
    cTanggal = Mid(txtTglFaktur.Text, 1, 2)
    cBulan = Mid(txtTglFaktur.Text, 4, 2)
    ctahun = Mid(txtTglFaktur.Text, 7)
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
    CekTanggal = CDate(txtTglFaktur.Text)
    On Error GoTo 0
    Exit Sub
SalahTanggal:
    If Len(Trim(txtTglFaktur.Text)) = 0 Then
       Exit Sub
    End If
    Dim x As String
    x = MsgBox("Format yang benar: hari-bulan-tahun" & Chr(13) & "Misalnya: 28-12-1974", vbOKOnly, "Penulisan Tanggal Salah")
    txtTglFaktur.SetFocus

End Sub
