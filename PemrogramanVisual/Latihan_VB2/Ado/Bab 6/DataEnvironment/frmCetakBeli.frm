VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCetakBeli 
   Caption         =   "Mencetak Data Pembelian"
   ClientHeight    =   4080
   ClientLeft      =   48
   ClientTop       =   276
   ClientWidth     =   5328
   LinkTopic       =   "Form1"
   ScaleHeight     =   4080
   ScaleWidth      =   5328
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbSize 
      Height          =   288
      Left            =   4224
      TabIndex        =   8
      Top             =   2976
      Width           =   972
   End
   Begin VB.ComboBox cmbFont 
      Height          =   288
      Left            =   1056
      TabIndex        =   7
      Top             =   2976
      Width           =   2700
   End
   Begin VB.TextBox txtTglSD 
      Height          =   288
      Left            =   2112
      TabIndex        =   3
      Top             =   1248
      Width           =   1260
   End
   Begin VB.TextBox txtTglDari 
      Height          =   288
      Left            =   2112
      TabIndex        =   2
      Top             =   960
      Width           =   1260
   End
   Begin VB.TextBox txtNamaSpl 
      Height          =   300
      Left            =   2112
      TabIndex        =   1
      Top             =   672
      Width           =   2988
   End
   Begin VB.TextBox txtKodeSpl 
      Height          =   288
      Left            =   2112
      TabIndex        =   0
      Top             =   384
      Width           =   972
   End
   Begin VB.Data dbSupplier 
      Caption         =   "Supplier"
      Connect         =   "Access"
      DatabaseName    =   "C:\LatVB6\Pembelian.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   324
      Left            =   3552
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "Supplier"
      Top             =   0
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.Data dbBarang 
      Caption         =   "Barang"
      Connect         =   "Access"
      DatabaseName    =   "C:\LatVB6\Pembelian.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   324
      Left            =   1824
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "Barang"
      Top             =   0
      Visible         =   0   'False
      Width           =   1644
   End
   Begin VB.ComboBox cmbPrinter 
      Height          =   288
      Left            =   1056
      TabIndex        =   6
      Top             =   2688
      Width           =   4140
   End
   Begin VB.Frame Frame1 
      Caption         =   "Urutan Data"
      Height          =   492
      Left            =   480
      TabIndex        =   13
      Top             =   1728
      Width           =   4428
      Begin VB.OptionButton optUrutNoFak 
         Caption         =   "Urut Nomor Faktur"
         Height          =   204
         Left            =   2400
         TabIndex        =   5
         Top             =   192
         Width           =   1548
      End
      Begin VB.OptionButton optUrutKodeBrg 
         Caption         =   "Urut Kode Barang"
         Height          =   204
         Left            =   192
         TabIndex        =   4
         Top             =   192
         Width           =   2028
      End
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   204
      Left            =   1056
      TabIndex        =   9
      Top             =   2400
      Width           =   3852
      _ExtentX        =   6795
      _ExtentY        =   360
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Data dbBeli 
      Caption         =   "Beli"
      Connect         =   "Access"
      DatabaseName    =   "C:\LatVB6\Pembelian.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   324
      Left            =   96
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "Beli"
      Top             =   0
      Visible         =   0   'False
      Width           =   1644
   End
   Begin VB.CommandButton cmdBatal 
      Caption         =   "Batal"
      Height          =   492
      Left            =   3072
      TabIndex        =   11
      Top             =   3456
      Width           =   1644
   End
   Begin VB.CommandButton cmdCetak 
      Caption         =   "Cetak"
      Height          =   492
      Left            =   672
      TabIndex        =   10
      Top             =   3456
      Width           =   1644
   End
   Begin VB.Label Label8 
      Caption         =   "Size"
      Height          =   204
      Left            =   3840
      TabIndex        =   20
      Top             =   2976
      Width           =   588
   End
   Begin VB.Label Label7 
      Caption         =   "Font"
      Height          =   204
      Left            =   384
      TabIndex        =   19
      Top             =   2976
      Width           =   588
   End
   Begin VB.Label Label4 
      Caption         =   "S/d Tanggal"
      Height          =   204
      Left            =   864
      TabIndex        =   18
      Top             =   1248
      Width           =   1260
   End
   Begin VB.Label Label3 
      Caption         =   "Mulai Tanggal"
      Height          =   204
      Left            =   864
      TabIndex        =   17
      Top             =   960
      Width           =   1260
   End
   Begin VB.Label Label2 
      Caption         =   "Nama"
      Height          =   204
      Left            =   864
      TabIndex        =   16
      Top             =   672
      Width           =   1164
   End
   Begin VB.Label Label1 
      Caption         =   "Kode Supplier"
      Height          =   204
      Left            =   864
      TabIndex        =   15
      Top             =   384
      Width           =   1164
   End
   Begin VB.Label Label6 
      Caption         =   "Printer"
      Height          =   204
      Left            =   384
      TabIndex        =   14
      Top             =   2688
      Width           =   876
   End
   Begin VB.Label Label5 
      Caption         =   "Proses"
      Height          =   204
      Left            =   384
      TabIndex        =   12
      Top             =   2400
      Width           =   588
   End
End
Attribute VB_Name = "frmCetakBeli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim P As Printer

Private Sub cmdBatal_Click()
    ProgressBar1.Value = ProgressBar1.Min
    frmCetakBeli.Hide
    frmMenu.Show
End Sub

Private Sub cmbPrinter_Click()
    'ulang sebanyak obyek printer
    For Each P In Printers
        'jika nama printer sama dengan yang dipilih
        If P.DeviceName = cmbPrinter.Text Then
           'set printer aktif dengan printer yang dipilih
           Set Printer = P
           'keluar dari pengulangan
           Exit For
        End If
    Next
End Sub

Private Sub cmbFont_Click()
    Printer.FontName = cmbFont.Text
End Sub

Private Sub cmbSize_Click()
    Printer.FontSize = cmbSize.Text
End Sub

Private Sub cmdCetak_Click()
    Dim MNo, MHal, MBaris, X As Integer
    Dim MNilai, MSubTotal, MTotal As Single
    Dim MGrs As String
    If Len(txtTglDari.Text) = 0 Or Len(txtTglSD.Text) = 0 Then
       Beep
       X = MsgBox("Data tanggal harus diisi!", vbOKOnly)
       Exit Sub
    End If
    MTgAwal = CDate(txtTglDari.Text)
    MTgAkhir = CDate(txtTglSD.Text)
    
    'On Error hanya dipasang jika program sudah benar
    On Error GoTo SalahCetak
    With dbBeli.Recordset
     ProgressBar1.Min = 1
     ProgressBar1.Max = .RecordCount
     'tentukan urutan data barang
     dbBarang.Recordset.Index = "KodeBrg"
     'tentukan urutan pencetakan berdasar pilihan
     If optUrutKodeBrg.Value = True Then
        .Index = "KodeBrg"
     Else
        .Index = "NoFaktur"
     End If
     
     'bawa pointer ke awal record yang dicetak
     .MoveFirst
     If Len(txtKodeSpl.Text) = 0 Then
        Do While Not (!TgFaktur >= MTgAwal)
           .MoveNext
           If .EOF Then
              Exit Do
           End If
        Loop
     Else
        Do While Not ((!TgFaktur >= MTgAwal) And (!KodeSpl = txtKodeSpl.Text))
           .MoveNext
           If .EOF Then
              Exit Do
           End If
        Loop
     End If
     'jika tidak ada data yang memenuhi syarat, batalkan mencetak
     If .EOF Then
        Beep
        X = MsgBox("Data tidak ada yang memenuhi syarat", vbOKOnly)
        Exit Sub
     End If
    
     'bawa head Printer ke awal halaman
     Printer.CurrentX = 0
     Printer.CurrentY = 0
     'mulai pengulangan
     MNo = 0
     MHal = 0
     MTotal = 0
     Do
        'cetak judul tabel
        MHal = MHal + 1
        'judul saja yang dicetak besar dan tebal
        Printer.FontSize = Val(cmbSize.Text) * 2
        Printer.FontBold = True
        Printer.Print "DAFTAR PEMBELIAN"
        Printer.Print "PT. XYZ"
        Printer.FontSize = cmbSize.Text
        Printer.FontBold = False
        Printer.Print
        Printer.Print "Kode Supplier: "; IIf(Len(txtKodeSpl.Text) = 0, "Semua", txtKodeSpl.Text)
        Printer.Print "Nama Supplier: "; IIf(Len(txtKodeSpl.Text) = 0, "Semua", txtNamaSpl.Text)
        Printer.Print "Periode      : "; MTgAwal & " s/d " & MTgAkhir;
        Printer.Print Tab(100); "Hal :"; Format(MHal, "###")
        MGrs = String$(115, "-")
        Printer.Print MGrs
        Printer.Print Tab(5); "No.";
        Printer.Print Tab(10); "Kode";
        Printer.Print Tab(17); "Nomor";
        Printer.Print Tab(29); "Tanggal";
        Printer.Print Tab(41); "Kode";
        Printer.Print Tab(49); "Nama";
        Printer.Print Tab(81); "Harga";
        Printer.Print Tab(91); "Banyak";
        Printer.Print Tab(101); "Nilai"
        Printer.Print Tab(10); "Spl.";
        Printer.Print Tab(17); "Faktur";
        Printer.Print Tab(29); "Faktur";
        Printer.Print Tab(41); "Barang";
        Printer.Print Tab(49); "Barang";
        Printer.Print Tab(81); "Satuan";
        Printer.Print Tab(91); "Barang";
        Printer.Print Tab(101); "Pembelian"
        Printer.Print MGrs
        MBaris = 0
        MSubTotal = 0
        'mulai pengulangan cetak isi tabel
        Do
           'lompati record yang tidak memenuhi syarat
           If Len(txtKodeSpl.Text) = 0 Then
              Do While Not (!TgFaktur >= MTgAwal)
                 .MoveNext
                 If .EOF Then
                    Exit Do
                 End If
              Loop
           Else
              Do While Not ((!TgFaktur >= MTgAwal) And (!KodeSpl = txtKodeSpl.Text))
                 .MoveNext
                 If .EOF Then
                    Exit Do
                 End If
              Loop
           End If
           If .EOF Then
              Exit Do
           End If
                      
           'sesuaikan record tabel Barang dengan kode barang di tabel Beli
           dbBarang.Recordset.Seek "=", !KodeBrg
           MNo = MNo + 1
           MNilai = !Banyak * !Harga
           ProgressBar1.Value = MNo
           
           Printer.Print Tab(1); RKanan(MNo, "###,###");
           Printer.Print Tab(10); !KodeSpl;
           Printer.Print Tab(17); !NoFaktur;
           Printer.Print Tab(29); !TgFaktur;
           Printer.Print Tab(41); !KodeBrg;
           Printer.Print Tab(49); dbBarang.Recordset!Nama;
           Printer.Print Tab(81); RKanan(!Harga, "#,###,###");
           Printer.Print Tab(91); RKanan(!Banyak, "#,###,###");
           Printer.Print Tab(101); RKanan(MNilai, "###,###,###")
           
           MSubTotal = MSubTotal + MNilai
           MBaris = MBaris + 1
           .MoveNext
           If .EOF Then
              Exit Do
           End If
        Loop Until MBaris > 55
        MTotal = MTotal + MSubTotal
        Printer.Print MGrs
        Printer.Print "Sub Total Pembelian:"; Tab(99); RKanan(MSubTotal, "#,###,###,###")
        Printer.Print "Total Pembelian:"; Tab(99); RKanan(MTotal, "#,###,###,###")
        Printer.Print MGrs
        Printer.NewPage
        If .EOF Then
           Exit Do
        End If
     Loop
     Printer.EndDoc
     ProgressBar1.Value = .RecordCount
    End With
    On Error GoTo 0
    cmdBatal_Click
    Exit Sub
SalahCetak:
    Beep
    X = MsgBox("Printer ERROR!" & Chr(13) & "Betulkan printer, lalu klik OK", vbOKCancel)
    If X = 0 Then
       Resume
    Else
       Printer.KillDoc
       cmdBatal_Click
    End If
End Sub

Private Sub Form_Activate()
    ProgressBar1.Visible = True
    cmdCetak.Default = True
    txtKodeSpl.SetFocus
    txtNamaSpl.Enabled = False
    txtNamaSpl.BackColor = vbButtonFace
    optUrutKodeBrg.Value = True
    'text cmbPrinter diisi nama printer aktif
    cmbPrinter.Text = Printer.DeviceName
    'jika belum ada daftar pilihan dalam cmbPrinter
    If cmbPrinter.ListCount = 0 Then
       'semua nama printer yang tersedia dimasukkan ke dalam daftar pilihan
       For Each P In Printers
           cmbPrinter.AddItem P.DeviceName
       Next
    End If
    'text cmbFont diisi nama font aktif
    cmbFont.Text = Printer.FontName
    'jika belum ada daftar pilihan dalam cmbFont
    If cmbFont.ListCount = 0 Then
       'semua nama font yang tersedia dimasukkan ke dalam daftar pilihan
       For i = 0 To Printer.FontCount - 1
           cmbFont.AddItem Printer.Fonts(i)
       Next
    End If
    'text cmbSize diisi ukuran font aktif
    cmbSize.Text = CInt(Printer.FontSize)
    'jika belum ada daftar pilihan dalam cmbSize
    If cmbSize.ListCount = 0 Then
       'isi daftar pilihan sesuai keinginan kita
       For i = 8 To 72 Step 2
           cmbSize.AddItem i
       Next
    End If
End Sub

Private Sub txtKodeSpl_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Function RKanan(NData, CFormat) As String
    'fungsi untuk format rata kanan suatu data numerik
    RKanan = Format(NData, CFormat)
    RKanan = Space(Len(CFormat) - Len(RKanan)) + RKanan
End Function

Private Sub txtKodeSpl_Change()
    If Len(txtKodeSpl.Text) < 5 Then
       Exit Sub
    End If
    With dbSupplier.Recordset
     .Index = "KodeSpl"
     .Seek "=", txtKodeSpl
     If .NoMatch Then
        Beep
        X = MsgBox("Kode Supplier tidak ada!", vbOKOnly)
        Exit Sub
     End If
     txtNamaSpl.Text = !Nama
    End With
End Sub

Private Sub txtKodeSpl_LostFocus()
    If Len(txtKodeSpl) > 0 And Len(txtKodeSpl.Text) < 5 Then
       Beep
       txtKodeSpl.Text = ""
    End If
End Sub

Private Sub txtTglDari_LostFocus()
    Dim MTgl As Date
    If Len(txtTglDari.Text) = 0 Then
       txtTglDari.Text = "01-01-1990"
       Exit Sub
    End If
    On Error GoTo SalahTanggal
    MTgl = CDate(txtTglDari.Text)
    On Error GoTo 0
    Exit Sub
SalahTanggal:
    Beep
    X = MsgBox("Tanggal tidak sah!" & Chr(13) & "Format tanggal adalah : dd-mm-yyyy", vbOKOnly)
    txtTglDari.SetFocus
End Sub

Private Sub txtTglSD_LostFocus()
    Dim MTgl As Date
    If Len(txtTglSD.Text) = 0 Then
       txtTglSD.Text = "01-01-2100"
       Exit Sub
    End If
    On Error GoTo SalahTanggal
    MTgl = CDate(txtTglSD.Text)
    On Error GoTo 0
    Exit Sub
SalahTanggal:
    Beep
    X = MsgBox("Tanggal tidak sah!" & Chr(13) & "Format tanggal adalah : dd-mm-yyyy", vbOKOnly)
    txtTglSD.SetFocus
End Sub

