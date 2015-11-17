VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCetakBrg 
   Caption         =   "Mencetak Data Barang"
   ClientHeight    =   4140
   ClientLeft      =   48
   ClientTop       =   276
   ClientWidth     =   6252
   LinkTopic       =   "Form1"
   ScaleHeight     =   4140
   ScaleWidth      =   6252
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Urutan Data"
      Height          =   492
      Left            =   960
      TabIndex        =   15
      Top             =   2400
      Width           =   4428
      Begin VB.OptionButton Option2 
         Caption         =   "Urut Nama Barang"
         Height          =   204
         Left            =   2208
         TabIndex        =   17
         Top             =   192
         Width           =   2028
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Urut Kode Barang"
         Height          =   204
         Left            =   192
         TabIndex        =   16
         Top             =   192
         Width           =   2028
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Mulai Dari"
      Height          =   972
      Left            =   96
      TabIndex        =   11
      Top             =   1344
      Width           =   2988
      Begin VB.TextBox Text3 
         Height          =   288
         Left            =   1152
         TabIndex        =   2
         Top             =   288
         Width           =   780
      End
      Begin VB.TextBox Text4 
         Height          =   288
         Left            =   1152
         TabIndex        =   3
         Top             =   576
         Width           =   1740
      End
      Begin VB.Label Label4 
         Caption         =   "Kode Barang"
         Height          =   204
         Left            =   96
         TabIndex        =   13
         Top             =   288
         Width           =   1068
      End
      Begin VB.Label Label3 
         Caption         =   "Nama"
         Height          =   204
         Left            =   96
         TabIndex        =   12
         Top             =   576
         Width           =   972
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Mulai Dari"
      Height          =   972
      Left            =   96
      TabIndex        =   8
      Top             =   288
      Width           =   2988
      Begin VB.TextBox Text2 
         Height          =   288
         Left            =   1152
         TabIndex        =   1
         Top             =   576
         Width           =   1740
      End
      Begin VB.TextBox Text1 
         Height          =   288
         Left            =   1152
         TabIndex        =   0
         Top             =   288
         Width           =   780
      End
      Begin VB.Label Label2 
         Caption         =   "Nama"
         Height          =   204
         Left            =   96
         TabIndex        =   10
         Top             =   576
         Width           =   972
      End
      Begin VB.Label Label1 
         Caption         =   "Kode Barang"
         Height          =   204
         Left            =   96
         TabIndex        =   9
         Top             =   288
         Width           =   1068
      End
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   204
      Left            =   1536
      TabIndex        =   5
      Top             =   3072
      Width           =   3852
      _ExtentX        =   6795
      _ExtentY        =   360
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\LatVB6\Pembelian.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   324
      Left            =   2112
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "Barang"
      Top             =   0
      Visible         =   0   'False
      Width           =   2028
   End
   Begin VB.CommandButton cmdBatal 
      Caption         =   "Batal"
      Height          =   492
      Left            =   3552
      TabIndex        =   6
      Top             =   3456
      Width           =   1644
   End
   Begin VB.CommandButton cmdCetak 
      Caption         =   "Cetak"
      Height          =   492
      Left            =   1056
      TabIndex        =   4
      Top             =   3456
      Width           =   1644
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmCetakBrg.frx":0000
      Height          =   1932
      Left            =   3168
      OleObjectBlob   =   "frmCetakBrg.frx":0014
      TabIndex        =   14
      Top             =   384
      Width           =   2988
   End
   Begin VB.Label Label5 
      Caption         =   "Proses"
      Height          =   204
      Left            =   864
      TabIndex        =   7
      Top             =   3072
      Width           =   588
   End
End
Attribute VB_Name = "frmCetakBrg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBatal_Click()
    frmCetakBrg.Hide
    frmMenu.Show
End Sub

Private Sub CetakUKode()
    'mencetak urut kode barang
    Dim MNo, MHal, MBaris As Integer
    Dim MGrs As String
    On Error GoTo SalahCetak
    With Data1.Recordset
     ProgressBar1.Min = 1
     ProgressBar1.Max = .RecordCount
     'tentukan urutan data
     .Index = "KodeBrg"
     'bawa pointer ke awal record yang dicetak
     If Len(Text1.Text) = 0 Then
        'jika kode barang awal dikosongkan,
        'anggap dari record pertama
        .MoveFirst
     Else
        .Seek "=", Text1.Text
     End If
     If Len(Text3.Text) = 0 Then
        'jika kode barang akhir di kosongkan,
        'anggap kode barangnya paling akhir
        MAkhir = "ZZZ"
     Else
        MAkhir = Text3.Text
     End If
     
     'tentukan font
     Printer.Font = "Courier New"
     'bawa head Printer ke awal halaman
     Printer.CurrentX = 0
     Printer.CurrentY = 0
     'mulai pengulangan
     MNo = 0
     MHal = 0
     Do While !Kode <= MAkhir
        'cetak judul tabel
        MHal = MHal + 1
        Printer.Print "DAFTAR BARANG";
        Printer.Print Tab(60); "Hal :"; Format(MHal, "###")
        MGrs = String$(70, "-")
        Printer.Print MGrs
        Printer.Print Tab(5); "No.";
        Printer.Print Tab(10); "Kode";
        Printer.Print Tab(17); "Nama";
        Printer.Print Tab(49); "Satuan";
        Printer.Print Tab(58); "Harga"
        Printer.Print MGrs
        MBaris = 0
        'mulai pengulangan cetak isi tabel
        Do While MBaris <= 55 And !Kode <= MAkhir
           MNo = MNo + 1
           ProgressBar1.Value = MNo
           Printer.Print Tab(1); RKanan(MNo, "###,###");
           Printer.Print Tab(10); !Kode;
           Printer.Print Tab(17); !Nama;
           Printer.Print Tab(49); !Satuan;
           Printer.Print Tab(56); RKanan(!Harga, "#,###,###")
           MBaris = MBaris + 1
           .MoveNext
           If .EOF Then
              Exit Do
           End If
        Loop
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

Private Sub CetakUNama()
    'mencetak urut kode barang
    Dim MNo, MHal, MBaris As Integer
    Dim MGrs As String
    On Error GoTo SalahCetak
    With Data1.Recordset
     ProgressBar1.Min = 1
     ProgressBar1.Max = .RecordCount
     'tentukan urutan data
     .Index = "NamaBrg"
     'bawa pointer ke awal record yang dicetak
     If Len(Text2.Text) = 0 Then
        'jika nama barang awal dikosongkan,
        'anggap dari record pertama
        .MoveFirst
     Else
        .Seek "=", Text2.Text
     End If
     If Len(Text4.Text) = 0 Then
        'jika nama barang akhir di kosongkan,
        'anggap nama barangnya paling akhir
        MAkhir = "ZZZ"
     Else
        MAkhir = Text4.Text
     End If
     
     'tentukan font
     Printer.Font = "Courier New"
     'bawa head Printer ke awal halaman
     Printer.CurrentX = 0
     Printer.CurrentY = 0
     'mulai pengulangan
     MNo = 0
     MHal = 0
     Do While !Nama <= MAkhir
        'cetak judul tabel
        MHal = MHal + 1
        Printer.Print "DAFTAR BARANG";
        Printer.Print Tab(60); "Hal :"; Format(MHal, "###")
        MGrs = String$(70, "-")
        Printer.Print MGrs
        Printer.Print Tab(5); "No.";
        Printer.Print Tab(10); "Kode";
        Printer.Print Tab(17); "Nama";
        Printer.Print Tab(49); "Satuan";
        Printer.Print Tab(58); "Harga"
        Printer.Print MGrs
        MBaris = 0
        'mulai pengulangan cetak isi tabel
        Do While MBaris <= 55 And !Nama <= MAkhir
           MNo = MNo + 1
           ProgressBar1.Value = MNo
           Printer.Print Tab(1); RKanan(MNo, "###,###");
           Printer.Print Tab(10); !Kode;
           Printer.Print Tab(17); !Nama;
           Printer.Print Tab(49); !Satuan;
           Printer.Print Tab(56); RKanan(!Harga, "#,###,###")
           MBaris = MBaris + 1
           .MoveNext
           If .EOF Then
              Exit Do
           End If
        Loop
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

Private Sub cmdCetak_Click()
    If Option1.Value = True Then
       CetakUKode
    Else
       CetakUNama
    End If
End Sub

Private Sub Form_Activate()
    ProgressBar1.Visible = True
    cmdCetak.Default = True
    Text1.SetFocus
    Option1.Value = True
End Sub

Private Sub Text1_Change()
    If Len(Text1.Text) = 0 Then
       Exit Sub
    End If
    With Data1.Recordset
     .Index = "KodeBrg"
     .Seek ">=", Text1.Text
     If Len(Text1.Text) < 6 Then
        Exit Sub
     End If
     .Seek "=", Text1.Text
     If .NoMatch Then
        X = MsgBox("Kode Barang tidak ada!" & Chr(13) & "Kosongkan untuk mulai dari awal", vbOKOnly)
        Beep
        Exit Sub
     End If
    End With
End Sub

Private Sub Text1_LostFocus()
    Text1.Text = Data1.Recordset!Kode
    Text2.Text = Data1.Recordset!Nama
End Sub

Private Sub Text2_LostFocus()
    Text1.Text = Data1.Recordset!Kode
    Text2.Text = Data1.Recordset!Nama
End Sub

Private Sub Text3_LostFocus()
    Text3.Text = Data1.Recordset!Kode
    Text4.Text = Data1.Recordset!Nama
End Sub

Private Sub Text4_LostFocus()
    Text3.Text = Data1.Recordset!Kode
    Text4.Text = Data1.Recordset!Nama
End Sub

Private Sub Text2_Change()
    If Len(Text2.Text) = 30 Then
       Exit Sub
    End If
    With Data1.Recordset
     .Index = "NamaBrg"
     .Seek ">=", Text2.Text
     If Len(Text2.Text) < 30 Then
        Exit Sub
     End If
     .Seek "=", Text2.Text
     If .NoMatch Then
        X = MsgBox("Nama Barang tidak ada!" & Chr(13) & "Kosongkan untuk mulai dari awal", vbOKOnly)
        Beep
        Exit Sub
     End If
    End With
End Sub

Private Sub Text3_Change()
    If Len(Text3.Text) = 0 Then
       Exit Sub
    End If
    With Data1.Recordset
     .Index = "KodeBrg"
     .Seek ">=", Text3.Text
     If Len(Text3.Text) < 6 Then
        Exit Sub
     End If
     .Seek "=", Text3.Text
     If .NoMatch Then
        X = MsgBox("Kode Barang tidak ada!" & Chr(13) & "Kosongkan untuk mulai dari awal", vbOKOnly)
        Beep
        Exit Sub
     End If
    End With
End Sub

Private Sub Text4_Change()
    If Len(Text4.Text) = 0 Then
       Exit Sub
    End If
    With Data1.Recordset
     .Index = "NamaBrg"
     .Seek ">=", Text4.Text
     If Len(Text4.Text) < 30 Then
        Exit Sub
     End If
     .Seek "=", Text4.Text
     If .NoMatch Then
        X = MsgBox("Nama Barang tidak ada!" & Chr(13) & "Kosongkan untuk mulai dari awal", vbOKOnly)
        Beep
        Exit Sub
     End If
    End With
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Function RKanan(NData, CFormat) As String
    'fungsi untuk format rata kanan suatu data numerik
    RKanan = Format(NData, CFormat)
    RKanan = Space(Len(CFormat) - Len(RKanan)) + RKanan
End Function

