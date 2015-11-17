VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "BACA DATA"
   ClientHeight    =   3636
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   4932
   LinkTopic       =   "Form1"
   ScaleHeight     =   3636
   ScaleWidth      =   4932
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Cetak"
      Height          =   372
      Left            =   3600
      TabIndex        =   2
      Top             =   3000
      Width           =   1092
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2532
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4452
      _ExtentX        =   7853
      _ExtentY        =   4466
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   252
      Left            =   240
      TabIndex        =   1
      Top             =   3120
      Width           =   2172
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim Nama, Kode, Gaji As String
    Dim TotalGaji As Currency
    Printer.FontSize = 12
    Open "C:\VB6\GAJI.DAT" For Input As #1
    'JUDUL
    Printer.Print Tab(10); "NAMA"; Tab(40); "GOLONGAN"; Tab(60); "GAJI"
    Printer.Print Tab(10); String(45, "=")
    Do Until EOF(1)
        Input #1, Nama, Kode, Gaji
        TotalGaji = TotalGaji + Gaji
        'ISI
        Printer.Print Tab(10); Nama; Tab(40); Kode; Tab(60); Gaji
    Loop
    Printer.Print Tab(10); String(45, "=")
    Printer.Print Tab(10); "TOTAL GAJI"; Tab(57); Format(TotalGaji, "Currency")
    Close #1
    Printer.EndDoc
End Sub

Private Sub Form_Load()
    Dim LI As ListItem
    Dim Nama, Kode, Gaji As String
    Dim TotalGaji As Currency
    
    ListView1.View = lvwReport
    ListView1.Sorted = True

    'Membuat Judul Kolom (ColumnHeaders) serta mengatur lebar.
    ListView1.ColumnHeaders.Add , , "NAMA", ListView1.Width / 3
    ListView1.ColumnHeaders.Add , , "GOLONGAN", ListView1.Width / 3
    ListView1.ColumnHeaders.Add , , "GAJI", ListView1.Width / 3
    
    Open "C:\VB6\GAJI.DAT" For Input As #1
    Do Until EOF(1)
        Input #1, Nama, Kode, Gaji
        'Masukkan item dan sub item ke dalam list.
        Set LI = ListView1.ListItems.Add(, , Nama)
        LI.SubItems(1) = Kode
        LI.SubItems(2) = Gaji
        TotalGaji = TotalGaji + Gaji
    Loop
    Close #1
    Label1.Caption = "Total Gaji = " + Format(TotalGaji, "Currency")
End Sub

