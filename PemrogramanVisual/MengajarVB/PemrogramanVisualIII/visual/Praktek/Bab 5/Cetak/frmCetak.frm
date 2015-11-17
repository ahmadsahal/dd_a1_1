VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Mencetak Data Barang"
   ClientHeight    =   735
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3735
   LinkTopic       =   "Form1"
   ScaleHeight     =   735
   ScaleWidth      =   3735
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdBatal 
      Caption         =   "Batal"
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton cmdCetak 
      Caption         =   "Cetak Data"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ConAVB As ADODB.Connection
Dim rsBarang As ADODB.Recordset
Private Sub Form_Load()
    'Create the connection
    
    Set ConAVB = New ADODB.Connection
    
    ConAVB.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "Persist Security Info=False;Data Source=" & App.Path & _
        "\AVB.mdb;Mode = readwrite"
    ConAVB.Open
End Sub
Private Sub cmdBatal_Click()
    End
End Sub

Private Sub cmdCetak_Click()
    Dim Grs As String
    Dim No As Integer
    Dim Hal As Integer
    Dim strsql As String
    
    Set rsBarang = New ADODB.Recordset
    strsql = "Select * from Barang"
    rsBarang.Open strsql, ConAVB, adOpenDynamic, adLockOptimistic, adCmdText
    
    'Tentukan font
    Printer.Font = "Times news Roman"
    'Bawa ke record barang paling atas
    rsBarang.MoveFirst
    'Bawa head printer ke awal halaman
    Printer.CurrentX = 0
    Printer.CurrentY = 0
    'Mulai pengulangan
    No = 0
    Hal = 0
    Do While Not rsBarang.EOF
        'Cetak judul tabel
        Hal = Hal + 1
        Printer.Print Tab(46); "Daftar Barang"
        Printer.Print Tab(78); "Hal :"; Format(Hal, "###")
        Grs = String$(92, "+")
        Printer.Print Tab(0); Grs;
        Printer.Print Tab(2); "No";
        Printer.Print Tab(8); "Kode";
        Printer.Print Tab(18); "Nama";
        Printer.Print Tab(38); "Harga";
        Printer.Print Tab(8); "Barang";
        Printer.Print Tab(18); "Barang";
        Printer.Print Tab(0); Grs;
        'Mulai  Pengulangan cetak isi tabel
        Do While Not rsBarang.EOF
            No = No + 1
            Printer.Print Tab(2); No;
            Printer.Print Tab(8); rsBarang![Kode Barang];
            Printer.Print Tab(18); rsBarang![Nama Barang];
            Printer.Print Tab(38); rsBarang![Harga satuan];
            rsBarang.MoveNext
        Loop
        Printer.Print Tab(0); Grs;
        Printer.NewPage
    Loop
    Printer.EndDoc
End Sub




