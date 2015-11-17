VERSION 5.00
Begin VB.Form frmCetakJual 
   Caption         =   "Mencetak Data Penjualan"
   ClientHeight    =   1380
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   5325
   LinkTopic       =   "Form1"
   ScaleHeight     =   1380
   ScaleWidth      =   5325
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtNoBon 
      Height          =   300
      Left            =   2040
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtTanggalBon 
      Height          =   300
      Left            =   2040
      TabIndex        =   2
      Top             =   480
      Width           =   1185
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Preview"
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   840
      Width           =   2055
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "Keluar"
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label8 
      Caption         =   "Tanggal Bon"
      Height          =   210
      Left            =   720
      TabIndex        =   5
      Top             =   480
      Width           =   870
   End
   Begin VB.Label Label7 
      Caption         =   "Nomor Bon"
      Height          =   210
      Left            =   720
      TabIndex        =   4
      Top             =   120
      Width           =   1065
   End
End
Attribute VB_Name = "frmCetakJual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim conAVB As ADODB.Connection
Dim rsPenjualan As ADODB.Recordset
Dim sw, nom As Integer

Private Sub cmdexit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    'Create the connection
    
    Set conAVB = New ADODB.Connection
    
    conAVB.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "Persist Security Info=False;Data Source=" & App.Path & _
        "\AVB8.mdb;Mode = readwrite"
    
    conAVB.Open
End Sub

Private Sub cmdPreview_Click()
    nom = 1
    sw = 1
    PreviewPenjualan.Show
    cetaklayar
End Sub
Private Sub cetaklayar()
    Dim grs As String
    Dim strsql As String
  
    Set rsPenjualan = New ADODB.Recordset
    strsql = "Select * from Penjualan where [No bon]= '" & txtNoBon & "' and [Tanggal Bon]= '" & txtTanggalBon & "'"
    rsPenjualan.Open strsql, conAVB, adOpenDynamic, adLockOptimistic, adCmdText
    With rsPenjualan
        .MoveFirst
        Do While Not .EOF
            If sw = 1 Then
                PreviewPenjualan.FontBold = True
                PreviewPenjualan.FontSize = 14
                PreviewPenjualan.Print
                PreviewPenjualan.Print "Laporan data Penjualan"
                PreviewPenjualan.FontSize = 10
                PreviewPenjualan.Print Tab(0); "No Bon: ";
                PreviewPenjualan.Print Tab(15); ![No Bon]
                PreviewPenjualan.Print Tab(0); "Tanggal: ";
                PreviewPenjualan.Print Tab(15); ![Tanggal Bon]
                grs = String$(92, "+")
                PreviewPenjualan.FontBold = False
                PreviewPenjualan.FontSize = 8
                PreviewPenjualan.Print
                PreviewPenjualan.Print Tab(0); grs;
                PreviewPenjualan.Print Tab(2); "No";
                PreviewPenjualan.Print Tab(5); "Kode";
                PreviewPenjualan.Print Tab(18); "Kode";
                PreviewPenjualan.Print Tab(32); "Harga";
                PreviewPenjualan.Print Tab(42); "Banyak";
                PreviewPenjualan.Print Tab(54); "Jumlah";
                PreviewPenjualan.Print Tab(5); "Pelanggan";
                PreviewPenjualan.Print Tab(18); "Barang ";
                PreviewPenjualan.Print Tab(32); "Barang";
                PreviewPenjualan.Print Tab(42); "Satuan";
                PreviewPenjualan.FontBold = False
                PreviewPenjualan.Print Tab(0); grs;
                sw = 0
            End If
            PreviewPenjualan.Print Tab(2); Format(nom, "###");
            PreviewPenjualan.Print Tab(5); ![Kode Pelanggan];
            PreviewPenjualan.Print Tab(18); ![Kode Barang];
            PreviewPenjualan.Print Tab(32); ![Harga Barang];
            PreviewPenjualan.Print Tab(42); ![Banyaknya Barang];
            PreviewPenjualan.Print Tab(54); ![Banyaknya Barang] * ![Harga Barang];
            
            .MoveNext
            nom = nom + 1
        Loop
        PreviewPenjualan.Print Tab(0); grs;
    End With
End Sub


