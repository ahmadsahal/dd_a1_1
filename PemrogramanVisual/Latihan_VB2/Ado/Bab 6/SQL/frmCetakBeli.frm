VERSION 5.00
Begin VB.Form frmCetakBeli 
   Caption         =   "Cetak Pembelian"
   ClientHeight    =   1380
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1380
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdexit 
      Caption         =   "Keluar"
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   840
      Width           =   2055
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Preview"
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   840
      Width           =   2055
   End
   Begin VB.TextBox txtTanggalFaktur 
      Height          =   300
      Left            =   1560
      TabIndex        =   1
      Top             =   480
      Width           =   1185
   End
   Begin VB.TextBox txtNoFaktur 
      Height          =   300
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Nomor Faktur"
      Height          =   210
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   1065
   End
   Begin VB.Label Label8 
      Caption         =   "Tanggal Bon"
      Height          =   210
      Left            =   240
      TabIndex        =   4
      Top             =   480
      Width           =   870
   End
End
Attribute VB_Name = "frmCetakBeli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim conAVB As ADODB.Connection
Dim rsPembelian As ADODB.Recordset
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
    PreviewPembelian.Show
    cetaklayar
End Sub
Private Sub cetaklayar()
    If txtNoFaktur.Text = "" Or txtTanggalFaktur = "" Then
        MsgBox "Data Tidak boleh kosong"
        Exit Sub
    Else
      Dim grs As String
      Dim strsql As String
    
      Set rsPembelian = New ADODB.Recordset
      strsql = "Select * from Pembelian where [No Faktur]= '" & txtNoFaktur & "' and [Tanggal Faktur]= '" & txtTanggalFaktur & "'"
      rsPembelian.Open strsql, conAVB, adOpenDynamic, adLockOptimistic, adCmdText
      With rsPembelian
          .MoveFirst
          Do While Not .EOF
              If sw = 1 Then
                  PreviewPembelian.FontBold = True
                  PreviewPembelian.FontSize = 14
                  PreviewPembelian.Print
                  PreviewPembelian.Print "Laporan data Pembelian"
                  PreviewPembelian.FontSize = 10
                  PreviewPembelian.Print Tab(0); "No Faktur: ";
                  PreviewPembelian.Print Tab(15); ![No Faktur]
                  PreviewPembelian.Print Tab(0); "Tanggal: ";
                  PreviewPembelian.Print Tab(15); ![Tanggal Faktur]
                  grs = String$(92, "+")
                  PreviewPembelian.FontBold = False
                  PreviewPembelian.FontSize = 8
                  PreviewPembelian.Print
                  PreviewPembelian.Print Tab(0); grs;
                  PreviewPembelian.Print Tab(2); "No";
                  PreviewPembelian.Print Tab(5); "Kode";
                  PreviewPembelian.Print Tab(18); "Kode";
                  PreviewPembelian.Print Tab(32); "Harga";
                  PreviewPembelian.Print Tab(42); "Banyak";
                  PreviewPembelian.Print Tab(54); "Jumlah";
                  PreviewPembelian.Print Tab(5); "Pemasok";
                  PreviewPembelian.Print Tab(18); "Barang ";
                  PreviewPembelian.Print Tab(32); "Barang";
                  PreviewPembelian.Print Tab(42); "Satuan";
                  PreviewPembelian.FontBold = False
                  PreviewPembelian.Print Tab(0); grs;
                  sw = 0
              End If
              PreviewPembelian.Print Tab(2); Format(nom, "###");
              PreviewPembelian.Print Tab(5); ![Kode Pemasok];
              PreviewPembelian.Print Tab(18); ![Kode Barang];
              PreviewPembelian.Print Tab(32); ![Harga Satuan];
              PreviewPembelian.Print Tab(42); ![Banyaknya Barang];
              PreviewPembelian.Print Tab(54); ![Banyaknya Barang] * ![Harga Satuan];
              
              .MoveNext
              nom = nom + 1
          Loop
          PreviewPembelian.Print Tab(0); grs;
      End With
    End If
End Sub



