VERSION 5.00
Begin VB.Form PreviewPembelian 
   Caption         =   "Preview Pembelian"
   ClientHeight    =   7710
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10080
   LinkTopic       =   "Form1"
   ScaleHeight     =   7710
   ScaleWidth      =   10080
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCetak 
      Caption         =   "Cetak"
      Height          =   375
      Left            =   5400
      TabIndex        =   1
      Top             =   6960
      Width           =   1695
   End
   Begin VB.CommandButton cmdKeluar 
      Caption         =   "Keluar"
      Height          =   375
      Left            =   7680
      TabIndex        =   0
      Top             =   6960
      Width           =   1575
   End
End
Attribute VB_Name = "PreviewPembelian"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim conAVB As ADODB.Connection
Dim rsPembelian As ADODB.Recordset
Dim sw, nom As Integer

Private Sub Form_Load()
    'Create the connection

    Set conAVB = New ADODB.Connection
    
    conAVB.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "Persist Security Info=False;Data Source=" & App.Path & _
        "\AVB8.mdb;Mode = readwrite"
    conAVB.Open
End Sub

Private Sub cmdCetak_Click()
    nom = 1
    sw = 1
    cetak
    Printer.EndDoc
End Sub
Private Sub cetak()
 Dim grs As String
    Dim strsql As String
  
    Set rsPembelian = New ADODB.Recordset
    strsql = "Select * from Pembelian where [No Faktur]= '" & frmCetakBeli.txtNoFaktur.Text & "'"
    rsPembelian.Open strsql, conAVB, adOpenDynamic, adLockOptimistic, adCmdText
    With rsPembelian
        .MoveFirst
        Printer.CurrentX = 0
        Printer.CurrentY = 0
        Do While Not .EOF
            If sw = 1 Then
                Printer.FontBold = True
                Printer.FontSize = 14
                Printer.Print
                Printer.Print "Laporan data Pembelian"
                Printer.FontSize = 10
                Printer.Print Tab(0); "No Faktur: ";
                Printer.Print Tab(15); ![No Faktur]
                Printer.Print Tab(0); "Tanggal: ";
                Printer.Print Tab(15); ![Tanggal Faktur]
               
                grs = String$(92, "+")
                Printer.FontBold = False
                Printer.FontSize = 8
                Printer.Print
                Printer.Print Tab(0); grs;
                Printer.Print Tab(2); "No";
                Printer.Print Tab(5); "Kode";
                Printer.Print Tab(18); "Kode";
                Printer.Print Tab(32); "Harga";
                Printer.Print Tab(42); "Banyak";
                Printer.Print Tab(54); "Jumlah";
                Printer.Print Tab(5); "Pemasok";
                Printer.Print Tab(18); "Barang ";
                Printer.Print Tab(32); "Barang";
                Printer.Print Tab(42); "Satuan";
                Printer.FontBold = False
                Printer.Print Tab(0); grs;
                sw = 0
            End If
            Printer.Print Tab(2); Format(nom, "###");
            Printer.Print Tab(5); ![Kode Pemasok];
            Printer.Print Tab(18); ![Kode Barang];
            Printer.Print Tab(32); ![Harga Satuan];
            Printer.Print Tab(42); ![Banyaknya Barang];
            Printer.Print Tab(54); ![Banyaknya Barang] * ![Harga Satuan];
            .MoveNext
            nom = nom + 1
        Loop
        Printer.Print Tab(0); grs;
    End With
    Printer.NewPage
End Sub

Private Sub cmdKeluar_Click()
    Unload Me
End Sub


