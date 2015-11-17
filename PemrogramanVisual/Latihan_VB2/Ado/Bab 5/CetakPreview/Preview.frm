VERSION 5.00
Begin VB.Form PreviewPenjualan 
   Caption         =   "Preview"
   ClientHeight    =   7395
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9210
   LinkTopic       =   "Form2"
   ScaleHeight     =   7395
   ScaleWidth      =   9210
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdKeluar 
      Caption         =   "Keluar"
      Height          =   375
      Left            =   7200
      TabIndex        =   1
      Top             =   6960
      Width           =   1575
   End
   Begin VB.CommandButton cmdCetak 
      Caption         =   "Cetak"
      Height          =   375
      Left            =   4920
      TabIndex        =   0
      Top             =   6960
      Width           =   1695
   End
End
Attribute VB_Name = "PreviewPenjualan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim ConAVB As ADODB.Connection
Dim rsBarang As ADODB.Recordset
Dim sw, nom As Integer

Private Sub Form_Load()
    'Create the connection
    
    Set ConAVB = New ADODB.Connection
    
    ConAVB.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
     "Persist Security Info=False;Data Source=" & App.Path & _
        "\AVB.mdb;Mode = readwrite"
    ConAVB.Open
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
  
    Set rsBarang = New ADODB.Recordset
    strsql = "Select * from Barang"
    rsBarang.Open strsql, ConAVB, adOpenDynamic, adLockOptimistic, adCmdText
    With rsBarang
        .MoveFirst
        Printer.CurrentX = 0
        Printer.CurrentY = 0
        Do While Not .EOF
            If sw = 1 Then
                Printer.FontBold = True
                Printer.FontSize = 14
                Printer.Print
                Printer.Print "Laporan data Barang"
                grs = String$(92, "+")
                Printer.FontBold = False
                Printer.FontSize = 10
                Printer.Print
                Printer.Print Tab(0); grs;
                Printer.Print Tab(2); "No";
                Printer.Print Tab(8); "Kode";
                Printer.Print Tab(18); "Nama";
                Printer.Print Tab(38); "Harga";
                Printer.Print Tab(8); "Barang";
                Printer.Print Tab(18); "Barang";
                Printer.Print Tab(38); "Barang";
                Printer.FontBold = False
                Printer.Print Tab(0); grs;
                sw = 0
            End If
            Printer.Print Tab(3); Format(nom, "###");
            Printer.Print Tab(8); ![Kode Barang];
            Printer.Print Tab(18); ![Nama Barang];
            Printer.Print Tab(38); ![Harga Barang]; ";"
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





    


