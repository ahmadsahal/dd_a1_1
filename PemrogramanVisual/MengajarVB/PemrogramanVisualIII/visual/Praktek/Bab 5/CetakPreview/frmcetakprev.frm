VERSION 5.00
Begin VB.Form frmCetakPemasok 
   Caption         =   "Cetak Pemasok"
   ClientHeight    =   615
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4875
   LinkTopic       =   "Form1"
   ScaleHeight     =   615
   ScaleWidth      =   4875
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdexit 
      Caption         =   "Keluar"
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Preview"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frmCetakPemasok"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ConAVB As ADODB.Connection
Dim rsBarang As ADODB.Recordset
Dim sw, nom As Integer

Private Sub cmdexit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    'Create the connection
  
    Set ConAVB = New ADODB.Connection
    
       ConAVB.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
    "Persist Security Info=False;Data Source=" & App.Path & _
    "\AVB.mdb;Mode = readwrite"
    ConAVB.Open
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
  
    Set rsBarang = New ADODB.Recordset
    strsql = "Select * from Barang"
    rsBarang.Open strsql, ConAVB, adOpenDynamic, adLockOptimistic, adCmdText
    With rsBarang
        .MoveFirst
        Do While Not .EOF
            If sw = 1 Then
                PreviewPenjualan.FontBold = True
                PreviewPenjualan.FontSize = 14
                PreviewPenjualan.Print
                PreviewPenjualan.Print "Laporan data Barang"
    grs = String$(92, "+")
                PreviewPenjualan.FontBold = False
                PreviewPenjualan.FontSize = 10
                PreviewPenjualan.Print
                PreviewPenjualan.Print Tab(0); grs;
                PreviewPenjualan.Print Tab(2); "No";
                PreviewPenjualan.Print Tab(8); "Kode";
                PreviewPenjualan.Print Tab(18); "Nama";
                PreviewPenjualan.Print Tab(38); "Harga";
                PreviewPenjualan.Print Tab(8); "Barang";
                PreviewPenjualan.Print Tab(18); "Barang";
                PreviewPenjualan.Print Tab(38); "Barang";
                PreviewPenjualan.FontBold = False
                PreviewPenjualan.Print Tab(0); grs;
                sw = 0
            End If
            PreviewPenjualan.Print Tab(3); Format(nom, "###");
            PreviewPenjualan.Print Tab(8); ![Kode Barang];
            PreviewPenjualan.Print Tab(18); ![Nama Barang];
            PreviewPenjualan.Print Tab(38); ![Harga Barang]; ";"
            .MoveNext
            nom = nom + 1
        Loop
        PreviewPenjualan.Print Tab(0); grs;
    End With
End Sub


