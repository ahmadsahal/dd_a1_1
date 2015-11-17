VERSION 5.00
Begin VB.Form frmCariBarang 
   Caption         =   "Cari Barang"
   ClientHeight    =   1845
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1845
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtKodeBarang 
      DataField       =   "Kode Barang"
      Enabled         =   0   'False
      Height          =   285
      Left            =   1320
      TabIndex        =   2
      Top             =   120
      Width           =   990
   End
   Begin VB.TextBox txtNamaBarang 
      DataField       =   "Nama Barang"
      Enabled         =   0   'False
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Top             =   480
      Width           =   3135
   End
   Begin VB.TextBox txtHargaBarang 
      DataField       =   "Harga Barang"
      Enabled         =   0   'False
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Top             =   840
      Width           =   1320
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Kode Barang:"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Nama Barang:"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   4
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Harga Barang:"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   1815
   End
   Begin VB.Menu mnuCariBarang 
      Caption         =   "Cari Barang"
   End
End
Attribute VB_Name = "frmCariBarang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Proyek     :Mencari data Barang
'tanggal    :22 November 2001
'Programmer :Kok Yung
'Descripsi  :Mencari data Barang pada database AVB8,
'            menggunakan Perintah SQL

Option Explicit
Dim conAVB    As ADODB.Connection
Dim Barang     As ADODB.Command
Dim rsBarang   As ADODB.Recordset
Dim mstrSQL As String

Private Sub Form_Load()
    'Create the Connection and get the data
    Set conAVB = New ADODB.Connection
    Set Barang = New ADODB.Command
    
    conAVB.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "Persist Security Info=False;Data Source=" & App.Path & _
        "\AVB8.mdb;Mode = readwrite"
    conAVB.Open

    
    'Create the Command object
    Set Barang = New ADODB.Command
    Set Barang.ActiveConnection = conAVB
    
    'Create the Recordset and set to all records
    Set rsBarang = New ADODB.Recordset
    'Set up SQL for all records
    
    mstrSQL = "Select * from Barang"
    Barang.CommandText = mstrSQL
    Set rsBarang = Barang.Execute
    
    'Bind the text boxes
    Set txtKodeBarang.DataSource = rsBarang
    Set txtNamaBarang.DataSource = rsBarang
    Set txtHargaBarang.DataSource = rsBarang
End Sub

Private Sub mnuCariBarang_Click()
    Dim mstrSQL As String
    Dim strKode     As String
        
    strKode = InputBox("Masukkan kode Barang yang akan anda cari." & vbCrLf & _
        "", "Cari Data Barang")

    mstrSQL = "Select [Kode Barang], [Nama Barang], [Harga Barang]from Barang " & _
              "Where  [Kode Barang] = '" & strKode & "'"
    Barang.CommandText = mstrSQL
    With rsBarang
        .Requery
        If .BOF And .EOF Then
            MsgBox "Kode Barang tidak ada", vbInformation, "Keterangan"
            Barang.CommandText = mstrSQL
            .Requery
        End If
    End With
End Sub
