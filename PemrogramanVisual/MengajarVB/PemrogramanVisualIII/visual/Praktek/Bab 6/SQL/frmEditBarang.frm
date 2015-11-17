VERSION 5.00
Begin VB.Form frmEditBarang 
   Caption         =   "Edit Barang"
   ClientHeight    =   2145
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4110
   LinkTopic       =   "Form1"
   ScaleHeight     =   2145
   ScaleWidth      =   4110
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cbobarang 
      Height          =   315
      Left            =   1320
      TabIndex        =   9
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox txtJumlahBarang 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1320
      TabIndex        =   8
      Top             =   1200
      Width           =   660
   End
   Begin VB.CommandButton cmdTutup 
      Caption         =   "Tutup"
      Height          =   375
      Left            =   2400
      TabIndex        =   6
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Enabled         =   0   'False
      Height          =   372
      Left            =   840
      TabIndex        =   5
      Top             =   1680
      Width           =   972
   End
   Begin VB.TextBox txtNamaBarang 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Top             =   480
      Width           =   2655
   End
   Begin VB.TextBox txtHargaBarang 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Top             =   840
      Width           =   1320
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Jumlah Barang:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Kode Barang:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Nama Barang:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Harga Barang:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1815
   End
End
Attribute VB_Name = "frmEditBarang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Proyek     :MengEdit data Barang
'tanggal    :22 November 2001
'Programmer :Kok Yung
'Descripsi  :Mengedit data barang pada database AVB8,
'            menggunakan perintah SQL
                                   
Option Explicit
Dim conAVB As ADODB.Connection
Dim rsBarang As ADODB.Recordset


Private Sub cbobarang_Click()
    'Retrieve the selected record
    Dim strsql  As String
    
    strsql = "Select * from Barang Where [Kode Barang] = '" & cbobarang.Text & "'"
    Set rsBarang = conAVB.Execute(strsql, , adCmdText)
    'Transfer from database
    
    With rsBarang
        txtNamaBarang.Text = ![Nama barang]
        txtHargaBarang.Text = ![Harga Barang]
        txtJumlahBarang.Text = ![Jumlah Barang]
    End With
    txtNamaBarang.Enabled = True
    txtHargaBarang.Enabled = True
    cmdEdit.Enabled = True

End Sub

Private Sub Form_Load()
    'Create the connection
    Dim strsql As String
    Set conAVB = New ADODB.Connection
    
    conAVB.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "Persist Security Info=False;Data Source=" & App.Path & _
        "\AVB8.mdb;Mode = readwrite"
    conAVB.Open
    
    'Buat recordset
    Set rsBarang = New ADODB.Recordset
    strsql = "Select * from Barang"
    Set rsBarang = conAVB.Execute(strsql, , adCmdText)
        
    Dim StrKodeBarang As String
    'Fill the combo box with names and patient numbers
    With rsBarang
        Do Until .EOF
            StrKodeBarang = ![Kode Barang]
            cbobarang.AddItem StrKodeBarang
            .MoveNext
        Loop
        .MoveFirst
    End With
End Sub

Private Sub cmdEdit_Click()
     'simpan record yang sedang aktif
    On Error GoTo HandleErrors
    Dim strsql As String
    
    strsql = "Update Barang " & _
             "Set [Nama Barang] = '" & txtNamaBarang.Text & "', " & _
             "[Harga Barang] = '" & txtHargaBarang.Text & "' " & _
             "Where [Kode Barang] = '" & cbobarang.Text & "'"
    conAVB.Execute strsql, , adCmdText
    nonaktifkan
    
cmdEdit_Click_Exit:
Exit Sub

HandleErrors:
    Dim strMessage As String
    Dim errDBError As ADODB.Error
    
    For Each errDBError In conAVB.Errors
        strMessage = strMessage & Err.Description & vbCrLf
    Next
    MsgBox strMessage, vbExclamation, "Provider Error"
   
End Sub
Private Sub nonaktifkan()
    txtNamaBarang.Enabled = False
    txtHargaBarang.Enabled = False
    cmdEdit.Enabled = False
End Sub
Private Sub cmdTutup_Click()
    Unload Me
End Sub


