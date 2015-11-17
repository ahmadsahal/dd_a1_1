VERSION 5.00
Begin VB.Form frmEditPelanggan 
   Caption         =   "Mengedit Data Pelanggan"
   ClientHeight    =   2370
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   5100
   LinkTopic       =   "Form1"
   ScaleHeight     =   2370
   ScaleWidth      =   5100
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtTeleponpelanggan 
      Height          =   285
      Left            =   1560
      TabIndex        =   8
      Top             =   1200
      Width           =   1980
   End
   Begin VB.TextBox txtAlamatpelanggan 
      Height          =   285
      Left            =   1560
      TabIndex        =   6
      Top             =   840
      Width           =   3375
   End
   Begin VB.TextBox txtNamaPelanggan 
      Height          =   285
      Left            =   1560
      TabIndex        =   4
      Top             =   480
      Width           =   3375
   End
   Begin VB.ComboBox cboPelanggan 
      Height          =   315
      Left            =   1560
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   2655
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Enabled         =   0   'False
      Height          =   615
      Left            =   1920
      TabIndex        =   0
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Telepon pelanggan:"
      Height          =   255
      Index           =   3
      Left            =   45
      TabIndex        =   7
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Alamat pelanggan:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Nama Pelanggan:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Kode Pelanggan:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
   Begin VB.Menu cmdKeluar 
      Caption         =   "Keluar"
   End
End
Attribute VB_Name = "frmEditPelanggan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Nama Program   : Ado Database
'Programmer     : Kok Yung
'Tanggal        : 11/2001
'Purpose        : mengedit data tanpa menggunakan
'                 objek visual ( dengan menggunakan perintah SQL)
'Folder         : ADO/BAB3/2_Mengedit Data

Option Explicit
Dim mblnIsDirty     As Boolean
Dim conAVB As ADODB.Connection
Dim rsPelanggan As ADODB.Recordset

Private Sub cmdKeluar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    'Membuat connection
    Dim strSQL  As String
    Dim strName As String
    Set conAVB = New ADODB.Connection
    
    conAVB.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "Persist Security Info=False;Data Source=" & App.Path & _
        "\AVB.mdb;Mode = readwrite"
    conAVB.Open
    
    'membuat recordset untuk combo box
    Set rsPelanggan = New ADODB.Recordset
    strSQL = "Select * from Pelanggan"
    Set rsPelanggan = conAVB.Execute(strSQL, , adCmdText)
    
    'Mengisi combo box dengan Kode Pelanggan
    With rsPelanggan
        Do Until .EOF
            strName = ![Kode Pelanggan]
            cboPelanggan.AddItem strName
            .MoveNext
        Loop
        .MoveFirst
    End With
    rsPelanggan.Close
End Sub
Private Sub cboPelanggan_Click()
 'tampilkan data Pelanggan yang dipilh
   'nama dipilih dari daftar
    'Perform the filter and hide controls
    Dim strSQL As String
    Set rsPelanggan = New ADODB.Recordset
    strSQL = "Select * from Pelanggan Where [Kode Pelanggan]= '" & cboPelanggan.Text & "'"
    Set rsPelanggan = conAVB.Execute(strSQL, , adCmdText)
    
 'transfer data dari database kemudian Tampilkan data
     
        With rsPelanggan
            If .BOF And .EOF Then
                Exit Sub
            Else
                
                txtNamaPelanggan.Text = "" & ![Nama Pelanggan]
                txtAlamatpelanggan.Text = "" & ![Alamat Pelanggan]
                txtTeleponpelanggan.Text = ![Telepon Pelanggan]
            End If
        End With
    cmdEdit.Enabled = True
    
    
End Sub

Private Sub ClearTextFields()
    'bersihkan text box untuk sebuah penambahan data
    txtNamaPelanggan.Text = ""
    txtAlamatpelanggan.Text = ""
    txtTeleponpelanggan.Text = ""
End Sub

Private Sub cmdEdit_Click()
 'simpan record yang sedang aktif
    On Error GoTo HandleErrors
    
    Dim strSQL As String
    
    strSQL = "Update Pelanggan " & _
             "Set [Nama Pelanggan] = '" & txtNamaPelanggan.Text & "', " & _
             "[Alamat Pelanggan] = '" & txtAlamatpelanggan.Text & "', " & _
             "[Telepon Pelanggan] = '" & txtTeleponpelanggan.Text & "' " & _
             "Where [Kode Pelanggan] = '" & cboPelanggan.Text & "'"
    conAVB.Execute strSQL, , adCmdText
    ClearTextFields
    
cmdSimpan_Click_Exit:
Exit Sub

HandleErrors:
    Dim strMessage As String
    Dim errDBError As ADODB.Error
    
    For Each errDBError In conAVB.Errors
        strMessage = strMessage & Err.Description & vbCrLf
    Next
    MsgBox strMessage, vbExclamation, "Provider Error"
   
End Sub


