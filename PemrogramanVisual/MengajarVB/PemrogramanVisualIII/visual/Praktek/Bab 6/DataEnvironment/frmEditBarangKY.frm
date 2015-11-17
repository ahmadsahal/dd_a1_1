VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmEditBarang 
   Caption         =   "Edit Barang"
   ClientHeight    =   2625
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4110
   LinkTopic       =   "Form1"
   ScaleHeight     =   2625
   ScaleWidth      =   4110
   StartUpPosition =   2  'CenterScreen
   Begin MSDataListLib.DataCombo cboBarang 
      Bindings        =   "frmEditBarangKY.frx":0000
      Height          =   315
      Left            =   1320
      TabIndex        =   9
      Top             =   360
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Kode Barang"
      Text            =   ""
      Object.DataMember      =   "Barang"
   End
   Begin VB.TextBox txtJumlahBarang 
      DataField       =   "Jumlah Barang"
      DataMember      =   "Barang"
      DataSource      =   "DE"
      Enabled         =   0   'False
      Height          =   285
      Left            =   1320
      TabIndex        =   8
      Top             =   1440
      Width           =   660
   End
   Begin VB.CommandButton cmdTutup 
      Caption         =   "Tutup"
      Height          =   375
      Left            =   2400
      TabIndex        =   6
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Enabled         =   0   'False
      Height          =   372
      Left            =   840
      TabIndex        =   5
      Top             =   2040
      Width           =   972
   End
   Begin VB.TextBox txtNamaBarang 
      DataField       =   "Nama Barang"
      DataMember      =   "Barang"
      DataSource      =   "DE"
      Enabled         =   0   'False
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Top             =   720
      Width           =   2655
   End
   Begin VB.TextBox txtHargaBarang 
      DataField       =   "Harga Barang"
      DataMember      =   "Barang"
      DataSource      =   "DE"
      Enabled         =   0   'False
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Top             =   1080
      Width           =   1320
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Jumlah Barang:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Kode Barang:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Nama Barang:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Harga Barang:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   1815
   End
End
Attribute VB_Name = "frmEditBarang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboBarang_Click(Area As Integer)
    Dim strBarang   As String
    Dim vntBookMark As Variant
    
        strBarang = "[Kode Barang] = '" & cboBarang & "'"
        With DE.rsBarang
            vntBookMark = .Bookmark     'Save pointer to current record
            .MoveFirst
            .Find strBarang
            If .EOF Then
               .Bookmark = vntBookMark 'Return to previous record
            End If
            txtNamaBarang.Text = ![Nama Barang]
            txtHargaBarang.Text = ![Harga Barang]
            txtJumlahBarang.Text = ![Jumlah Barang]
        End With

    txtNamaBarang.Enabled = True
    txtHargaBarang.Enabled = True
    cmdEdit.Enabled = True

End Sub

Private Sub cmdEdit_Click()
    'Simpan record yang anda sedang ditampilkan
    
    On Error GoTo HandleErrors 'jalankan penyaringan kesalahan / error utk penanggannan kesalahan
    DE.rsBarang.Update             'simpan record
    cmdEdit.Enabled = False
    txtNamaBarang.Enabled = False
    txtHargaBarang.Enabled = False
    
    
cmdEdit_Click_Exit:
    Exit Sub

HandleErrors: 'penyaringan kesalahan / error
    Dim strMessage As String
    Dim errDBError As ADODB.Error
    
    For Each errDBError In DE.conAVB.Errors
        strMessage = strMessage & errDBError.Description & vbCrLf
    Next
    MsgBox strMessage, vbExclamation, "Provider error"
    On Error GoTo 0     'matikan fungsi penyaringan kesalahan / error
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub


