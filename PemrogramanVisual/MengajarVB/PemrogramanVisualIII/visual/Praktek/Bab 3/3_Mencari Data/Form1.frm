VERSION 5.00
Begin VB.Form frmEditPemasok 
   Caption         =   "Mencari Data Kode Pemasok"
   ClientHeight    =   2190
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   5760
   LinkTopic       =   "Form1"
   ScaleHeight     =   2190
   ScaleWidth      =   5760
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Height          =   495
      Left            =   1680
      TabIndex        =   12
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   5175
      Begin VB.TextBox txtCari 
         Height          =   285
         Left            =   1920
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdCari 
         Caption         =   "Cari"
         Height          =   255
         Left            =   3240
         TabIndex        =   9
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Cari Kode Pemasok"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.TextBox txtKodePemasok 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1935
      TabIndex        =   3
      Top             =   120
      Width           =   990
   End
   Begin VB.TextBox txtNamaPemasok 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1935
      TabIndex        =   2
      Top             =   480
      Width           =   3375
   End
   Begin VB.TextBox txtAlamatPemasok 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Top             =   840
      Width           =   3375
   End
   Begin VB.TextBox txtNoTelepon 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1935
      TabIndex        =   0
      Top             =   1200
      Width           =   1980
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Kode Pemasok:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   135
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Nama Pemasok:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   510
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Alamat Pemasok:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   900
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "No Telepon:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   4
      Top             =   1275
      Width           =   1815
   End
   Begin VB.Menu mnucari 
      Caption         =   "Cari"
   End
   Begin VB.Menu mnuKeluar 
      Caption         =   "Keluar"
   End
End
Attribute VB_Name = "frmEditPemasok"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Nama Program   : Ado Database
'Programmer     : Kok Yung
'Tanggal        : 11/2001
'Purpose        : Mencari data tanpa menggunakan
'                 objek visual ( dengan menggunakan perintah SQL)
'Folder         : ADO/BAB3/3_Mencari Data

Option Explicit
Dim conAVB As ADODB.Connection
Dim rsPemasok As ADODB.Recordset
Dim Pemasok     As ADODB.Command

Private Sub cmdEdit_Click()
 'simpan record yang sedang aktif
    On Error GoTo HandleErrors
    Dim strSQL As String
    
    strSQL = "Update Pemasok  " & _
             "Set [Nama Pemasok] = '" & txtNamaPemasok.Text & "', " & _
            "[Alamat Pemasok] = '" & txtAlamatPemasok.Text & "', " & _
             "[No Telepon] = '" & txtNoTelepon.Text & "' " & _
             "Where [Kode Pemasok] = '" & txtKodePemasok.Text & "'"
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

Private Sub Form_Load()
    'Create the connection
    Dim strSQL As String
    Set conAVB = New ADODB.Connection
    
    conAVB.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "Persist Security Info=False;Data Source=" & App.Path & _
        "\AVB.mdb;Mode = readwrite"
    conAVB.Open
    
    'Create the recordset
    Set rsPemasok = New ADODB.Recordset
    strSQL = "Select * from Pemasok"
    
    rsPemasok.Open strSQL, conAVB, adOpenDynamic, adLockOptimistic, adCmdText
End Sub
Private Sub cmdCari_Click()
    'Perform the filter and hide controls
    Dim strSQL As String
    Set rsPemasok = New ADODB.Recordset
    strSQL = "Select * from Pemasok Where [Kode Pemasok]= '" & txtCari.Text & "'"
    Set rsPemasok = conAVB.Execute(strSQL, , adCmdText)
    
    With rsPemasok
        If .BOF And .EOF Then
            MsgBox "Data yang dicari tidak ada.", vbInformation, "Keterangan"
        Else
            txtKodePemasok.Text = ![Kode Pemasok]
            txtNamaPemasok.Text = ![Nama Pemasok]
            txtAlamatPemasok.Text = ![Alamat Pemasok]
            txtNoTelepon.Text = ![No Telepon]
            txtNamaPemasok.Enabled = True
            txtAlamatPemasok.Enabled = True
            txtNoTelepon.Enabled = True
            txtNamaPemasok.SetFocus
        End If
    End With
    Frame1.Visible = False
End Sub
Private Sub mnucari_Click()
    Frame1.Visible = True
    txtCari.SetFocus
End Sub

Private Sub mnuKeluar_Click()
    Unload Me
End Sub
Private Sub ClearTextFields()
    'bersihkan text box untuk sebuah penambahan data
    txtKodePemasok.Text = ""
    txtNamaPemasok.Text = ""
    txtAlamatPemasok.Text = ""
    txtNoTelepon.Text = ""
    txtCari.Text = ""
End Sub

