VERSION 5.00
Begin VB.Form frmEditPemasok 
   Caption         =   "Edit Data Pemasok"
   ClientHeight    =   2925
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   5385
   LinkTopic       =   "Form1"
   ScaleHeight     =   2925
   ScaleWidth      =   5385
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      Width           =   5055
      Begin VB.CommandButton cmdNext 
         Caption         =   ">"
         Default         =   -1  'True
         Height          =   372
         Left            =   2520
         TabIndex        =   14
         Top             =   240
         Width           =   972
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "<"
         Height          =   372
         Left            =   1320
         TabIndex        =   13
         Top             =   240
         Width           =   972
      End
      Begin VB.CommandButton cmdFirst 
         Caption         =   "I<"
         Height          =   372
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   972
      End
      Begin VB.CommandButton cmdLast 
         Caption         =   ">I"
         Height          =   372
         Left            =   3720
         TabIndex        =   11
         Top             =   240
         Width           =   972
      End
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Height          =   375
      Left            =   720
      TabIndex        =   9
      Top             =   2400
      Width           =   1695
   End
   Begin VB.CommandButton cmdTutup 
      Caption         =   "Tutup"
      Height          =   375
      Left            =   2640
      TabIndex        =   8
      Top             =   2400
      Width           =   1575
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
      Height          =   285
      Left            =   1935
      TabIndex        =   2
      Top             =   480
      Width           =   3375
   End
   Begin VB.TextBox txtAlamatPemasok 
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Top             =   840
      Width           =   3375
   End
   Begin VB.TextBox txtNoTelepon 
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
End
Attribute VB_Name = "frmEditPemasok"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Program    :Mengedit Data Pemasok
'Programer  :Kok Yung
'Tanggal    :15 November 2001
'Deskripsi  :Mengedit data Pemasok dalam database AVB8 menggunakan
'            Perintah SQL dan tombol navigasi
'folder     :C:\ADO\Bab 6\SQL

Option Explicit
Dim conAVB As ADODB.Connection
Dim rsPemasok As ADODB.Recordset
Dim Pemasok     As ADODB.Command

Private Sub cmdEdit_Click()
 'simpan record yang sedang aktif
    On Error GoTo HandleErrors
    Dim strsql As String
    
    strsql = "Update Pemasok " & _
             "Set [Nama Pemasok] = '" & txtNamaPemasok.Text & "', " & _
             "[Alamat Pemasok] = '" & txtAlamatPemasok.Text & "', " & _
             "[No Telepon] = '" & txtNoTelepon.Text & "' " & _
             "Where [Kode Pemasok] = '" & txtKodePemasok.Text & "'"
    conAVB.Execute strsql, , adCmdText
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

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
    Dim strsql As String
    Set conAVB = New ADODB.Connection
    
    conAVB.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "Persist Security Info=False;Data Source=" & App.Path & _
        "\AVB8.mdb;Mode = readwrite"
    conAVB.Open
    
    'Create the recordset
    Set rsPemasok = New ADODB.Recordset
    strsql = "Select * from Pemasok"
    
    rsPemasok.Open strsql, conAVB, adOpenDynamic, adLockOptimistic, adCmdText
End Sub


Private Sub cmdFirst_Click()
    On Error Resume Next
    rsPemasok.MoveFirst
    TampilkanData
End Sub

Private Sub cmdLast_Click()
    On Error Resume Next
    rsPemasok.MoveLast
    TampilkanData
End Sub

Private Sub cmdNext_Click()
    On Error Resume Next
    With rsPemasok
        .MoveNext
        If .EOF Then
            .MoveFirst
        End If
        TampilkanData
    End With
End Sub

Private Sub cmdPrevious_Click()
    On Error Resume Next
    With rsPemasok
        .MovePrevious
        If .BOF Then
            .MoveLast
        End If
        TampilkanData
    End With
End Sub
Private Sub TampilkanData()
    'Transfer from database
        With rsPemasok
            txtKodePemasok.Text = ![Kode Pemasok]
            txtNamaPemasok.Text = ![Nama Pemasok]
            txtAlamatPemasok.Text = ![ALamat Pemasok]
            txtNoTelepon.Text = ![no telepon]
        End With
    
End Sub

Private Sub ClearTextFields()
    'bersihkan text box untuk sebuah penambahan data
            txtKodePemasok.Text = ""
            txtNamaPemasok.Text = ""
            txtAlamatPemasok.Text = ""
            txtNoTelepon.Text = ""
End Sub






