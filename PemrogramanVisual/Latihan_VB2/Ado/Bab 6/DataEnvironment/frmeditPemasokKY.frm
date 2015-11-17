VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmEditPemasok 
   Caption         =   "Mencari Data Kode Pemasok"
   ClientHeight    =   2190
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   5430
   LinkTopic       =   "Form1"
   ScaleHeight     =   2190
   ScaleWidth      =   5430
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Height          =   495
      Left            =   1680
      TabIndex        =   7
      Top             =   1560
      Width           =   2055
   End
   Begin VB.TextBox txtNamaPemasok 
      DataField       =   "Nama Pemasok"
      DataMember      =   "Pemasok"
      DataSource      =   "DE"
      Enabled         =   0   'False
      Height          =   285
      Left            =   1935
      TabIndex        =   2
      Top             =   480
      Width           =   3375
   End
   Begin VB.TextBox txtAlamatPemasok 
      DataField       =   "Alamat Pemasok"
      DataMember      =   "Pemasok"
      DataSource      =   "DE"
      Enabled         =   0   'False
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Top             =   840
      Width           =   3375
   End
   Begin VB.TextBox txtNoTelepon 
      DataField       =   "No Telepon"
      DataMember      =   "Pemasok"
      DataSource      =   "DE"
      Enabled         =   0   'False
      Height          =   285
      Left            =   1935
      TabIndex        =   0
      Top             =   1200
      Width           =   1980
   End
   Begin MSDataListLib.DataCombo cboPemasok 
      Bindings        =   "frmeditPemasokKY.frx":0000
      Height          =   315
      Left            =   1920
      TabIndex        =   8
      Top             =   120
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Kode Pemasok"
      Text            =   ""
      Object.DataMember      =   "Pemasok"
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Kode Pemasok:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   135
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Nama Pemasok:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   510
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Alamat Pemasok:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   900
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "No Telepon:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   1275
      Width           =   1815
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
Private Sub cboPemasok_Click(Area As Integer)
    Dim strPemasok   As String
    Dim vntBookMark As Variant
    
        strPemasok = "[Kode Pemasok] = '" & cboPemasok & "'"
        With DE.rsPemasok
            vntBookMark = .Bookmark
            .MoveFirst
            .Find strPemasok
            If .EOF Then
               .Bookmark = vntBookMark
            End If
            txtNamaPemasok.Text = ![Nama Pemasok]
            txtAlamatPemasok.Text = ![alamat Pemasok]
            txtNoTelepon.Text = ![No Telepon]
                       
        End With

    txtNamaPemasok.Enabled = True
    txtAlamatPemasok.Enabled = True
    txtNoTelepon.Enabled = True
    cmdEdit.Enabled = True

End Sub

Private Sub cmdEdit_Click()
      'Edit record yang anda sedang ditampilkan
    On Error GoTo HandleErrors 'jalankan penyaringan kesalahan / error utk penanggannan kesalahan
    DE.rsPemasok.Update             'Edit record
      
    txtNamaPemasok.Enabled = False
    txtAlamatPemasok.Enabled = False
    txtNoTelepon.Enabled = False
    cmdEdit.Enabled = False
    
cmdEdit_Click_Exit:
    Exit Sub

HandleErrors: 'penyaringan kesalahan / error
    Dim strMessage As String
    Dim errDBError As ADODB.Error
    
    For Each errDBError In DE.conAVB.Errors
        strMessage = strMessage & errDBError.Description & vbCrLf
    Next
    MsgBox strMessage, vbExclamation, " Data Kembar"
    On Error GoTo 0     'matikan fungsi penyaringan kesalahan / error
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub mnuKeluar_Click()
    Unload Me
End Sub
