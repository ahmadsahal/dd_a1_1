VERSION 5.00
Begin VB.Form frmEditPelanggan 
   Caption         =   "Mengedit Data Pelanggan"
   ClientHeight    =   3150
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   5700
   LinkTopic       =   "Form1"
   ScaleHeight     =   3150
   ScaleWidth      =   5700
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   960
      TabIndex        =   9
      Top             =   1560
      Width           =   3615
      Begin VB.CommandButton cmdNext 
         Caption         =   ">"
         Default         =   -1  'True
         Height          =   255
         Left            =   1920
         TabIndex        =   13
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "<"
         Height          =   255
         Left            =   1080
         TabIndex        =   12
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdFirst 
         Caption         =   "I<"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdLast 
         Caption         =   ">I"
         Height          =   255
         Left            =   2760
         TabIndex        =   10
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.TextBox txtTeleponpelanggan 
      DataField       =   "Telepon pelanggan"
      DataMember      =   "Pelanggan"
      DataSource      =   "DE"
      Height          =   285
      Left            =   1800
      TabIndex        =   8
      Top             =   1230
      Width           =   1980
   End
   Begin VB.TextBox txtAlamatpelanggan 
      DataField       =   "Alamat pelanggan"
      DataMember      =   "Pelanggan"
      DataSource      =   "DE"
      Height          =   285
      Left            =   1770
      TabIndex        =   6
      Top             =   850
      Width           =   3375
   End
   Begin VB.TextBox txtNamaPelanggan 
      DataField       =   "Nama Pelanggan"
      DataMember      =   "Pelanggan"
      DataSource      =   "DE"
      Height          =   285
      Left            =   1770
      TabIndex        =   4
      Top             =   470
      Width           =   3375
   End
   Begin VB.TextBox txtKodePelanggan 
      DataField       =   "Kode Pelanggan"
      DataMember      =   "Pelanggan"
      DataSource      =   "DE"
      Enabled         =   0   'False
      Height          =   285
      Left            =   1770
      TabIndex        =   2
      Top             =   120
      Width           =   990
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Enabled         =   0   'False
      Height          =   615
      Left            =   2040
      TabIndex        =   0
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Telepon pelanggan:"
      Height          =   255
      Index           =   3
      Left            =   165
      TabIndex        =   7
      Top             =   1260
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Alamat pelanggan:"
      Height          =   255
      Index           =   2
      Left            =   165
      TabIndex        =   5
      Top             =   885
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Nama Pelanggan:"
      Height          =   255
      Index           =   1
      Left            =   165
      TabIndex        =   3
      Top             =   495
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Kode Pelanggan:"
      Height          =   255
      Index           =   0
      Left            =   165
      TabIndex        =   1
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
'Proyek     :Menginput data Pelanggan
'tanggal    :22 November 2001
'Programmer :Kok Yung
'Descripsi  :Menampilakan, menambah, dan menghapus data pada database AVB2,
'            menggunakan ADO dan DE
                                   
Option Explicit

Private Sub cmdEdit_Click()
     'Edit record yang anda sedang ditampilkan
    On Error GoTo HandleErrors 'jalankan penyaringan kesalahan / error utk penanggannan kesalahan
    
    DE.rsPelanggan.Update             'Edit record
   
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
Private Sub cmdFirst_Click()
    'Move to first record
    On Error Resume Next
    DE.rsPelanggan.MoveFirst
End Sub

Private Sub cmdKeluar_Click()
    Unload Me
End Sub

Private Sub cmdLast_Click()
    'Move to last record
    
    On Error Resume Next
    DE.rsPelanggan.MoveLast
End Sub

Private Sub cmdNext_Click()
    'Move to next record
    
    On Error Resume Next
    With DE.rsPelanggan
        .MoveNext
        If .EOF Then
            .MoveLast
        End If
    End With
End Sub

Private Sub cmdPrevious_Click()
    'Move to previous record
    
    On Error Resume Next
    With DE.rsPelanggan
        .MovePrevious
        If .BOF Then
            .MoveFirst
        End If
    End With
End Sub

Private Sub cmdTutup_Click()
'Keluar dari proyek
    Unload Me
End Sub

Private Sub txtNamaPelanggan_Change()
    cmdEdit.Enabled = True
End Sub
