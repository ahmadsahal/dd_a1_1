VERSION 5.00
Begin VB.Form Form28 
   Caption         =   "Mencari Data"
   ClientHeight    =   3225
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5475
   LinkTopic       =   "Form1"
   ScaleHeight     =   3225
   ScaleWidth      =   5475
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtCari 
      Height          =   285
      Left            =   2040
      TabIndex        =   14
      Top             =   2760
      Width           =   855
   End
   Begin VB.CommandButton cmdCari 
      Caption         =   "Cari"
      Height          =   495
      Left            =   3120
      TabIndex        =   13
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton cmdLast 
      Caption         =   ">I"
      Height          =   372
      Left            =   3960
      TabIndex        =   11
      Top             =   1920
      Width           =   972
   End
   Begin VB.CommandButton cmdFirst 
      Caption         =   "I<"
      Height          =   372
      Left            =   360
      TabIndex        =   10
      Top             =   1920
      Width           =   972
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "<"
      Height          =   372
      Left            =   1560
      TabIndex        =   9
      Top             =   1920
      Width           =   972
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   5055
      Begin VB.CommandButton cmdNext 
         Caption         =   ">"
         Height          =   375
         Left            =   2640
         TabIndex        =   15
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.TextBox txtNoTelepon 
      DataField       =   "No Telepon"
      DataMember      =   "Pemasok"
      DataSource      =   "DE"
      Enabled         =   0   'False
      Height          =   285
      Left            =   1950
      TabIndex        =   7
      Top             =   1230
      Width           =   1980
   End
   Begin VB.TextBox txtAlamatPemasok 
      DataField       =   "Alamat Pemasok"
      DataMember      =   "Pemasok"
      DataSource      =   "DE"
      Enabled         =   0   'False
      Height          =   285
      Left            =   1950
      TabIndex        =   5
      Top             =   855
      Width           =   3375
   End
   Begin VB.TextBox txtNamaPemasok 
      DataField       =   "Nama Pemasok"
      DataMember      =   "Pemasok"
      DataSource      =   "DE"
      Enabled         =   0   'False
      Height          =   285
      Left            =   1950
      TabIndex        =   3
      Top             =   465
      Width           =   3375
   End
   Begin VB.TextBox txtKodePemasok 
      DataField       =   "Kode Pemasok"
      DataMember      =   "Pemasok"
      DataSource      =   "DE"
      Enabled         =   0   'False
      Height          =   285
      Left            =   1950
      TabIndex        =   1
      Top             =   90
      Width           =   990
   End
   Begin VB.Label Label1 
      Caption         =   "Kode Pemasok"
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "No Telepon:"
      Height          =   255
      Index           =   3
      Left            =   105
      TabIndex        =   6
      Top             =   1275
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Alamat Pemasok:"
      Height          =   255
      Index           =   2
      Left            =   105
      TabIndex        =   4
      Top             =   900
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Nama Pemasok:"
      Height          =   255
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   510
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Kode Pemasok:"
      Height          =   255
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   135
      Width           =   1815
   End
End
Attribute VB_Name = "Form28"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Nama Program   : Ado Database
'Programmer     : Kok Yung
'Tanggal        : 11/2001
'Purpose        : mencari data menggunakan objek visual Data Environment
'Folder         : ADO/BAB2/3_Mencari Data


Private Sub cmdCari_Click()

'Cari berdasarkan Kode Pemasok
    Dim strCari  As String
    Dim vntBookMark As Variant
    
    If txtCari.Text = "" Then
        MsgBox "Masukkan Kode Pemasok sebelum anda menekan tombol OK"
        txtCari.SetFocus
    Else
        strCari = "[Kode Pemasok] = '" & txtCari & "'"
        With DE.rsPemasok
            vntBookMark = .Bookmark     'Save pointer to current record
            .MoveFirst
            .Find strCari
            If .EOF Then
                MsgBox "Kode Tidak Ada", vbExclamation, "Perhatian"
                .Bookmark = vntBookMark 'Return to previous record
            End If
        End With
    End If
    txtCari.Text = ""
End Sub

Private Sub cmdFirst_Click()
'Pindah ke record pertama
    DE.rsPemasok.MoveFirst
End Sub
Private Sub cmdLast_click()
'Pindah kerecord terakhir
    DE.rsPemasok.MoveLast
End Sub

Private Sub cmdNext_Click()
'Pindah satu record ke belakang/ kearah record terakhir
    With DE.rsPemasok
        .MoveNext
        If .EOF Then
            .MoveFirst
        End If
    End With
End Sub

Private Sub cmdPrevious_click()
'Pindah satu record ke depan/ kearah record pertama
    With DE.rsPemasok
        .MovePrevious
        If .BOF Then
            .MoveLast
        End If
    End With
End Sub

