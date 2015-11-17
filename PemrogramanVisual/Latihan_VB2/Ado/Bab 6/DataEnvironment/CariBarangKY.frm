VERSION 5.00
Begin VB.Form frmCariBarang 
   Caption         =   "Cari Barang"
   ClientHeight    =   1635
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   5115
   LinkTopic       =   "Form1"
   ScaleHeight     =   1635
   ScaleWidth      =   5115
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtJumlahBarang 
      DataField       =   "Jumlah Barang"
      DataMember      =   "Barang"
      DataSource      =   "DE"
      Height          =   285
      Left            =   1440
      TabIndex        =   7
      Top             =   1185
      Width           =   660
   End
   Begin VB.TextBox txtHargaBarang 
      DataField       =   "Harga Barang"
      DataMember      =   "Barang"
      DataSource      =   "DE"
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Top             =   810
      Width           =   1320
   End
   Begin VB.TextBox txtNamaBarang 
      DataField       =   "Nama Barang"
      DataMember      =   "Barang"
      DataSource      =   "DE"
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Top             =   420
      Width           =   3375
   End
   Begin VB.TextBox txtKodeBarang 
      DataField       =   "Kode Barang"
      DataMember      =   "Barang"
      DataSource      =   "DE"
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   45
      Width           =   990
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Jumlah Barang:"
      Height          =   255
      Index           =   3
      Left            =   195
      TabIndex        =   6
      Top             =   1230
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Harga Barang:"
      Height          =   255
      Index           =   2
      Left            =   195
      TabIndex        =   4
      Top             =   855
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Nama Barang:"
      Height          =   255
      Index           =   1
      Left            =   195
      TabIndex        =   2
      Top             =   465
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Kode Barang:"
      Height          =   255
      Index           =   0
      Left            =   195
      TabIndex        =   0
      Top             =   90
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

Private Sub mnuCariBarang_Click()
         
    strKode = InputBox("Masukkan kode Barang yang akan anda cari." & vbCrLf & _
        "", "Cari Data Barang")

    'Cari berdasarkan Kode Pemasok

    Dim vntBookMark As Variant
    
        strCari = "[Kode Barang] = '" & strKode & "'"
        With DE.rsBarang
            vntBookMark = .Bookmark     'Save pointer to current record
            .MoveFirst
            .Find strCari
            If .EOF Then
                MsgBox "Kode Tidak Ada", vbExclamation, "Perhatian"
                .Bookmark = vntBookMark 'Return to previous record
            End If
        End With
    
End Sub

