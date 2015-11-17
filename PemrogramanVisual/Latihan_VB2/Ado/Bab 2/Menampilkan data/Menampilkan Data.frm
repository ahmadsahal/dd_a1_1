VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form28 
   Caption         =   "Menampilkan, mencari dan mengurutkan Data"
   ClientHeight    =   3510
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7860
   LinkTopic       =   "Form1"
   ScaleHeight     =   3510
   ScaleWidth      =   7860
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid dgdVisit 
      Bindings        =   "Menampilkan Data.frx":0000
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   5318
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DataMember      =   "Pemasok"
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "Kode Pemasok"
         Caption         =   "Kode Pemasok"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Nama Pemasok"
         Caption         =   "Nama Pemasok"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Alamat Pemasok"
         Caption         =   "Alamat Pemasok"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "No Telepon"
         Caption         =   "No Telepon"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1244.976
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   3179.906
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1335.118
         EndProperty
      EndProperty
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "Informasi Data Pemasok"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   1
      Top             =   0
      Width           =   3255
   End
   Begin VB.Menu mnuKeluar 
      Caption         =   "Keluar"
   End
   Begin VB.Menu mnuCari 
      Caption         =   "Cari"
      Begin VB.Menu cariKode 
         Caption         =   "Kode Pemasok"
      End
      Begin VB.Menu mnuCariNama 
         Caption         =   "Nama Pemasok"
      End
   End
   Begin VB.Menu mnuurut 
      Caption         =   "Urutkan"
      Begin VB.Menu mnuUrutKode 
         Caption         =   "berdasarkan Kode Pemasok"
      End
      Begin VB.Menu mnuurutnama 
         Caption         =   "berdasarkan Nama Pemasok"
      End
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
'Purpose        : menampilkan data menggunakan objek visual Data Environment
'Folder         : ADO/BAB2/Menampilkan Data

Option Explicit
    
Private Sub cariKode_Click()
'Cari berdasarkan nama Pemasok
    Dim strKode     As String
    Dim vntBookmark As Variant
    
    On Error GoTo HandleError
    strKode = InputBox("Tulis Kode Pemasok Yang anda Cari." & vbCrLf & _
        "Kode Pemasok sudah benar", "Cari Kode Pemasok")
    With DE.rsPemasok
        vntBookmark = .Bookmark
        .Find "[Kode Pemasok] Like '" & strKode & "*'"
        If .EOF Then
            MsgBox "Data Tidak ditemukan'" & strKode & "'", vbInformation, "Cari Kode Pemasok"
            .Bookmark = vntBookmark
        End If
    End With
    
CariKode_Click_Exit:
    Exit Sub
    
HandleError:
    MsgBox "Proses tidak dapat dilakukan.", vbInformation, "Perhatian"
    On Error GoTo 0
End Sub

Private Sub mnuCariNama_Click()
 'Cari berdasarkan nama Pemasok
    Dim strNama     As String
    Dim vntBookmark As Variant
    
    On Error GoTo HandleError
    strNama = InputBox("Tulis Nama Pemasok Yang anda Cari." & vbCrLf & _
        "Nama Pemasok sudah benar", "Cari Nama Pemasok")
    With DE.rsPemasok
        vntBookmark = .Bookmark
        .Find "[Nama Pemasok] Like '" & strNama & "*'"
        If .EOF Then
            MsgBox "Data Tidak ditemukan'" & strNama & "'", vbInformation, "Cari nama Pemasok"
            .Bookmark = vntBookmark
        End If
    End With
    
mnuCariNama_Click_Exit:
    Exit Sub
    
HandleError:
    MsgBox "Proses tidak dapat dilakukan.", vbInformation, "Perhatian"
    On Error GoTo 0
End Sub

Private Sub mnuKeluar_Click()
'Keluar dari proyek
    Unload Me
End Sub
Private Sub mnuUrutKode_Click()
    DE.rsPemasok.Sort = "[Kode Pemasok]"
End Sub

Private Sub mnuurutnama_Click()
    DE.rsPemasok.Sort = "[Nama Pemasok]"
End Sub
