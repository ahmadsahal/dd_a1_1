VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "BACA DATA"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   ScaleHeight     =   3375
   ScaleWidth      =   4935
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView ListView1 
      Height          =   2532
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4452
      _ExtentX        =   7858
      _ExtentY        =   4471
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   252
      Left            =   240
      TabIndex        =   1
      Top             =   2880
      Width           =   2172
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Dim LI As ListItem
    Dim Nama, Kode, Gaji As String
    Dim TotalGaji As Currency
    
    ListView1.View = lvwReport
    ListView1.Sorted = True 'Menyortir data

    'Membuat Judul Kolom (ColumnHeaders) serta mengatur lebar.
    ListView1.ColumnHeaders.Add , , "NAMA", ListView1.Width / 3
    ListView1.ColumnHeaders.Add , , "GOLONGAN", ListView1.Width / 3
    ListView1.ColumnHeaders.Add , , "GAJI", ListView1.Width / 3
    On Error GoTo Salah
    Open App.Path & "\GAJI.DAT" For Input As #1
    Do Until EOF(1)
        Input #1, Nama, Kode, Gaji
        'Masukkan item dan sub item ke dalam list.
        Set LI = ListView1.ListItems.Add(, , Nama)
        LI.SubItems(1) = Kode
        LI.SubItems(2) = Gaji
        TotalGaji = TotalGaji + Gaji
    Loop
    Close #1
    Label1.Caption = "Total Gaji = " + Format(TotalGaji, "Currency")
    Exit Sub
Salah:
    MsgBox "File C:\VB6\GAJI.DAT tidak ditemukan!"
    End
End Sub

'Jika garis-garis tipis seperti dalam Exel
'ingin dimasukkan, tambahkan perintah ini:
'ListView1.GridLines = True di bawah perintah
'ListView1.Sorted = True

