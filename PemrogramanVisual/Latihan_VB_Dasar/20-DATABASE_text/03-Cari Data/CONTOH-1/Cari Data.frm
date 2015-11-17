VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "CARI DATA"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4425
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4425
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView ListView1 
      Height          =   1452
      Left            =   360
      TabIndex        =   2
      Top             =   960
      Width           =   3732
      _ExtentX        =   6588
      _ExtentY        =   2566
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.TextBox Text1 
      Height          =   288
      Left            =   2640
      TabIndex        =   1
      Top             =   240
      Width           =   1332
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   252
      Left            =   360
      TabIndex        =   3
      Top             =   2640
      Width           =   3732
   End
   Begin VB.Label Label1 
      Caption         =   "Kode golongan akan dicari (1-3)"
      Height          =   252
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2412
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Dim LI As ListItem
    ListView1.View = lvwReport
    ListView1.Sorted = True
    'Membuat Judul Kolom (ColumnHeaders) dan mengatur lebar.
    ListView1.ColumnHeaders.Add , , "NAMA", ListView1.Width / 3
    ListView1.ColumnHeaders.Add , , "GOLONGAN", ListView1.Width / 3
    ListView1.ColumnHeaders.Add , , "GAJI", ListView1.Width / 3
End Sub

Private Sub CariData()
    Dim Nama, Kode, Gaji As String
    Dim Ada As Integer
    Dim TotalGaji As Currency
    ListView1.ListItems.Clear
    Open App.Path & "\GAJI.DAT" For Input As #1
    Do Until EOF(1)
        Input #1, Nama, Kode, Gaji
        If Text1 = Kode Then
            'Masukkan data ke dalam list.
            Set LI = ListView1.ListItems.Add(, , Nama)
            LI.SubItems(1) = Kode
            LI.SubItems(2) = Gaji
            'Jumlahkan gaji
            TotalGaji = TotalGaji + Gaji
            Ada = Ada + 1
        End If
    Loop
    Close #1
    If Ada = 0 Then MsgBox "Kode golongan tersebut tidak ada (1-3)!"
    If Ada <> 0 Then
        Text1 = ""
        Label2.Caption = "Total Gaji = " + Format(TotalGaji, "Currency")
    End If
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If Text1 <> "" And KeyCode = 13 Then CariData
End Sub

