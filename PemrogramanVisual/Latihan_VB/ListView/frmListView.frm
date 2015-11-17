VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmListView 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "List View For Database"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCari 
      Caption         =   "&Cari"
      Height          =   495
      Left            =   5160
      TabIndex        =   3
      Top             =   4920
      Width           =   1575
   End
   Begin VB.TextBox txtNama 
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   5040
      Width           =   1935
   End
   Begin MSComctlLib.ListView lvwData 
      Height          =   4695
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   8281
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
      Caption         =   "&Nama"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   5040
      Width           =   735
   End
End
Attribute VB_Name = "frmListView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

   Dim myDb As Database
   Dim myRs As Recordset
   Dim myList As ListItem

   Private Sub BuatKolom()
      With lvwData.ColumnHeaders
        .Add , , "No", lvwData.Width * 1 / 10
        .Add , , "Nama", lvwData.Width * 1 / 3
        .Add , , "Alamat", lvwData.Width * 2 / 3
      End With
      lvwData.View = lvwReport
   End Sub
 
   Private Sub BukaTabel():
   Dim i
   On Error Resume Next
       While Not myRs.EOF
       i = i + 1
          Set myList = lvwData.ListItems.Add _
                       (, , CStr(i))
          myList.SubItems(1) = myRs!Nama
          myList.SubItems(2) = myRs!Alamat
          myRs.MoveNext
       Wend
   End Sub

   Private Sub Form_Load()
        Set myDb = DBEngine.Workspaces(0) _
        .OpenDatabase(App.Path & "\data.mdb")
        Set myRs = myDb.OpenRecordset _
          ("tblTemanKu")
         Call BuatKolom
         Call BukaTabel
        myRs.Close
        myDb.Close
    End Sub


Private Sub cmdCari_Click()
    Dim mySQL$, i
    i = i + 1
    'Menentukan alamat database
    Set myDb = DBEngine.Workspaces(0) _
        .OpenDatabase(App.Path & "\data.mdb")
    'kalimat SQL
    mySQL$ = "SELECT * FROM TBLTEMANKU " & _
           " WHERE NAMA = '" + _
           CStr(Me.txtNama.Text) + "'"
    'tentukan recordset
    Set myRs = myDb.OpenRecordset(mySQL)
    'Jika data ada
    If myRs.RecordCount <> 0 Then
    'kosongkan list view
        Me.lvwData.ListItems.Clear
        'isi list view dengan data
        Set myList = lvwData.ListItems.Add _
                      (, , i)
              myList.SubItems(1) = myRs!Nama
              myList.SubItems(2) = myRs!Alamat
    Else
        MsgBox "Data tidak ketemu"
    End If
End Sub


