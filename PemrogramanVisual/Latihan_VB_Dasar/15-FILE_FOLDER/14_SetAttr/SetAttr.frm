VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2400
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   3744
   LinkTopic       =   "Form1"
   ScaleHeight     =   2400
   ScaleWidth      =   3744
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'SetAttr: Untuk mengubah atribut file

Private Sub Form_Load()
    Dim NamaFile As String
    NamaFile = InputBox("Ketik nama file: ")
    SetAttr NamaFile, vbHidden  ' Menyembunyikan File
    MsgBox "File sudah disembunyikan!"
    SetAttr NamaFile, vbArchive ' Wee, muncul kembali
    MsgBox "File dinormalkan kembali!"
    End
 End Sub
