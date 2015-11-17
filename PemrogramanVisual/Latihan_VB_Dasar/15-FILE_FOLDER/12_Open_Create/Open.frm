VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   ScaleHeight     =   2400
   ScaleWidth      =   3750
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Open: Membuat file

Private Sub Form_Load()
    Dim NamaBarang, Jumlah, HargaSatuan As String
    NamaBarang = "TV"
    Jumlah = "3"
    HargaSatuan = "2000000"
    Open App.Path & "\DATA.DAT" For Append As #1
        Write #1, NamaBarang, Jumlah, HargaSatuan
    Close
    MsgBox "File DATA.DAT sudah dibuat"
    End
End Sub
