VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Melihat tanggal file"
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
'FileDateTime: Untuk menampilkan tanggal dan jam file
'    Bentuk umum:
'FileDateTime(NamaFile)

Private Sub Form_Load()
    Dim NamaFile As String
    On Error GoTo Salah
    NamaFile = InputBox("Ketik nama sebuah file (contoh: C:\DATA.DAT)")
    MsgBox ("Tanggal dan jam file: " & FileDateTime(NamaFile))
    End
Salah:
    MsgBox "File tidak ditemukan!"
    End
End Sub
