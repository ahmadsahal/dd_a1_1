VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Melihat ukuran file"
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
'FileLen: Untuk menampilkan tanggal dan jam file
'    Bentuk umum:
'FileLen(NamaFile)

Private Sub Form_Load()
    Dim NamaFile As String
    On Error GoTo Salah
    NamaFile = InputBox("Ketik nama sebuah file (contoh: C:\NASKAH.DOC)")
    MsgBox ("Ukuran file: " & FileLen(NamaFile) & " byte")
    End
Salah:
    MsgBox "File tidak ditemukan!"
    End
End Sub

