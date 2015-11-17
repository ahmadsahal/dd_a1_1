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
'Name: Untuk mengubah nama file atau folder
'          Bentuk umum:
' Name "NamaLama" As "NamaBaru"

Private Sub Form_Load()
    Dim NamaLama, NamaBaru As String
    On Error GoTo Salah
    NamaLama = InputBox("Nama file atau folder akan diubah (contoh: C:\SURAT.DOC):")
    NamaBaru = InputBox("Diubah menjadi (contoh: C:\NASKAH.DOC):")
    Name NamaLama As NamaBaru
    MsgBox "Sukses, nama file sudah diubah!"
    End
Salah:
    MsgBox "File tidak ditemukan!"
    End
End Sub
