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
'FileCopy: Untuk mengcopy file

Private Sub Form_Load()
    Dim Asal, Tujuan As String
    On Error GoTo Salah
    Asal = InputBox("Nama file akan dicopy (contoh: C:\SURAT.DOC):")
    Tujuan = InputBox("Dicopy ke (contoh: D:\SURAT.DOC):")
    FileCopy Asal, Tujuan
    MsgBox "Sukses, file sudah dicopy!"
    End
Salah:
    MsgBox "Error, file tidak dapat dicopy!"
    End
End Sub
