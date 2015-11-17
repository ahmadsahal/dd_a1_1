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
'Input: Membuka file untuk dibaca

Private Sub Form_Activate()
    Dim NamaBarang, Jumlah, HargaSatuan As String
    Dim TotalHarga As Currency
    On Error GoTo Salah
   If Dir(App.Path & "\Data.dat") = "" Then
   Open App.Path & "\Data.dat" For Output As #1
   
   Write #1, NamaBarang, Jumlah, HargaSatuan
   Close #1
   End If

    Exit Sub
Salah:
    MsgBox "File belum dibuat, buat dahulu dengan Open!"
    End
End Sub

