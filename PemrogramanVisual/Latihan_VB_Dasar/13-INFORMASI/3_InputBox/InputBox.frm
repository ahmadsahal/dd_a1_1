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
'InputBox: Meminta masukan dari si pemakai

Private Sub Form_Load()
    Dim NamaFolder As String
    NamaFolder = InputBox("Nama folder akan dibuat (contoh: C:\PROGRAM)", "Membuat folder")
    On Error GoTo Salah
    MkDir (NamaFolder)
    MsgBox ("Folder sukses dibuat!")
    End
Salah:
    MsgBox ("Folder gagal dibuat!")
    End
End Sub
