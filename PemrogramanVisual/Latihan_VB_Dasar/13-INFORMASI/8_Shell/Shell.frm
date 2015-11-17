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
'Shell: Menjalankan file EXE.
'Bentuk umum: Shell "NamaPath"

Private Sub Form_Load()
    Dim NamaFile As String
    NamaFile = InputBox("Nama file exe akan dijalankan (contoh: c:\programku.exe):", "Menjalankan file EXE")
    Shell NamaFile, vbMaximizedFocus
End Sub
