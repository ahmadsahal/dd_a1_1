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
Private Sub Form_Activate()
    Cls
    GoSub JUDUL
    Print "SONI", 32, "SURABAYA"
    Print "AFNI", 27, "MALANG"
    GoTo Akhir
JUDUL:
    Print "DAFTAR ANGGOTA"
    Print
    Print "NAMA", "UMUR", "ALAMAT"
    Print
    Return
Akhir:
End Sub

