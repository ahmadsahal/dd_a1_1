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
    Form1.FontSize = 12
    Judul = "PT SURYA KENCANA"
    CurrentX = (ScaleWidth - TextWidth(Judul)) / 2
    Print Judul
    Form1.FontSize = 8
    Print
    Print Tab(5); "NO"; Tab(20); "NAMA"
    Print Tab(5); "URUT"; Tab(20); "KARYAWAN"
    Print
End Sub

