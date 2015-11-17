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
Private Sub Form_Activate()
    Dim TglMasuk, TglKeluar As String
    Dim Selisih As Integer
    TglMasuk = "02/11/03"
    TglKeluar = "12/12/03"
    Print "Tanggal Masuk = "; TglMasuk
    Print "Tanggal Keluar = "; TglKeluar
    Selisih = DateDiff("d", TglMasuk, TglKeluar)
    Print "Selisih tanggal = "; Selisih; " hari"
End Sub

