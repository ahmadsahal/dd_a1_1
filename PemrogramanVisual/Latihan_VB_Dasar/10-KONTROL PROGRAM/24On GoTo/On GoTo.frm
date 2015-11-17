VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2244
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   3744
   LinkTopic       =   "Form1"
   ScaleHeight     =   2244
   ScaleWidth      =   3744
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
    Dim Pilihan As Byte
    Pilihan = InputBox("1=Tambah, 2=Kurang, 3=Keluar. Pilihan Anda[1..3]: ")
    On Pilihan GoTo Satu, Dua, Tiga

Satu:
    Print "TAMBAH": Print
    Print "10 + 20 = "; 20 + 10
    GoTo Akhir

Dua:
    Print "KURANG": Print
    Print "30-10"; 30 - 10
    GoTo Akhir
    
Tiga: End

Akhir:

End Sub

