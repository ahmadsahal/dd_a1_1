VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3264
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   3744
   LinkTopic       =   "Form1"
   ScaleHeight     =   3264
   ScaleWidth      =   3744
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
'Chr: Mencetak karakter ASCII (0-255)
'     0-32 tidak dapat dicetak
'Asc: Mencetak kode karakter ASCII
    Print Chr(65)  'A
    Print Asc("A") '65
    Print
    Rem mencetak semua karakter ASCII
    For n = 33 To 255
        Print "Chr("; n; ") = "; Chr(n)
    Next
End Sub

