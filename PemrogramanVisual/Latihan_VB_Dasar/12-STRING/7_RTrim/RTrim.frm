VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4485
   LinkTopic       =   "Form1"
   ScaleHeight     =   2400
   ScaleWidth      =   4485
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
    Print "1. Panjang text sebelum pakai Rtrim() adalah=" & Len("Tom     ") & " character"
    Print "2. Hasil text dengan menggunakan Rtrim() adalah " & RTrim("Tom     ") ' "Tom"
    Print "3. Panjang text sesudah pakai Rtrim() adalah=" & RTrim(Len("Tom     ")) & " character"
End Sub

 
