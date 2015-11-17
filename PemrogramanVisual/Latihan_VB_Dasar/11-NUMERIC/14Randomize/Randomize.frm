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
'Randomize: Agar bilangan acak berbeda
'setiap kali program dijalankan
    Randomize
    For i = 1 To 10
        Print Rnd(i)
    Next i
End Sub

