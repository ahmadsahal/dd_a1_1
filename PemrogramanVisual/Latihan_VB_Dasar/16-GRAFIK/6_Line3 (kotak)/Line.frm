VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Kotak"
   ClientHeight    =   2268
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   3336
   LinkTopic       =   "Form1"
   ScaleHeight     =   2268
   ScaleWidth      =   3336
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Paint()
    Line (450, 450)-Step(2000, 1000), vbRed, BF
End Sub
