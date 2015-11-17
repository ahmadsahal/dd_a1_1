VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Kotak"
   ClientHeight    =   2160
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   3240
   LinkTopic       =   "Form1"
   ScaleHeight     =   2160
   ScaleWidth      =   3240
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Paint()
    Line (450, 450)-Step(2000, 1000), vbBlack, B
End Sub
