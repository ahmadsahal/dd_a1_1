VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Pset"
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
Private Sub Form_Paint()
    For i = 1 To 1000
        DrawWidth = Rnd * 10 + 1
        PSet (Rnd * ScaleWidth, Rnd * ScaleHeight), _
            RGB(Rnd * 255, Rnd * 255, Rnd * 255)
    Next
End Sub
