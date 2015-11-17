VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "FillStyle"
   ClientHeight    =   2316
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   3348
   LinkTopic       =   "Form1"
   ScaleHeight     =   2316
   ScaleWidth      =   3348
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Paint()
    FillStyle = 7
    Line (100, 100)-Step(2000, 2000), vbBlack, B
End Sub
