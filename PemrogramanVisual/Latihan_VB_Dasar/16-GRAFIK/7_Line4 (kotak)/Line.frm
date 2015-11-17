VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "LIne"
   ClientHeight    =   2148
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   3036
   LinkTopic       =   "Form1"
   ScaleHeight     =   2148
   ScaleWidth      =   3036
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Paint()
    Line (500, 500)-Step(2000, 1000), RGB(64, 64, 64), BF
    Line (450, 450)-Step(2000, 1000), vbYellow, BF
    Line (450, 450)-Step(2000, 1000), vbBlack, B
End Sub
