VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3612
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   3744
   LinkTopic       =   "Form1"
   ScaleHeight     =   3612
   ScaleWidth      =   3744
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Paint()
    Const PI = 3.14159265358979
    FillStyle = vbFSSolid
    FillColor = vbBlue
    Circle (ScaleWidth / 2 + 200, ScaleHeight / 2 - 200), _
    1500, vbBlack, -(PI * 2), -(PI / 2)
    FillColor = vbCyan
    Circle (ScaleWidth / 2, ScaleHeight / 2), _
    1500, vbBlack, -(PI / 2), -(PI * 2)
End Sub
