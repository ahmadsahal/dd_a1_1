VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "FillStyle"
   ClientHeight    =   2448
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   3528
   LinkTopic       =   "Form1"
   ScaleHeight     =   2448
   ScaleWidth      =   3528
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FillStyle: Untuk memasukkan pola
'FillStyle = 1 sampai 7

Private Sub Form_Paint()
    FillStyle = 7
    Circle (1500, 1100), 1000
End Sub
