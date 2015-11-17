VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "DrawWidth"
   ClientHeight    =   2412
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   3300
   LinkTopic       =   "Form1"
   ScaleHeight     =   2412
   ScaleWidth      =   3300
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'DrawWidth: Untuk membuat border

Private Sub Form_Paint()
    DrawWidth = 5   ' Border setebal 5 pixel
    FillStyle = vbFSSolid
    FillColor = vbRed
    Circle (1200, 1200), 1000, vbGreen
End Sub
