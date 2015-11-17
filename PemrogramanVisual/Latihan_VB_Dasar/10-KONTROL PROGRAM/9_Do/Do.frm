VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Do"
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
'Do: Perintah untuk membuat loop
'(pengulangan)

Private Sub Form_Activate()
    Dim i As Integer
  
    Do
        i = i + 1
        Print i
    Loop Until i = 10

End Sub

