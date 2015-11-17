VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2400
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   3744
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   10.2
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
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
    Print FormatNumber(10000)
    'Akan menghasilkan 10.000,00 (2 desimal)
    Print FormatNumber(10000, 1)
    'Akan menghasilkan 10.000,00 (1 desimal)
    Print FormatNumber(10000, 0)
    'Akan menghasilkan 10.000 (tanpa desimal)

    Print FormatNumber("10000")
    'Akan menghasilkan 10.000,00
    Print FormatNumber("10000", 0)
End Sub


