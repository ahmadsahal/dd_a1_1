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
    Print "2 + 3 ="; 2 + 3
    Print "5 - 3 ="; 5 - 3
    Print "2 x 3 = "; 2 * 3
    Print "11 / 2 ="; 11 / 2
    Print "11 \ 2 ="; 11 \ 2
    Print "2 ^ 3 ="; 2 ^ 3
    Print "10 Mod 3 ="; 10 Mod 3
    Print "Mico & Pardosi ="; "Mico" & "Pardosi"
End Sub
