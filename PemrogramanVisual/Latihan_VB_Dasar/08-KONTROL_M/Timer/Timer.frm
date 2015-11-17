VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1524
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   2880
   LinkTopic       =   "Form1"
   ScaleHeight     =   1524
   ScaleWidth      =   2880
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Left            =   1200
      Top             =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Timer1.Enabled = True
    Timer1.Interval = 1000
End Sub

Private Sub Timer1_Timer()
    Cls
    Form1.FontSize = 10
    Print Time$
End Sub

