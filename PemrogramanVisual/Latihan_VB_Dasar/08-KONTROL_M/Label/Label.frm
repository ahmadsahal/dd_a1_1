VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2400
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   3744
   LinkTopic       =   "Form1"
   ScaleHeight     =   2400
   ScaleWidth      =   3744
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   372
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3252
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Label1.FontName = "Arial"
    Label1.FontSize = 14
    Label1.FontBold = True
    Label1.ForeColor = vbRed
    Label1.Alignment = 2
    Label1.Caption = "NILAI RATA-RATA"
End Sub

