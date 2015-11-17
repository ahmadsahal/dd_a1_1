VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1428
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   3816
   LinkTopic       =   "Form1"
   ScaleHeight     =   1428
   ScaleWidth      =   3816
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   372
      Left            =   2400
      TabIndex        =   0
      Top             =   120
      Width           =   1212
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Time: Untuk menampilkan jam sekarang

Private Sub Form_Load()
    Label1.FontSize = 14
    Label1.Caption = Time$
End Sub
