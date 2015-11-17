VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1536
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   3672
   LinkTopic       =   "Form1"
   ScaleHeight     =   1536
   ScaleWidth      =   3672
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   492
      Left            =   2040
      TabIndex        =   0
      Top             =   120
      Width           =   1572
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Date: Untuk menampilkan tanggal hari ini

Private Sub Form_Load()
    Label1.FontSize = 14
    Label1.Caption = Date
End Sub

