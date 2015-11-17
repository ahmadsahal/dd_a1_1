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
   Begin VB.OptionButton Option3 
      Caption         =   "Merah"
      Height          =   372
      Left            =   360
      TabIndex        =   3
      Top             =   1080
      Width           =   972
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Biru"
      Height          =   372
      Left            =   360
      TabIndex        =   2
      Top             =   720
      Width           =   972
   End
   Begin VB.OptionButton Option1 
      Caption         =   "HItam"
      Height          =   372
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "Microsft Visual Basic"
      Height          =   252
      Left            =   1560
      TabIndex        =   0
      Top             =   720
      Width           =   2052
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Label1.FontBold = True
End Sub

Private Sub Option1_Click()
    Label1.ForeColor = vbBlack
End Sub

Private Sub Option2_Click()
    Label1.ForeColor = vbBlue
End Sub

Private Sub Option3_Click()
    Label1.ForeColor = vbRed
End Sub
