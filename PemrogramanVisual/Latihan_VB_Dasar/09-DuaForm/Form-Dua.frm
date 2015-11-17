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
   Begin VB.CommandButton Command2 
      Caption         =   "Keluar"
      Height          =   372
      Left            =   2160
      TabIndex        =   1
      Top             =   840
      Width           =   1212
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Buka Form2"
      Height          =   372
      Left            =   2160
      TabIndex        =   0
      Top             =   240
      Width           =   1212
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Form2.Show
End Sub

Private Sub Command2_Click()
    End
End Sub






