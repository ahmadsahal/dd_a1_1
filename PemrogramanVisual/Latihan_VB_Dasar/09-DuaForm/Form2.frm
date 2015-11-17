VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   2736
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   3372
   LinkTopic       =   "Form2"
   ScaleHeight     =   2736
   ScaleWidth      =   3372
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   1572
      Left            =   360
      Picture         =   "Form2.frx":0000
      ScaleHeight     =   1524
      ScaleWidth      =   2604
      TabIndex        =   1
      Top             =   360
      Width           =   2652
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Keluar"
      Height          =   372
      Left            =   1080
      TabIndex        =   0
      Top             =   2160
      Width           =   1212
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub
