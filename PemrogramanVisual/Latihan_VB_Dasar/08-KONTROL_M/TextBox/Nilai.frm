VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2796
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   3948
   LinkTopic       =   "Form1"
   ScaleHeight     =   2796
   ScaleWidth      =   3948
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TextRata 
      Height          =   288
      Left            =   1560
      TabIndex        =   7
      Top             =   1680
      Width           =   1332
   End
   Begin VB.TextBox TextPraktek 
      Height          =   288
      Left            =   1560
      TabIndex        =   6
      Top             =   1080
      Width           =   1332
   End
   Begin VB.TextBox TextTeori 
      Height          =   288
      Left            =   1560
      TabIndex        =   5
      Top             =   720
      Width           =   1332
   End
   Begin VB.TextBox TextNama 
      Height          =   288
      Left            =   1560
      TabIndex        =   4
      Top             =   360
      Width           =   2052
   End
   Begin VB.Label Label4 
      Caption         =   "Nilai Rata-Rata"
      Height          =   252
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   1212
   End
   Begin VB.Label Label3 
      Caption         =   "Nilai Praktek"
      Height          =   252
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   1092
   End
   Begin VB.Label Label2 
      Caption         =   "Nilai Teori"
      Height          =   252
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "Nama SIswa"
      Height          =   252
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1092
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub TextTeori_Change()
    TextRata = (Val(TextTeori) + Val(TextPraktek)) / 2
End Sub

Private Sub TextPraktek_Change()
    TextRata = (Val(TextTeori) + Val(TextPraktek)) / 2
End Sub

Private Sub TextRata_Change()
    TextRata = (Val(TextTeori) + Val(TextPraktek)) / 2
End Sub

