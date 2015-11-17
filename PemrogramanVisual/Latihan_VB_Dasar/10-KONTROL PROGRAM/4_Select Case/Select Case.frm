VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2400
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   3936
   LinkTopic       =   "Form1"
   ScaleHeight     =   2400
   ScaleWidth      =   3936
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TextKet 
      Height          =   288
      Left            =   1920
      TabIndex        =   6
      Top             =   1560
      Width           =   1332
   End
   Begin VB.TextBox TextNDH 
      Height          =   288
      Left            =   1920
      TabIndex        =   2
      Top             =   1200
      Width           =   1332
   End
   Begin VB.TextBox TextNDA 
      Height          =   288
      Left            =   1920
      TabIndex        =   1
      Top             =   840
      Width           =   1332
   End
   Begin VB.TextBox TextNama 
      Height          =   288
      Left            =   1920
      TabIndex        =   0
      Top             =   480
      Width           =   1692
   End
   Begin VB.Label Label4 
      Caption         =   "Keterangan"
      Height          =   252
      Left            =   240
      TabIndex        =   7
      Top             =   1560
      Width           =   1572
   End
   Begin VB.Label Label3 
      Caption         =   "Nama siswa"
      Height          =   252
      Left            =   240
      TabIndex        =   5
      Top             =   480
      Width           =   972
   End
   Begin VB.Label Label2 
      Caption         =   "Nilai dengan huruf"
      Height          =   252
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   1572
   End
   Begin VB.Label Label1 
      Caption         =   "Nilai dengan angka"
      Height          =   252
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   1572
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub TextNDA_Change()
    Select Case Val(TextNDA)
        Case 90 To 100
            TextNDH = "A"
            TextKet = "SANGAT BAIK"
        Case 70 To 89
            TextNDH = "B"
            TextKet = "BAIK"
        Case 60 To 69
            TextNDH = "C"
            TextKet = "CUKUP"
        Case 0 To 59
            TextNDH = "D"
            TextKet = "KURANG"
    End Select
End Sub

