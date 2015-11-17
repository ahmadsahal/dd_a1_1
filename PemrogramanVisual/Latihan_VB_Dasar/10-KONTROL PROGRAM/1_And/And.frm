VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "OPERATOR 'AND'"
   ClientHeight    =   2556
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   3744
   LinkTopic       =   "Form1"
   ScaleHeight     =   2556
   ScaleWidth      =   3744
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TextKeterangan 
      Height          =   288
      Left            =   1560
      TabIndex        =   8
      Top             =   1920
      Width           =   1572
   End
   Begin VB.TextBox TextPraktek 
      Height          =   288
      Left            =   1560
      TabIndex        =   6
      Top             =   1560
      Width           =   1572
   End
   Begin VB.TextBox TextTeori 
      Height          =   288
      Left            =   1560
      TabIndex        =   4
      Top             =   1200
      Width           =   1572
   End
   Begin VB.TextBox TextNama 
      Height          =   288
      Left            =   1560
      TabIndex        =   0
      Top             =   840
      Width           =   1932
   End
   Begin VB.Label Label5 
      Caption         =   "Keterangan"
      Height          =   252
      Left            =   240
      TabIndex        =   7
      Top             =   1920
      Width           =   972
   End
   Begin VB.Label Label4 
      Caption         =   "Nilai Praktek"
      Height          =   252
      Left            =   240
      TabIndex        =   5
      Top             =   1560
      Width           =   972
   End
   Begin VB.Label Label3 
      Caption         =   "Nilai Teori"
      Height          =   252
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   972
   End
   Begin VB.Label Label2 
      Caption         =   "LULUS apabila NILAI TEORI >=60 dan NILAI PRAKTEK >=60, selainnya GAGAL."
      Height          =   492
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   3252
   End
   Begin VB.Label Label1 
      Caption         =   "Nama Praktikan"
      Height          =   252
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   1212
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub TextPraktek_Change()
    If Val(TextTeori) >= 60 And Val(TextPraktek) >= 60 Then
        TextKeterangan = "LULUS"
    Else
        TextKeterangan = "GAGAL"
    End If
End Sub

Private Sub TextTeori_Change()
    If Val(TextTeori) >= 60 And Val(TextPraktek) >= 60 Then
        TextKeterangan = "LULUS"
    Else
        TextKeterangan = "GAGAL"
    End If
End Sub
