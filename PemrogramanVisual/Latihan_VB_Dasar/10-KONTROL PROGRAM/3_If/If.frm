VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2580
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   3648
   LinkTopic       =   "Form1"
   ScaleHeight     =   2580
   ScaleWidth      =   3648
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TextKeterangan 
      Height          =   288
      Left            =   1560
      TabIndex        =   8
      Top             =   1920
      Width           =   1812
   End
   Begin VB.TextBox TextRata 
      Height          =   288
      Left            =   1560
      TabIndex        =   7
      Top             =   1560
      Width           =   1812
   End
   Begin VB.TextBox TextPraktek 
      Height          =   288
      Left            =   1560
      TabIndex        =   6
      Top             =   960
      Width           =   1812
   End
   Begin VB.TextBox TextTeori 
      Height          =   288
      Left            =   1560
      TabIndex        =   5
      Top             =   600
      Width           =   1812
   End
   Begin VB.TextBox TextNama 
      Height          =   288
      Left            =   1560
      TabIndex        =   3
      Top             =   240
      Width           =   1812
   End
   Begin VB.Label Label5 
      Caption         =   "Keterangan"
      Height          =   252
      Left            =   240
      TabIndex        =   9
      Top             =   1920
      Width           =   1092
   End
   Begin VB.Label Label4 
      Caption         =   "Nama Siswa"
      Height          =   252
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   1212
   End
   Begin VB.Label Label3 
      Caption         =   "Nilai Rata-rata"
      Height          =   252
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   1092
   End
   Begin VB.Label Label2 
      Caption         =   "Nilai Praktek"
      Height          =   252
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "Nilai Teori"
      Height          =   252
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   852
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub TextTeori_Change()
    TextRata = (Val(TextTeori) + (Val(TextPraktek))) / 2
    If Val(TextRata) >= 60 Then
        TextKeterangan = "LULUS"
    Else
        TextKeterangan = "GAGAL"
    End If
End Sub

Private Sub TextPraktek_Change()
    TextRata = (Val(TextTeori) + (Val(TextPraktek))) / 2
    If Val(TextRata) >= 60 Then
        TextKeterangan = "LULUS"
    Else
        TextKeterangan = "GAGAL"
    End If
End Sub

