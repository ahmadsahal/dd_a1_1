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
   Begin VB.TextBox TextKeterangan 
      Height          =   288
      Left            =   1560
      TabIndex        =   2
      Top             =   1440
      Width           =   1812
   End
   Begin VB.TextBox TextNilai 
      Height          =   288
      Left            =   1560
      TabIndex        =   1
      Top             =   720
      Width           =   1812
   End
   Begin VB.TextBox TextNama 
      Height          =   288
      Left            =   1560
      TabIndex        =   0
      Top             =   360
      Width           =   1812
   End
   Begin VB.Label Label4 
      Caption         =   "Nilai harus diketik dalam huruf besar, mis: C"
      Height          =   252
      Left            =   240
      TabIndex        =   6
      Top             =   1920
      Width           =   3252
   End
   Begin VB.Label Label3 
      Caption         =   "Keterangan"
      Height          =   252
      Left            =   240
      TabIndex        =   5
      Top             =   1440
      Width           =   972
   End
   Begin VB.Label Label2 
      Caption         =   "Nilai (A-F)"
      Height          =   252
      Left            =   240
      TabIndex        =   4
      Top             =   720
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "Nama siswa"
      Height          =   252
      Left            =   240
      TabIndex        =   3
      Top             =   360
      Width           =   972
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub TextKaterangan_Change()

End Sub

Private Sub TextNilai_Change()
    If TextNilai = "A" Or TextNilai = "B" Or TextNilai = "C" Then
        TextKeterangan = "LULUS"
    Else
        TextKeterangan = "GAGAL"
    End If
End Sub

