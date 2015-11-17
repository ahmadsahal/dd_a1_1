VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Menghitung selisih jam"
   ClientHeight    =   2400
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   3504
   LinkTopic       =   "Form1"
   ScaleHeight     =   2400
   ScaleWidth      =   3504
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   288
      Left            =   1320
      TabIndex        =   7
      Top             =   1800
      Width           =   1332
   End
   Begin VB.CommandButton Command1 
      Caption         =   "HITUNG"
      Height          =   372
      Left            =   240
      TabIndex        =   6
      Top             =   1200
      Width           =   972
   End
   Begin VB.TextBox Text2 
      Height          =   288
      Left            =   1320
      TabIndex        =   3
      Top             =   720
      Width           =   1332
   End
   Begin VB.TextBox Text1 
      Height          =   288
      Left            =   1320
      TabIndex        =   0
      Top             =   240
      Width           =   1332
   End
   Begin VB.Label Label6 
      Caption         =   "jam"
      Height          =   252
      Left            =   2760
      TabIndex        =   9
      Top             =   1800
      Width           =   372
   End
   Begin VB.Label Label5 
      Caption         =   "Lama kerja"
      Height          =   252
      Left            =   240
      TabIndex        =   8
      Top             =   1800
      Width           =   972
   End
   Begin VB.Label Label4 
      Caption         =   "hh:mm"
      Height          =   252
      Left            =   2760
      TabIndex        =   5
      Top             =   720
      Width           =   612
   End
   Begin VB.Label Label3 
      Caption         =   "Selesai kerja"
      Height          =   252
      Left            =   240
      TabIndex        =   4
      Top             =   720
      Width           =   972
   End
   Begin VB.Label Label2 
      Caption         =   "hh:mm"
      Height          =   252
      Left            =   2760
      TabIndex        =   2
      Top             =   240
      Width           =   612
   End
   Begin VB.Label Label1 
      Caption         =   "Mulai kerja"
      Height          =   252
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   972
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Text3 = (TimeValue(Text2) - TimeValue(Text1)) * 24
End Sub
