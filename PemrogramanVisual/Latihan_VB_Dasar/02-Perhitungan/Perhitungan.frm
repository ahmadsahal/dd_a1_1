VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3264
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   4200
   LinkTopic       =   "Form1"
   ScaleHeight     =   3264
   ScaleWidth      =   4200
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   288
      Left            =   1680
      TabIndex        =   6
      Top             =   2280
      Width           =   2052
   End
   Begin VB.TextBox Text2 
      Height          =   288
      Left            =   1680
      TabIndex        =   3
      Top             =   1440
      Width           =   2052
   End
   Begin VB.TextBox Text1 
      Height          =   288
      Left            =   1680
      TabIndex        =   2
      Top             =   960
      Width           =   2052
   End
   Begin VB.Label Label4 
      Caption         =   "Total"
      Height          =   252
      Left            =   360
      TabIndex        =   5
      Top             =   2280
      Width           =   492
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "PERHITUNGAN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   1080
      TabIndex        =   4
      Top             =   240
      Width           =   2172
   End
   Begin VB.Label Label3 
      Caption         =   "Harga per unit"
      Height          =   252
      Left            =   360
      TabIndex        =   1
      Top             =   1440
      Width           =   1092
   End
   Begin VB.Label Label2 
      Caption         =   "Jumlah unit"
      Height          =   252
      Left            =   360
      TabIndex        =   0
      Top             =   960
      Width           =   972
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text1_Change()
    Text3 = Val(Text1) * Val(Text2)
End Sub

Private Sub Text2_Change()
    Text3 = Val(Text1) * Val(Text2)
End Sub

