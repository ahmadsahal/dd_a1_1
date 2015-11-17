VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Hari dan Bulan"
   ClientHeight    =   2100
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   3276
   LinkTopic       =   "Form1"
   ScaleHeight     =   2100
   ScaleWidth      =   3276
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List2 
      Height          =   1392
      ItemData        =   "ListBox2.frx":0000
      Left            =   1680
      List            =   "ListBox2.frx":0002
      TabIndex        =   1
      Top             =   360
      Width           =   1092
   End
   Begin VB.ListBox List1 
      Height          =   1392
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   1092
   End
   Begin VB.Label Label2 
      Caption         =   "BULAN"
      Height          =   252
      Left            =   1680
      TabIndex        =   3
      Top             =   120
      Width           =   1092
   End
   Begin VB.Label Label1 
      Caption         =   "HARI"
      Height          =   252
      Left            =   480
      TabIndex        =   2
      Top             =   120
      Width           =   972
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    For i = 1 To 7
        List1.AddItem WeekdayName(i)
    Next
    
    For i = 1 To 12
        List2.AddItem MonthName(i)
    Next
End Sub

