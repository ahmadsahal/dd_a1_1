VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Menjumlahkan 2 bilangan"
   ClientHeight    =   2280
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   3744
   LinkTopic       =   "Form1"
   ScaleHeight     =   2280
   ScaleWidth      =   3744
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "ULANG"
      Height          =   372
      Left            =   2400
      TabIndex        =   6
      Top             =   1680
      Width           =   972
   End
   Begin VB.TextBox Text3 
      Height          =   288
      Left            =   1920
      TabIndex        =   2
      Top             =   1200
      Width           =   1452
   End
   Begin VB.TextBox Text2 
      Height          =   288
      Left            =   1920
      TabIndex        =   1
      Top             =   600
      Width           =   1452
   End
   Begin VB.TextBox Text1 
      Height          =   288
      Left            =   1920
      TabIndex        =   0
      Top             =   240
      Width           =   1452
   End
   Begin VB.Label Label3 
      Caption         =   "Jumlah"
      Height          =   252
      Left            =   240
      TabIndex        =   5
      Top             =   1200
      Width           =   732
   End
   Begin VB.Label Label2 
      Caption         =   "Bilangan kedua"
      Height          =   252
      Left            =   240
      TabIndex        =   4
      Top             =   600
      Width           =   1212
   End
   Begin VB.Label Label1 
      Caption         =   "Bilangan  pertama"
      Height          =   252
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   1452
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Call: Perintah untuk memanggil sub program

Function Jumlah() As Currency
    Jumlah = Val(Text1) + Val(Text2)
End Function

Private Sub Command1_Click()
    Kosongkan 'Atau Call Kosongkan
End Sub

Private Sub Text1_Change()
    Text3 = Jumlah
End Sub

Private Sub Text2_Change()
    Text3 = Jumlah
End Sub

Private Sub Kosongkan()
    Text1 = ""
    Text2 = ""
    Text3 = ""
    Text1.SetFocus
End Sub

