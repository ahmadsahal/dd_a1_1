VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1920
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   3720
   LinkTopic       =   "Form1"
   ScaleHeight     =   1920
   ScaleWidth      =   3720
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   372
      Left            =   2400
      TabIndex        =   1
      Top             =   1080
      Width           =   972
   End
   Begin VB.TextBox Text1 
      Height          =   288
      Left            =   1440
      TabIndex        =   0
      Top             =   480
      Width           =   1932
   End
   Begin VB.Label Label1 
      Caption         =   "Tanggal"
      Height          =   252
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   972
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If IsDate(Text1) Then
        MsgBox ("Benar, yang Anda ketik adalah tanggal!")
    Else
        MsgBox ("Salah, yang Anda ketik bukan tanggal!")
    End If
End Sub

