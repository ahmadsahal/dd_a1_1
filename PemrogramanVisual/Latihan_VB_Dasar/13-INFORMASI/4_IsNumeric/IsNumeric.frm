VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1560
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   3744
   LinkTopic       =   "Form1"
   ScaleHeight     =   1560
   ScaleWidth      =   3744
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   372
      Left            =   2400
      TabIndex        =   1
      Top             =   840
      Width           =   972
   End
   Begin VB.TextBox Text1 
      Height          =   288
      Left            =   1680
      TabIndex        =   0
      Top             =   360
      Width           =   1692
   End
   Begin VB.Label Label1 
      Caption         =   "Ketik suatu data"
      Height          =   252
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   1212
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If IsNumeric(Text1) Then
        MsgBox ("Data yang Anda ketik adalah Numeric!")
    Else
        MsgBox ("Data yang Anda ketik adalah String!")
    End If
End Sub
