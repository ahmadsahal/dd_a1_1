VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "MsgBox"
   ClientHeight    =   2028
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   3648
   LinkTopic       =   "Form1"
   ScaleHeight     =   2028
   ScaleWidth      =   3648
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   372
      Left            =   2280
      TabIndex        =   4
      Top             =   1320
      Width           =   972
   End
   Begin VB.TextBox Text2 
      Height          =   288
      Left            =   1320
      TabIndex        =   3
      Top             =   840
      Width           =   1932
   End
   Begin VB.TextBox Text1 
      Height          =   288
      Left            =   1320
      TabIndex        =   1
      Top             =   360
      Width           =   1932
   End
   Begin VB.Label Label2 
      Caption         =   "Alamat"
      Height          =   252
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   732
   End
   Begin VB.Label Label1 
      Caption         =   "Nama"
      Height          =   252
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   732
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If Text1 = "" Or Text2 = "" Then
        Beep
        MsgBox ("Isi dulu dong data Anda!")
    Else
        MsgBox ("Nama Anda: " & Text1 & "  Alamat: " & Text2)
    End If
End Sub
