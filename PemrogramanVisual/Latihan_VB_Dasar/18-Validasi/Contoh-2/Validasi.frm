VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Validasi"
   ClientHeight    =   2424
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   3588
   LinkTopic       =   "Form1"
   ScaleHeight     =   2424
   ScaleWidth      =   3588
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   372
      Left            =   2400
      TabIndex        =   2
      Top             =   1680
      Width           =   972
   End
   Begin VB.TextBox Text2 
      Height          =   288
      Left            =   1320
      TabIndex        =   1
      Top             =   1200
      Width           =   1932
   End
   Begin VB.TextBox Text1 
      Height          =   288
      Left            =   1320
      TabIndex        =   0
      Top             =   720
      Width           =   1932
   End
   Begin VB.Label Label3 
      Caption         =   "Tombol OK tidak akan hidup apabila salah satu data  tidak diisi."
      Height          =   372
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   3132
   End
   Begin VB.Label Label2 
      Caption         =   "Alamat"
      Height          =   252
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   612
   End
   Begin VB.Label Label1 
      Caption         =   "Nama"
      Height          =   252
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   612
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Command1.Enabled = False
End Sub

Private Sub Text1_Change()
    If Text1 = "" Or Text2 = "" Then
        Command1.Enabled = False
    Else
        Command1.Enabled = True
    End If
End Sub

Private Sub Text2_Change()
    If Text1 = "" Or Text2 = "" Then
        Command1.Enabled = False
    Else
        Command1.Enabled = True
    End If
End Sub
