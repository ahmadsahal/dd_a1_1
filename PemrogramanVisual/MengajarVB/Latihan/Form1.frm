VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4245
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9405
   LinkTopic       =   "Form1"
   ScaleHeight     =   4245
   ScaleWidth      =   9405
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   5040
      TabIndex        =   3
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox txtNama 
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   960
      TabIndex        =   2
      Top             =   1560
      Width           =   6375
   End
   Begin VB.Label Label1 
      Caption         =   "Nama"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
If UCase(txtNama.Text) = "SINI" Then Label2.Caption = "Disinilah tempatnya"
End Sub
