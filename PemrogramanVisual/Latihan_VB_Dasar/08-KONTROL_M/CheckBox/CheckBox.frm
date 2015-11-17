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
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   492
      Left            =   1200
      TabIndex        =   0
      Top             =   240
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   252
      Left            =   1200
      TabIndex        =   1
      Top             =   960
      Width           =   1572
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
    If Check1.Value = 1 Then
        Check1.Caption = "Hidup"
        Label1.Visible = True
        Label1.Caption = "Aku dihidupkan"
    Else
        Check1.Caption = "Mati"
        Label1.Visible = False
    End If
End Sub

Private Sub Form_Load()
    Check1.Value = 1
    Check1.Caption = "Hidup"
    Label1.Visible = True
    Label1.Caption = "Aku dihidupkan"
End Sub
