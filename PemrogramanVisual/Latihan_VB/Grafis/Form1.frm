VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   5925
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7725
   DrawWidth       =   3
   LinkTopic       =   "Form1"
   ScaleHeight     =   5925
   ScaleWidth      =   7725
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdStart 
      Caption         =   "&Start"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1455
   End
   Begin VB.Timer tmrGrafis 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   240
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i, z As Integer

Private Sub cmdStart_Click()
    Me.tmrGrafis.Enabled = True
End Sub

Private Sub Form_Load()
    i = 100
    z = 1
End Sub

Private Sub tmrGrafis_Timer()
    Dim x As Integer, y As Integer
    i = i + 100
    z = z + 1
    x = Me.ScaleWidth \ 2
    y = Me.ScaleHeight \ 2
    Circle (x, y), i
    ForeColor = QBColor(z)
    If z = 15 Then z = 1
    If i > ScaleWidth - 5000 Then
        Cls
        i = 100
    End If
End Sub
