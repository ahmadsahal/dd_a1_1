VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   5700
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7260
   LinkTopic       =   "Form4"
   ScaleHeight     =   5700
   ScaleWidth      =   7260
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   5880
      TabIndex        =   2
      Top             =   5160
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   0
      Left            =   3000
      TabIndex        =   1
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   720
      Top             =   2400
   End
   Begin VB.Label lblClock 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   840
      TabIndex        =   0
      Top             =   360
      Width           =   2895
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
For i = 1 To 10
    Load Text1(i)
' Move the new control where you need it, and resize it.
Text1(i).Move (i * 1), (i * 400), 800, 350
' Set other properties as required.
Text1(i).MaxLength = 10

' Finally make it visible.
Text1(i).Visible = True
Next i

End Sub

Private Sub Timer1_Timer()
Dim strTime As String
    strTime = Time$
    If Mid$(lblClock.Caption, 3, 1) = ":" Then
        Mid$(strTime, 3, 1) = " "
        Mid$(strTime, 6, 1) = " "
    End If
    lblClock.Caption = strTime
End Sub
