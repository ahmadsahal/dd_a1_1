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
   Begin VB.ListBox List1 
      Height          =   1392
      Left            =   840
      TabIndex        =   0
      Top             =   360
      Width           =   1332
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    List1.AddItem "Henri Dunan", 0
    List1.AddItem "Ike Nurianty", 1
    List1.AddItem "Dion Arkade", 2
    List1.AddItem "Yusi Ismiati", 3
    List1.AddItem 5000, 4
    List1.AddItem 8000, 5
    List1.AddItem 6000, 6
    List1.AddItem 7000, 7
End Sub

Private Sub List1_Click()
    If List1.ListIndex = -1 Then
        MsgBox "Anda tidak memilih apa-apa"
    Else
        MsgBox "Anda memilih " & List1.Text & "(" & List1.ListIndex & ")"
    End If
End Sub
