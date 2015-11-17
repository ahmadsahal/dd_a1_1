VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7725
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   7725
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4440
      TabIndex        =   7
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton cmdSelectAll 
      Caption         =   "Select &All"
      Height          =   375
      Left            =   2040
      TabIndex        =   6
      Top             =   2760
      Width           =   1095
   End
   Begin VB.ListBox lstRight 
      Height          =   2595
      Left            =   4920
      MultiSelect     =   2  'Extended
      TabIndex        =   5
      Top             =   0
      Width           =   2175
   End
   Begin VB.CommandButton cmdBackAll 
      Caption         =   "<<"
      Height          =   495
      Left            =   3480
      TabIndex        =   4
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "<"
      Height          =   495
      Left            =   3480
      TabIndex        =   3
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton cmdMoveAll 
      Caption         =   ">>"
      Height          =   495
      Left            =   3480
      TabIndex        =   2
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   ">"
      Height          =   495
      Left            =   3480
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.ListBox lstLeft 
      Height          =   2595
      ItemData        =   "Form1.frx":0000
      Left            =   1080
      List            =   "Form1.frx":0013
      MultiSelect     =   2  'Extended
      TabIndex        =   0
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdMove_Click()
    ' Move one item from left to right.
    If lstLeft.ListIndex >= 0 Then
        lstRight.AddItem lstLeft.Text
        lstLeft.RemoveItem lstLeft.ListIndex
    End If
End Sub

Private Sub cmdMoveAll_Click()
    ' Move all items from left to right.
    Do While lstLeft.ListCount
        lstRight.AddItem lstLeft.List(0)
        lstLeft.RemoveItem 0
    Loop
End Sub

Private Sub cmdBack_Click()
    ' Move one item from right to left.
    If lstRight.ListIndex >= 0 Then
        lstLeft.AddItem lstRight.Text
        lstRight.RemoveItem lstRight.ListIndex
    End If
End Sub

Private Sub cmdBackAll_Click()
    ' Move all items from right to left.
    Do While lstRight.ListCount
        lstLeft.AddItem lstRight.List(0)
        lstRight.RemoveItem 0
    Loop
End Sub


Private Sub cmdSelectAll_Click()
    Dim i As Long, saveIndex As Long, saveTop As Long
    ' Save current state.
    saveIndex = lstRight.ListIndex
    saveTop = lstRight.TopIndex
    ' Make the list box invisible to avoid flickering.
    lstRight.Visible = False
    ' Change the select state for all items.
    For i = 0 To lstRight.ListCount - 1
        lstRight.Selected(i) = True
    Next
    ' Restore original state, and make the list box visible again.
    lstRight.TopIndex = saveTop
    lstRight.ListIndex = saveIndex
    lstRight.Visible = True
End Sub

Private Sub lstLeft_DblClick()
    ' Simulate a click on the Move button.
    cmdMove.Value = True
End Sub

Private Sub lstRight_DblClick()
    ' Simulate a click on the Back button.
    cmdBack.Value = True
End Sub


Private Sub lstLeft_Click()
    cmdOK.Enabled = (lstLeft.SelCount > 0)
End Sub

Private Sub cmdOK_Click()

Dim i As Long
For i = 0 To lstLeft.ListCount - 1
    If lstLeft.Selected(i) Then lstRight.AddItem lstLeft.List(i)
Next
 
For i = 0 To lstLeft.ListCount - 1
    lstLeft.Selected(i) = False
Next
End Sub
 



 


