VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form5 
   Caption         =   "Form5"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form5"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.Slider Slider1 
      Height          =   630
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   1111
      _Version        =   393216
      SelectRange     =   -1  'True
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StartSelection As Single

Private Sub Slider1_MouseDown(Button As Integer, Shift As Integer, _
    x As Single, y As Single)
    If Shift = vbShiftMask Then
        ' If the shift key is being pressed, enter select range mode.
        Slider1.SelectRange = True
        Slider1.SelLength = 0
        StartSelection = Slider1.Value
    Else
        ' Else cancel any active select range mode.
        Slider1.SelectRange = False
    End If
End Sub

 
Private Sub Slider1_Scroll()
    If Slider1.SelectRange Then
        ' The indicator is being moved in SelectRange mode.
        If Slider1.Value > StartSelection Then
            Slider1.SelStart = StartSelection
            Slider1.SelLength = Slider1.Value - StartSelection
        Else
            Slider1.SelStart = Slider1.Value
            Slider1.SelLength = StartSelection - Slider1.Value
        End If
    End If
End Sub

 

