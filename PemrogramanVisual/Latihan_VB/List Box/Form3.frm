VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   5655
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6750
   LinkTopic       =   "Form3"
   ScaleHeight     =   5655
   ScaleWidth      =   6750
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   6375
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   4935
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton cmdFiller 
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   5160
      Width           =   1575
   End
   Begin VB.PictureBox picCanvas 
      Height          =   2055
      Left            =   240
      Picture         =   "Form3.frx":0000
      ScaleHeight     =   1995
      ScaleWidth      =   2115
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Const SB_WIDTH = 300    ' width of vertical scrollbars
Const SB_HEIGHT = 300   ' height of horizontal scrollbars

Private Sub Form_Resize()
    ' Resize the scroll bars along the form.
    HScroll1.Move 0, ScaleHeight - SB_HEIGHT, ScaleWidth - SB_WIDTH
    VScroll1.Move ScaleWidth - SB_WIDTH, 0, SB_WIDTH, _
        ScaleHeight - SB_HEIGHT
    cmdFiller.Move ScaleWidth - SB_WIDTH, ScaleHeight - SB_HEIGHT, _
        SB_WIDTH, SB_HEIGHT

    ' Put these controls on top.
    HScroll1.ZOrder
    VScroll1.ZOrder
    cmdFiller.ZOrder
    picCanvas.BorderStyle = 0

    ' A click on the arrow moves one pixel.
    HScroll1.SmallChange = ScaleX(1, vbPixels, vbTwips)
    VScroll1.SmallChange = ScaleY(1, vbPixels, vbTwips)
    ' A click on the scroll bar moves 16 pixels.
    HScroll1.LargeChange = HScroll1.SmallChange * 16
    VScroll1.LargeChange = VScroll1.SmallChange * 16

    ' If the form is larger than the picCanvas picture box,
    ' we don't need to show the corresponding scroll bar.
    If ScaleWidth < picCanvas.Width + SB_WIDTH Then
        HScroll1.Visible = True
        HScroll1.Max = picCanvas.Width + SB_WIDTH - ScaleWidth
    Else
        HScroll1.Value = 0
       ' HScroll1.Visible = False
    End If
    If ScaleHeight < picCanvas.Height + SB_HEIGHT Then
        VScroll1.Visible = True
        VScroll1.Max = picCanvas.Height + SB_HEIGHT - ScaleHeight
    Else
        VScroll1.Value = 0
     '   VScroll1.Visible = False
    End If
    ' Make the filler control visible only if necessary.
    cmdFiller.Visible = (HScroll1.Visible Or VScroll1.Visible)
    MoveCanvas
End Sub

Sub MoveCanvas()
    picCanvas.Move -HScroll1.Value, -VScroll1.Value
End Sub

 

