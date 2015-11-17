VERSION 5.00
Begin VB.Form frmPaint 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Paint"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5790
   DrawStyle       =   5  'Transparent
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   5790
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboLine 
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      ItemData        =   "frmPaint.frx":0000
      Left            =   3120
      List            =   "frmPaint.frx":0002
      TabIndex        =   4
      Top             =   3000
      Width           =   2055
   End
   Begin VB.ComboBox cboWarna 
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      ItemData        =   "frmPaint.frx":0004
      Left            =   120
      List            =   "frmPaint.frx":0006
      TabIndex        =   3
      Top             =   3000
      Width           =   2175
   End
   Begin VB.PictureBox picPaint 
      DrawStyle       =   5  'Transparent
      FillStyle       =   0  'Solid
      Height          =   2655
      Left            =   0
      ScaleHeight     =   100
      ScaleMode       =   0  'User
      ScaleWidth      =   100
      TabIndex        =   2
      Top             =   0
      Width           =   4935
   End
   Begin VB.VScrollBar vscWarna 
      Height          =   2655
      LargeChange     =   10
      Left            =   4920
      Max             =   500
      SmallChange     =   10
      TabIndex        =   1
      Top             =   0
      Width           =   255
   End
   Begin VB.HScrollBar hscWarna 
      Height          =   255
      LargeChange     =   10
      Left            =   0
      Max             =   500
      SmallChange     =   10
      TabIndex        =   0
      Top             =   2640
      Width           =   4935
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2640
      TabIndex        =   6
      Top             =   3000
      Width           =   180
   End
   Begin VB.Label lblY 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5280
      TabIndex        =   5
      Top             =   1320
      UseMnemonic     =   0   'False
      Width           =   165
   End
End
Attribute VB_Name = "frmPaint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim T, L, i, x, y, WPaint, WLine

Private Sub cboLine_Change()
    Paint
End Sub

Private Sub cboLine_Click()
    Paint
End Sub

Private Sub cboWarna_Change()
    Paint
End Sub

Private Sub cboWarna_Click()
    Paint
End Sub

Sub Form_Activate()
    Me.cboWarna.ListIndex = 0
    Me.cboLine.ListIndex = 0
    Paint
End Sub

Private Sub Paint()
    WPaint = Me.cboWarna.ListIndex
    WLine = Me.cboLine.ListIndex
    With Me.picPaint
        T = .ScaleHeight
        L = .ScaleWidth
            For i = 0 To T Step T / 300
                Select Case WPaint
                    Case Is = 0
                        .FillColor = _
                        RGB(255 - _
                        (i * 255 \ T), x, y)
                    Case Is = 1
                        .FillColor = _
                        RGB(x, 255 - _
                        (i * 255 \ T), y)
                    Case Is = 2
                        .FillColor = _
                        RGB(x, y, 255 - _
                        (i * 255 \ T))
                End Select
                Select Case WLine
                    Case 0
                        picPaint.Line _
                        (0, i - 1)- _
                        (L, i - 1), , B
                    Case 1
                        picPaint.Line _
                        (i, L - 1)- _
                        (L, i - 1), , B
                    Case 2
                        picPaint.Line _
                        (i, L - 1)- _
                        (L, i - L), , B
                    End Select
            Next i
    End With
End Sub

Private Sub Form_Load()
    With Me.cboWarna
        .AddItem "Merah"
        .AddItem "Hijau"
        .AddItem "Biru"
    End With
    With Me.cboLine
        .AddItem "Vertical"
        .AddItem "Diagonal"
        .AddItem "Horizontal"
    End With
End Sub

Private Sub hscWarna_Change()
    x = Me.hscWarna.Value
    Paint
    Me.lblX.Caption = x
End Sub

Private Sub picPaint_Click()
    SavePicture Me.picPaint.Image, App.Path & "\Coba.jpg"
End Sub

Private Sub vscWarna_Change()
    y = Me.vscWarna.Value
    Paint
    Me.lblY.Caption = y
End Sub
