VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "INDEX PROGRAM"
   ClientHeight    =   4005
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7860
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MouseIcon       =   "Form1.frx":0442
   ScaleHeight     =   4005
   ScaleWidth      =   7860
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picSplitter 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H00808080&
      Height          =   4800
      Left            =   3360
      ScaleHeight     =   2090.126
      ScaleMode       =   0  'User
      ScaleWidth      =   1404
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   135
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   240
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":074C
            Key             =   "C"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0A68
            Key             =   "D"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0EBC
            Key             =   "B"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1310
            Key             =   "A"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView trvProgram 
      Height          =   3975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   7011
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   882
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      FullRowSelect   =   -1  'True
      ImageList       =   "ImageList1"
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ListView lvwProgram 
      Height          =   3735
      Left            =   3720
      TabIndex        =   2
      Top             =   120
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   6588
      View            =   1
      Arrange         =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      TextBackground  =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Image imgSplitter 
      Height          =   4785
      Left            =   3000
      MouseIcon       =   "Form1.frx":1764
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   120
      Width           =   135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mbMoving As Boolean
Const sglSplitLimit = 500

Private Sub Form_Load()
Dim i
    With trvProgram.Nodes
        .Add , 4, "A", "Master", "A", "A"
        .Add 1, 4, "B", "Adodc", "B", "B"
        .Add 1, 4, "C", "ListView", "B", "B"
        .Add 1, 4, "D", "File", "B", "B"
        .Add 1, 4, "E", "Multimedia", "B", "B"
        .Add 1, 4, "F", "Mdi", "B", "B"
        .Add 1, 4, "G", "Zodiac", "B", "B"
        .Add 2, 4, "H", "Objek", "C", "C"
        .Add 2, 4, "I", "Keterangan", "C", "C"
        .Add 3, 4, "J", "Objek", "C", "C"
        .Add 3, 4, "K", "Keterangan", "C", "C"
        .Add 4, 4, "L", "Objek", "C", "C"
        .Add 4, 4, "M", "Keterangan", "C", "C"
        .Add 5, 4, "N", "Objek", "C", "C"
        .Add 5, 4, "O", "Keterangan", "C", "C"
        .Add 6, 4, "P", "Objek", "C", "C"
        .Add 6, 4, "Q", "Keterangan", "C", "C"
        .Add 7, 4, "R", "Objek", "C", "C"
        .Add 7, 4, "S", "Keterangan", "C", "C"
       ' .Add , 3, "App", "Aplikasi", "D", "D"
    End With
    Me.lvwProgram.ListItems.Add , "A", "Master", , "A"
End Sub

Private Sub Form_Resize(): On Error Resume Next
    trvProgram.Move 0, 0, trvProgram.Width, ScaleHeight
    lvwProgram.Move lvwProgram.Left, 0, ScaleWidth - trvProgram.Width - _
    160, ScaleHeight
    SizeControls imgSplitter.Left
End Sub

Private Sub imgSplitter_MouseDown(Button As Integer, _
Shift As Integer, X As Single, Y As Single)
    With imgSplitter
        picSplitter.Move .Left, .Top, .Width \ 2, _
        .Height - 20
    End With
    picSplitter.Visible = True
    mbMoving = True
End Sub


Private Sub imgSplitter_MouseMove(Button As Integer, _
Shift As Integer, X As Single, Y As Single)
    Dim sglPos As Single
    
    If mbMoving Then
        sglPos = X + imgSplitter.Left
        If sglPos < sglSplitLimit Then
            picSplitter.Left = sglSplitLimit
        ElseIf sglPos > Me.Width - sglSplitLimit Then
            picSplitter.Left = Me.Width - sglSplitLimit
        Else
            picSplitter.Left = sglPos
        End If
    End If
End Sub


Private Sub imgSplitter_MouseUp(Button As Integer, _
Shift As Integer, X As Single, Y As Single)
    SizeControls picSplitter.Left
    picSplitter.Visible = False
    mbMoving = False
End Sub

Private Sub trvprogram_DragDrop(Source As Control, _
X As Single, Y As Single)
    If Source = imgSplitter Then
        SizeControls X
    End If
End Sub


Sub SizeControls(X As Single)
    On Error Resume Next
    If X < 1500 Then X = 1500
    If X > (Me.Width - 1500) Then X = Me.Width - 1000
    trvProgram.Width = X
    imgSplitter.Left = X
    lvwProgram.Left = X + 80
    lvwProgram.Width = Me.Width - (trvProgram.Width + 280)
    imgSplitter.Top = trvProgram.Top
    imgSplitter.Height = trvProgram.Height
End Sub

Private Sub trvProgram_NodeClick( _
ByVal Node As MSComctlLib.Node)
    With Me.lvwProgram.ListItems
    Select Case Node.Key
        Case Is = "A"
            .Clear
            .Add , "B", "Adodc", , "B"
            .Add , "C", "ListView", , "B"
            .Add , "D", "File", , "B"
            .Add , "E", "Multimedia", , "B"
            .Add , "F", "MDI", , "B"
            .Add , "G", "Zodiac", , "B"
    End Select
    End With
End Sub
