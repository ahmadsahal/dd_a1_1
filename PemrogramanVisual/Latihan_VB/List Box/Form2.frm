VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   7740
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7470
   LinkTopic       =   "Form2"
   ScaleHeight     =   7740
   ScaleWidth      =   7470
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command9 
      Caption         =   "[ Z ]"
      Height          =   375
      Left            =   4320
      TabIndex        =   10
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton Command8 
      Caption         =   "\"
      Height          =   375
      Left            =   3600
      TabIndex        =   9
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton Command7 
      Caption         =   "||||"
      Height          =   375
      Left            =   2880
      TabIndex        =   8
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton Command6 
      Caption         =   "C"
      Height          =   375
      Left            =   6000
      TabIndex        =   7
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton Command5 
      Caption         =   "4"
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Caption         =   "3"
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "2"
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "1"
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   615
   End
   Begin VB.PictureBox Picture2 
      Height          =   6855
      Left            =   0
      Picture         =   "Form2.frx":0000
      ScaleHeight     =   6795
      ScaleWidth      =   4515
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   6720
      TabIndex        =   1
      Top             =   0
      Width           =   615
   End
   Begin VB.PictureBox Picture1 
      Height          =   7935
      Left            =   120
      ScaleHeight     =   7875
      ScaleWidth      =   11595
      TabIndex        =   0
      Top             =   600
      Width           =   11655
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a

Private Sub Command1_Click()
   Picture1.PaintPicture Picture2.Picture, _
   (Picture1.ScaleWidth - 3000) / 2, _
    (Picture1.ScaleHeight - 2000) / 2
End Sub

Private Sub Command2_Click()
Picture1.PaintPicture Picture2.Picture, 0, 0, , , , , _
Picture2.ScaleWidth / 2, Picture2.ScaleHeight / 2
End Sub

Private Sub Command5_Click()
wi = Picture2.ScaleWidth / 2
he = Picture2.ScaleHeight / 2
Picture1.PaintPicture Picture2.Picture, _
wi, he, , , wi, he, wi, he
End Sub

Private Sub Command3_Click()
wi = Picture2.ScaleWidth / 2
he = Picture2.ScaleHeight / 2
Picture1.PaintPicture Picture2.Picture, _
0, he, , , 0, he, _
wi, he
End Sub

Private Sub Command4_Click()
wi = Picture2.ScaleWidth / 2
he = Picture2.ScaleHeight / 2
Picture1.PaintPicture Picture2.Picture, _
wi, 0, , , wi, 0, _
wi, he
End Sub

Private Sub Command6_Click()
Me.Picture1.Cls
End Sub

Private Sub Command7_Click()
    X = 0
Do While X < Picture1.ScaleWidth
    Y = 0
    ' For each column, start at the top and work downward.
    Do While Y < Picture1.ScaleHeight
        Picture1.PaintPicture Picture2.Picture, X, Y, , , 0, 0
        ' Next row
        Y = Y + Picture2.ScaleHeight
    Loop
    ' Next column
    X = X + Picture2.ScaleWidth
Loop

End Sub

Private Sub Command8_Click()
 
Picture1.PaintPicture Picture2.Picture, _
    Picture2.ScaleWidth, 0, -Picture2.ScaleWidth
' Flip vertically.
Picture1.PaintPicture Picture2.Picture, 0, _
    Picture2.ScaleHeight, , -Picture2.ScaleHeight
' Flip the image on both axes.
Picture1.PaintPicture Picture2.Picture, Picture2.ScaleWidth, _
    Picture2.ScaleHeight, -Picture2.ScaleWidth, -Picture2.ScaleHeight
End Sub

Private Sub Command9_Click()
If a = 0 Then
    Picture1.PaintPicture Picture2.Picture, 0, 0, _
    Picture2.ScaleWidth * 2, Picture2.ScaleHeight * 2
    a = 1
Else: a = 0
    Picture1.PaintPicture Picture2.Picture, 0, 0, _
    Picture2.ScaleWidth * 4, Picture2.ScaleHeight * 4
End If
End Sub

Private Sub Form_Load(): a = 0
' StdPicture's Width and Height properties are expressed in
' Himetric units.
With Picture1
    Width = CInt(.ScaleX(.Picture.Width, vbHimetric, vbPixels))
    Height = CInt(.ScaleY(.Picture.Height, vbHimetric, _
        vbPixels))
End With

End Sub

