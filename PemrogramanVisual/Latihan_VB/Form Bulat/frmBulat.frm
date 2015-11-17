VERSION 5.00
Begin VB.Form frmBulat 
   BackColor       =   &H0000FFFF&
   Caption         =   "Form1"
   ClientHeight    =   3495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3720
   LinkTopic       =   "Form1"
   ScaleHeight     =   3495
   ScaleWidth      =   3720
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Form Bulat"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   1200
      Width           =   2175
   End
End
Attribute VB_Name = "frmBulat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function _
CreateEllipticRgn Lib "gdi32" _
( _
ByVal X1 As Long, _
ByVal Y1 As Long, _
ByVal X2 As Long, _
ByVal Y2 As Long _
) As Long

Private Declare Function _
SetWindowRgn Lib "user32" _
( _
ByVal hWnd As Long, _
ByVal hRgn As Long, _
ByVal bRedraw As Boolean _
) As Long

Private Sub Form_Load()
    SetWindowRgn hWnd, _
    CreateEllipticRgn _
    (10, 10, 250, 250), _
    True
    frmEllips.Show
End Sub


