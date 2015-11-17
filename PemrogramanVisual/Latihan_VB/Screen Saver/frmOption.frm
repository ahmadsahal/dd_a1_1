VERSION 5.00
Begin VB.Form frmOption 
   BorderStyle     =   0  'None
   ClientHeight    =   2055
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4785
   DrawStyle       =   5  'Transparent
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   20.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   137
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   319
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraOption 
      BackColor       =   &H00000000&
      Caption         =   "Lines"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   1335
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   4815
      Begin VB.TextBox txtLineCount 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   5
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   4
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   3
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox txtLineWidth 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Line &Count"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   375
         Index           =   0
         Left            =   1680
         TabIndex        =   6
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Line &Width"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefDbl A-Z

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdOK_Click()
On Error Resume Next
    SaveSetting "VianScr", "options", _
    "linewidth", txtLineWidth.Text
    SaveSetting "VianScr", "options", _
    "lineCount", txtLineCount.Text
    Unload Me
End Sub

Private Sub Form_Load()
    Left = (Screen.Width - Width) \ 2
    Top = (Screen.Height - Height) \ 2
    txtLineCount.Text = GetSetting( _
    "VianScr", "options", "lineCount", "50")
    txtLineWidth.Text = GetSetting( _
    "VianScr", "options", "linewidth", "1")
End Sub

Private Sub Form_Paint()
Dim t, L, y
    t = 60
    L = Me.ScaleWidth
        For y = 0 To t Step t / 777
            Me.FillColor = _
            RGB(0, 100, 255 - _
            (y * 255 \ t))
            Me.Line _
            (-1, y - 1)- _
            (L, y + 1), , B
        Next y
    DrawBackGround
End Sub

Sub DrawBackGround() ' Mencetak label pada form
CurrentX = 15: CurrentY = -2: ForeColor = QBColor(4)
Print "Created By Viansastra"
CurrentX = 10: CurrentY = 5: ForeColor = QBColor(14)
Print "Created By Viansastra"
End Sub
