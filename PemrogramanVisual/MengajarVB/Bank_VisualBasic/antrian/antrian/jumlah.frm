VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Jumlah Antrian"
   ClientHeight    =   1830
   ClientLeft      =   1950
   ClientTop       =   2145
   ClientWidth     =   8445
   LinkTopic       =   "Form2"
   Picture         =   "jumlah.frx":0000
   ScaleHeight     =   1830
   ScaleWidth      =   8445
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "&Clear"
      Height          =   435
      Left            =   5760
      TabIndex        =   4
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   435
      Left            =   7080
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   435
      Left            =   4440
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   4440
      TabIndex        =   0
      Top             =   360
      Width           =   3855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "JUMLAH ANTRIAN"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   4095
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim no_antri, no_antri_panggil As Integer

Private Sub Command1_Click()
Form1.Textantri = Text1.Text
Form1.Command6.Enabled = True
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
Text1.Text = ""
End Sub

