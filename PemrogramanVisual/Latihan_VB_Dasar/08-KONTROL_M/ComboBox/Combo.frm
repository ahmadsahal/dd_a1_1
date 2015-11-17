VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2532
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   3792
   LinkTopic       =   "Form1"
   ScaleHeight     =   2532
   ScaleWidth      =   3792
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   372
      Left            =   2280
      TabIndex        =   4
      Top             =   1440
      Width           =   972
   End
   Begin VB.ComboBox Combo1 
      Height          =   288
      Left            =   1440
      TabIndex        =   3
      Text            =   "Pria"
      Top             =   840
      Width           =   972
   End
   Begin VB.TextBox Text1 
      Height          =   288
      Left            =   1440
      TabIndex        =   2
      Top             =   360
      Width           =   1812
   End
   Begin VB.Label Label2 
      Caption         =   "Jenis kelamin"
      Height          =   252
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   1092
   End
   Begin VB.Label Label1 
      Caption         =   "Nama Anda"
      Height          =   252
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1092
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    MsgBox "Nama= " & Text1 & ", Jenis kelamin = " & Combo1.Text, vbOKOnly, "Data Anda"
End Sub

Private Sub Form_Load()
    Combo1.AddItem "Pria", 0
    Combo1.AddItem "Wanita", 1
End Sub

