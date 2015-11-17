VERSION 5.00
Begin VB.Form frmStatement1 
   Caption         =   "Dengan Option Explicit"
   ClientHeight    =   2670
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3960
   Icon            =   "frmStatement1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   3960
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtHasil 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdHitung 
      Caption         =   "&Hitung"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   0
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "10+20"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   2850
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   3
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   570
   End
End
Attribute VB_Name = "frmStatement1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

   Dim Angka1 As Integer, Angka2 As Integer
   Dim Angka3 As Integer

   Private Sub Form_Load()
      Angka1 = 10
      Angka2 = 20
        ' menampilkan frmstatement2
      frmStatement2.Show
      frmStatement2.Left = Me.Left + 4100
      frmStatement2.Top = Me.Top
   End Sub

   Private Sub cmdHitung_Click()
     ' terjadi kesalahan dalam penulisan
     ' Variabel Angka1
     Angka3 = Agnka1 + Angka2
     Me.txtHasil.Text = Val(Angka3)
   End Sub
