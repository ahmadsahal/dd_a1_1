VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Math Class"
   ClientHeight    =   4305
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5055
   LinkTopic       =   "Form1"
   ScaleHeight     =   4305
   ScaleWidth      =   5055
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Operasi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2652
      Left            =   2160
      TabIndex        =   7
      Top             =   1320
      Width           =   2532
      Begin VB.OptionButton Option4 
         Caption         =   "Sisa (Modulus)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   240
         TabIndex        =   11
         Top             =   1800
         Width           =   2052
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Pangkat"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   240
         TabIndex        =   10
         Top             =   1320
         Width           =   2052
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Bagi"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   240
         TabIndex        =   9
         Top             =   840
         Width           =   2052
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Kali"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Value           =   -1  'True
         Width           =   2052
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Hitung"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   852
      Left            =   480
      TabIndex        =   6
      Top             =   2640
      Width           =   1332
   End
   Begin VB.TextBox Text2 
      Height          =   372
      Left            =   600
      TabIndex        =   1
      Top             =   1560
      Width           =   972
   End
   Begin VB.TextBox Text1 
      Height          =   372
      Left            =   600
      TabIndex        =   0
      Top             =   600
      Width           =   972
   End
   Begin VB.Label LblHasil 
      Caption         =   "HasilOp"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2280
      TabIndex        =   5
      Top             =   720
      Width           =   2172
   End
   Begin VB.Label Label3 
      Caption         =   "Hasil"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2280
      TabIndex        =   4
      Top             =   120
      Width           =   1092
   End
   Begin VB.Label Label2 
      Caption         =   "Nilai Akhir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   600
      TabIndex        =   3
      Top             =   1080
      Width           =   1092
   End
   Begin VB.Label Label1 
      Caption         =   "Nilai Awal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   600
      TabIndex        =   2
      Top             =   120
      Width           =   1092
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim x As New Kali

'kalau tidak diisi salah satu maka OUT!!
If Text1.Text = "" Then Exit Sub
If Text2.Text = "" Then Exit Sub

'pengisian nilai pada properti
x.Awal = Val(Text1.Text)
x.Akhir = Val(Text2.Text)


If Option1.Value = True Then
    x.Kali
    LblHasil.Caption = x.Hasil
End If

If Option2.Value = True Then
    x.Bagi
    LblHasil.Caption = x.Hasil
End If

If Option3.Value = True Then
    x.Pangkat x.Awal, x.Akhir
    LblHasil.Caption = x.Hasil
End If
If Option4.Value = True Then
    LblHasil.Caption = x.Sisa(x.Awal, x.Akhir)
End If

'menghancurkan objek dari memori
Set x = Nothing

End Sub

Private Sub Form_Load()
LblHasil.Caption = ""
Option1.Value = True
Text1.Text = "0"
Text2.Text = "0"
End Sub
