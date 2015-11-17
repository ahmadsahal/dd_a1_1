VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "HOTEL"
   ClientHeight    =   2664
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   3456
   LinkTopic       =   "Form1"
   ScaleHeight     =   2664
   ScaleWidth      =   3456
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Lama 
      Height          =   288
      Left            =   960
      TabIndex        =   7
      Top             =   1800
      Width           =   1212
   End
   Begin VB.CommandButton Command1 
      Caption         =   "HITUNG"
      Height          =   372
      Left            =   240
      TabIndex        =   6
      Top             =   1200
      Width           =   852
   End
   Begin VB.TextBox Keluar 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1057
         SubFormatType   =   3
      EndProperty
      Height          =   288
      Left            =   960
      TabIndex        =   1
      Top             =   720
      Width           =   1212
   End
   Begin VB.TextBox Masuk 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1057
         SubFormatType   =   3
      EndProperty
      Height          =   288
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   1212
   End
   Begin VB.Label Label5 
      Caption         =   "Lama"
      Height          =   252
      Left            =   240
      TabIndex        =   9
      Top             =   1800
      Width           =   612
   End
   Begin VB.Label Label8 
      Caption         =   "hari"
      Height          =   252
      Left            =   2400
      TabIndex        =   8
      Top             =   1800
      Width           =   372
   End
   Begin VB.Label Label7 
      Caption         =   "mm/dd/yy"
      Height          =   252
      Left            =   2400
      TabIndex        =   5
      Top             =   720
      Width           =   852
   End
   Begin VB.Label Label6 
      Caption         =   "mm/dd/yy"
      Height          =   252
      Left            =   2400
      TabIndex        =   4
      Top             =   240
      Width           =   972
   End
   Begin VB.Label Label4 
      Caption         =   "Keluar"
      Height          =   252
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   612
   End
   Begin VB.Label Label3 
      Caption         =   "Masuk"
      Height          =   252
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   732
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Lama = DateDiff("d", Masuk, Keluar)
End Sub

