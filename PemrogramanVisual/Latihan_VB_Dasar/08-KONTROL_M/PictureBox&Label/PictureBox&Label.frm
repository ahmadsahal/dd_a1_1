VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "PictureBox & Label"
   ClientHeight    =   2400
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   3744
   LinkTopic       =   "Form1"
   ScaleHeight     =   2400
   ScaleWidth      =   3744
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   612
      Left            =   120
      ScaleHeight     =   564
      ScaleWidth      =   3444
      TabIndex        =   0
      Top             =   120
      Width           =   3492
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "INPUT DATA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   3012
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
