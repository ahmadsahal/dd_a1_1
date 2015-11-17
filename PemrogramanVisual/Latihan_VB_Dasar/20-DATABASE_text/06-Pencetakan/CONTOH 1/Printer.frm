VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2352
   ClientLeft      =   48
   ClientTop       =   480
   ClientWidth     =   3744
   LinkTopic       =   "Form1"
   ScaleHeight     =   2352
   ScaleWidth      =   3744
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Cetak"
      Height          =   372
      Left            =   360
      TabIndex        =   1
      Top             =   1080
      Width           =   1212
   End
   Begin VB.TextBox Text1 
      Height          =   288
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   2292
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Printer.FontSize = 12
    Printer.Print Tab(10); Text1
    Printer.EndDoc
End Sub

