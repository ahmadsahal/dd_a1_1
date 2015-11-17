VERSION 5.00
Begin VB.Form frmLogo 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   7185
   ClientLeft      =   270
   ClientTop       =   1425
   ClientWidth     =   8370
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Enabled         =   0   'False
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H80000004&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   7185
   ScaleWidth      =   8370
   Begin VB.Label lblCorp 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Copyright (C) 2003 By Viansastra"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2280
      TabIndex        =   0
      Top             =   5040
      UseMnemonic     =   0   'False
      Width           =   3015
   End
   Begin VB.Image imaLogo 
      Height          =   3555
      Left            =   2280
      Picture         =   "frmLogo.frx":0000
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   3015
   End
End
Attribute VB_Name = "frmLogo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
    Me.imaLogo.Move (Me.Width - Me.imaLogo.Width) \ 2, _
    (Me.Height - Me.imaLogo.Height) \ 2
    Me.lblCorp.Move (Me.imaLogo.Left), _
    (Me.imaLogo.Top + Me.imaLogo.Height)
End Sub
