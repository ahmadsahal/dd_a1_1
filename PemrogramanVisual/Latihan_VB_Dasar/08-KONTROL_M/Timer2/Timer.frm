VERSION 5.00
Begin VB.Form Jam 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   2400
   ClientLeft      =   36
   ClientTop       =   420
   ClientWidth     =   3744
   LinkTopic       =   "Form1"
   ScaleHeight     =   2400
   ScaleWidth      =   3744
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   1320
      Top             =   1320
   End
   Begin VB.Label TampilJam 
      Caption         =   "Label1"
      Height          =   612
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   2172
   End
End
Attribute VB_Name = "Jam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    TampilJam.Top = ScaleTop
    TampilJam.Left = ScaleLeft
    TampilJam.Width = ScaleWidth
    TampilJam.Height = ScaleHeight
End Sub

Private Sub Timer1_Timer()
    If Jam.WindowState = vbNormal Then
        TampilJam.Caption = CStr(Time)
        Jam.Caption = Format(Date, "Long Date")
    Else
        Jam.Caption = CStr(Time)
    End If
End Sub

