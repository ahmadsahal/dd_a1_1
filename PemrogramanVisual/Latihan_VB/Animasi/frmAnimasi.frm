VERSION 5.00
Begin VB.Form frmAnimasi 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Animasi Label"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4755
   Icon            =   "frmAnimasi.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   4755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrAni2 
      Interval        =   100
      Left            =   720
      Top             =   1080
   End
   Begin VB.Timer tmrAni1 
      Interval        =   100
      Left            =   720
      Top             =   600
   End
   Begin VB.CommandButton cmdKeluar 
      Caption         =   "&Keluar"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1560
      TabIndex        =   2
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label lblAni2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   1560
      Width           =   75
   End
   Begin VB.Label lblAni1 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4875
   End
End
Attribute VB_Name = "frmAnimasi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' deklarasi variable level general
Dim X$, Y$
Dim n As Byte
' saat form akan ditampilkan

Private Sub Form_Load()
    X$ = " SELAMAT DATANG PROGRAMER..."
    Y$ = " << Copyright(C) 2003 " & _
    "By Viansastra >> "
End Sub

' event pergerakan oleh timer
Private Sub tmrAni1_Timer()
    n = n + 1
    Me.lblAni1.Caption = Left(X$, n)
    If n > Len(X$) Then n = 0
End Sub

' event pergerakan oleh timer
Private Sub tmrAni2_Timer()
    Y$ = Right(Y$, Len _
    (Y$) - 1) & Left(Y$, 1)
    Me.lblAni2.Caption = Y$
 End Sub

' event mengklik tombol keluar
Private Sub cmdKeluar_Click()
   Do
      DoEvents
      Me.Left = Trim(Str(Int(Me.Left) - 30))
   Loop Until Me.Left < -Screen.Width
        Unload Me
End Sub


