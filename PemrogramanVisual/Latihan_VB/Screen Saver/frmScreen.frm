VERSION 5.00
Begin VB.Form frmScreen 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   3780
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6915
   ControlBox      =   0   'False
   DrawWidth       =   10
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   26.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmScreen.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   6915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   5400
      Top             =   4320
   End
   Begin VB.Timer tmrGrafis 
      Interval        =   10
      Left            =   1560
      Top             =   1560
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Created By Viansastra"
      ForeColor       =   &H0080FF80&
      Height          =   735
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   6210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Created By Viansastra"
      ForeColor       =   &H0000FFFF&
      Height          =   735
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5730
   End
End
Attribute VB_Name = "frmScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

DefLng A-Z
Private Declare Function _
SystemParametersInfo Lib "user32" Alias _
"SystemParametersInfoA" (ByVal uAction As Long, _
ByVal uParam As Long, _
ByRef lpvParam As Any, _
ByVal fuWinIni As Long) As Long

Private Declare Function ShowCursor Lib "user32" ( _
ByVal bShow As Long) As Long

Const SPI_SETSCREENSAVEACTIVE = 17

Private j, k
Private z
    
Private Sub Form_KeyDown(KeyCode As _
Integer, Shift As Integer)
    Unload Me: End
End Sub

Private Sub Form_Load():
Dim x
j = 10: z = 1
    k = GetSetting("VianScr", _
    "options", "lineCount", 50)
    Me.DrawWidth = GetSetting( _
    "VianScr", "options", _
    "linewidth", 4)
    x = SystemParametersInfo( _
    SPI_SETSCREENSAVEACTIVE, _
    0, ByVal 0&, 0)
        Select Case UCase$( _
        Left$(Command$, 2))
           Case "/P"
                Unload Me: End
                Exit Sub
            Case "/C"
                frmOption.Show
                Unload Me
                Exit Sub
            Case "/A"
                MsgBox "Ngga ada " & _
                "password di " & _
                "Screen Saver ini"
                Unload Me: End
               Exit Sub
            Case "/S"
                x = ShowCursor(False)
                tmrGrafis.Enabled = True
                Case Else
                   Unload Me: End
                  x = ShowCursor(True)
                   Exit Sub
        End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim x
     x = SystemParametersInfo( _
     SPI_SETSCREENSAVEACTIVE, 2, ByVal 0&, 0)
     x = ShowCursor(True)
End Sub

Private Sub Form_MouseMove( _
Button As Integer, Shift As Integer, _
x As Single, y As Single)

Static XAkhir, YAkhir
Dim XSekarang
Dim YSekarang
'
    XSekarang = x
    YSekarang = y
        If XAkhir = 0 And _
        YAkhir = 0 Then
            XAkhir = XSekarang
            YAkhir = YSekarang
            Exit Sub
        End If
        If XSekarang <> XAkhir _
        Or YSekarang <> YAkhir Then
            Unload Me: End
        End If
End Sub

Private Sub tmrGrafis_Timer()
    Dim x1 As Integer, y1 As Integer
    j = j + k
    z = z + 1
    x1 = Me.ScaleWidth \ 2
    y1 = Me.ScaleHeight \ 2
    Circle (x1, y1), j
    ForeColor = QBColor(z)
    If z = 15 Then z = 1
    If j > ScaleWidth - 5000 Then
        Cls
        j = 100
    End If
End Sub
