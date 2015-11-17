VERSION 5.00
Begin VB.Form frmWaktu 
   BackColor       =   &H00FFFF00&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Waktu berjalan"
   ClientHeight    =   345
   ClientLeft      =   255
   ClientTop       =   1695
   ClientWidth     =   2400
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Bookman Old Style"
      Size            =   12
      Charset         =   0
      Weight          =   300
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmWaktu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   345
   ScaleWidth      =   2400
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   540
      Left            =   -600
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   375
   End
   Begin VB.Timer tmrTop 
      Interval        =   1
      Left            =   1920
      Top             =   600
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   20.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   660
      Index           =   0
      Left            =   -360
      TabIndex        =   1
      Top             =   -60
      Width           =   2685
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   20.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   660
      Index           =   1
      Left            =   -240
      TabIndex        =   2
      Top             =   -60
      Width           =   2565
   End
End
Attribute VB_Name = "frmWaktu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos _
Lib "user32" (ByVal hwnd As Long, _
ByVal hWndInsertAfter As Long, _
ByVal X As Long, ByVal Y As Long, _
ByVal cx As Long, ByVal cy As Long, _
ByVal wFlags As Long) As Long
'
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const HWND_TOPMOST = -1

Private Sub cmdCancel_Click()
    Cancel = MsgBox("Anda yakin ingin" & _
    " membatalkan Sleep Timer", _
    vbYesNo + vbQuestion + _
    vbDefaultButton2) = vbNo
    If Not Cancel Then
        Unload Me: On Error Resume Next
        End
    Else
        If MyLostTime <> 0 Then
           frmSleep.tmrSleep.Enabled = True
        End If
    End If
End Sub

Private Sub Form_Load()
'form Always On Top
    SetWindowPos hwnd, HWND_TOPMOST, _
0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub

Private Sub Form_Resize()
    Me.Move Screen.Width - Me.Width, Screen.Height - Me.Height
End Sub

