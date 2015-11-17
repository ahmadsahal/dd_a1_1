VERSION 5.00
Begin VB.Form frmSleep 
   BackColor       =   &H0000FFFF&
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   2250
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   6405
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawStyle       =   5  'Transparent
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Bookman Old Style"
      Size            =   12
      Charset         =   0
      Weight          =   600
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmSleep.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   6405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Batal"
      Height          =   375
      Left            =   4080
      TabIndex        =   6
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Set"
      Default         =   -1  'True
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Timer tmrAniLabel 
      Interval        =   100
      Left            =   600
      Top             =   1920
   End
   Begin VB.Timer tmrNow 
      Interval        =   1
      Left            =   1080
      Top             =   1920
   End
   Begin VB.Timer tmrSleep 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   0
      Top             =   1920
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      DrawWidth       =   10
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2400
      ScaleHeight     =   735
      ScaleWidth      =   3015
      TabIndex        =   0
      Top             =   960
      Width           =   3015
   End
   Begin VB.Label lblNow 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   105
      TabIndex        =   4
      Top             =   960
      Width           =   1365
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   1680
      X2              =   6120
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00008080&
      BorderWidth     =   2
      Index           =   0
      X1              =   1680
      X2              =   6120
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sleep Timer"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   27.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Index           =   0
      Left            =   2160
      TabIndex        =   3
      Top             =   0
      Width           =   3570
   End
   Begin VB.Label lblStart 
      BackStyle       =   0  'Transparent
      Caption         =   "lblStart"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblFinish 
      BackStyle       =   0  'Transparent
      Caption         =   "lblFinis"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   1560
      X2              =   1560
      Y1              =   120
      Y2              =   2160
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      Index           =   0
      X1              =   1560
      X2              =   1560
      Y1              =   120
      Y2              =   2160
   End
   Begin VB.Menu mnuH 
      Caption         =   "harga"
      Visible         =   0   'False
      Begin VB.Menu mnuAbout 
         Caption         =   "&Tentang.."
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Tutup"
      End
   End
End
Attribute VB_Name = "frmSleep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function _
ExitWindowsEx Lib "user32" _
( _
ByVal uFlags As Long, _
ByVal dwReserved As Long _
) As Long

Private Awal
Private MySecond, MyMinute, MyHour, MyTime
Private MyLostTime, MyInput, MyTimerEnd
Private RMinute, RHour, RSecond, i

Private Sub cmdCancel_Click():
On Error Resume Next
    Unload Me
End Sub

Private Sub cmdOK_Click():  On Error Resume Next
If Awal = 0 Then
    MyInput = InputBox("Masukkan waktu dalam menit", _
        "Sleep Timer", 15)
    If MyInput <> "" Then
            Me.lblStart.Caption = Format(Time, "hh:mm:ss")
            MyLostTime = Val(MyInput)
            MyMinute = Minute(lblStart)
            MySecond = Second(lblStart)
            MyHour = Hour(lblStart)
            RSecond = MySecond
            RMinute = MyMinute + MyInput
            If RMinute > 59 Then
                RMinute = _
                RMinute - 60: RHour = MyHour + 1
                If RMinute > 59 Then RMinute = _
                RMinute - 60: RHour = MyHour + 2
                If RMinute > 59 Then RMinute = _
                RMinute - 60: RHour = MyHour + 3
                If RMinute > 59 Then RMinute = _
                RMinute - 60: RHour = MyHour + 4
                If RMinute > 59 Then RMinute = _
                RMinute - 60: RHour = MyHour + 5
                If RMinute > 59 Then RMinute = _
                RMinute - 60: RHour = MyHour + 6
                If RMinute > 59 Then RMinute = _
                RMinute - 60: RHour = MyHour + 7
            Else
                RHour = MyHour
            End If
            cmdOK.Caption = "Simpan"
            Awal = 1
            MyTimerEnd = Format(RHour, "00") + _
            ":" + Format(RMinute, "00") + ":" + _
            Format(RSecond, "00")
            Me.lblFinish.Caption = CStr(MyTimerEnd)
            Me.tmrSleep.Enabled = True
    End If
Else:  Me.Hide: frmWaktu.Show
End If
End Sub
Sub Tutup():
    ExitWindowsEx &H45, 1
End Sub
Private Sub Form_Load()
    Load frmWaktu
    frmWaktu.Hide
    MyMinute = 0
    Awal = 0: MyHour = 0
    pic.CurrentX = 0
    pic.CurrentY = 0
    Me.pic.Font.Size = 30
    Me.pic.ForeColor = vbRed
    Me.pic.Font.Bold = True
    pic.Print "00:00:00"
    '
    pic.CurrentX = 0 - 40
    pic.CurrentY = 0 - 40
    Me.pic.Font.Size = 30
    Me.pic.ForeColor = vbYellow
    Me.pic.Font.Bold = True
    pic.Print "00:00:00"
    '
    i = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, _
UnloadMode As Integer):
If MyLostTime <> 0 Then
    Cancel = MsgBox("Anda yakin ingin membatalkan Sleep Timer", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo
    If Not Cancel Then
        Unload Me: On Error Resume Next
        End
    Else
        If MyLostTime <> 0 Then Me.tmrSleep.Enabled = True
    End If
Else
    End
End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
End Sub

Private Sub tmrAniLabel_Timer()
    i = i + 1
    Me.lblTitle(0).ForeColor = QBColor(i)
    If i = 15 Then i = 0
End Sub

Private Sub tmrNow_Timer()
    Me.lblNow.Caption = Format(Time, "hh:mm:ss")
End Sub

Private Sub tmrSleep_Timer(): Dim X, Y
On Error Resume Next
    Me.lblNow.Caption = Format(Time, "hh:mm:ss")
    MyTime = Time - TimeValue(Me.lblFinish.Caption)
    pic.Cls
    X = 0
    Y = 0
    
    pic.CurrentX = X
    pic.CurrentY = Y
    Me.pic.Font.Size = 30
    Me.pic.ForeColor = vbRed
    Me.pic.Font.Bold = True
    pic.Print Format(MyTime, "hh:mm:ss")
    
    pic.CurrentX = X - 40
    pic.CurrentY = Y - 40
    Me.pic.Font.Size = 30
    Me.pic.ForeColor = vbYellow
    Me.pic.Font.Bold = True
    pic.Print Format(MyTime, "hh:mm:ss")
    
    frmWaktu.lblTime(0).Caption = Format(MyTime, "hh:mm:ss")
    frmWaktu.lblTime(1).Caption = Format(MyTime, "hh:mm:ss")
    frmWaktu.Caption = "  Waktu tersisa " & Format(MyTime, "hh:mm:ss")
    Me.pic.ToolTipText = Format(MyTime, "hh:mm:ss")
    On Error Resume Next
    If Format(MyTime, "hh:mm:ss") = "00:00:00" Then Tutup
End Sub


