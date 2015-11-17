VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Begin VB.Form frmMultimedia 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4080
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6300
   Icon            =   "frmMultimedia.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   6300
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   6165
      _Version        =   393216
      Style           =   1
      Tab             =   2
      TabHeight       =   520
      BackColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Player"
      TabPicture(0)   =   "frmMultimedia.frx":044A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "shaMulti"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "dlgMulti"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "picLayar"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "picBar"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "picButton"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Play List"
      TabPicture(1)   =   "frmMultimedia.frx":0466
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "imgOpen1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "imgHapus1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "imgHapus"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "imgOpen"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "lstLagu"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "About"
      TabPicture(2)   =   "frmMultimedia.frx":0482
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "lblAbout"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "imgMultimedia"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "pclMultimedia"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "tmrLabel"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "tmrAbout"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).ControlCount=   5
      Begin VB.Timer tmrAbout 
         Interval        =   1
         Left            =   120
         Top             =   1320
      End
      Begin VB.Timer tmrLabel 
         Interval        =   100
         Left            =   600
         Top             =   1320
      End
      Begin VB.ListBox lstLagu 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1950
         ItemData        =   "frmMultimedia.frx":049E
         Left            =   -74880
         List            =   "frmMultimedia.frx":04A0
         MultiSelect     =   2  'Extended
         TabIndex        =   8
         ToolTipText     =   "Lagu"
         Top             =   480
         Width           =   5775
      End
      Begin VB.PictureBox picButton 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   -74880
         ScaleHeight     =   1095
         ScaleWidth      =   5775
         TabIndex        =   7
         Top             =   2160
         Width           =   5775
         Begin VB.Image imgBack1 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   1200
            Picture         =   "frmMultimedia.frx":04A2
            Stretch         =   -1  'True
            ToolTipText     =   "Previous"
            Top             =   120
            Width           =   1335
         End
         Begin VB.Image imgNext1 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   2880
            Picture         =   "frmMultimedia.frx":1904
            Stretch         =   -1  'True
            ToolTipText     =   "Browse..."
            Top             =   120
            Width           =   1335
         End
         Begin VB.Image imgStop1 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   3720
            Picture         =   "frmMultimedia.frx":2D66
            Stretch         =   -1  'True
            ToolTipText     =   "Browse..."
            Top             =   600
            Width           =   1335
         End
         Begin VB.Image imgPause1 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   2040
            Picture         =   "frmMultimedia.frx":41C8
            Stretch         =   -1  'True
            ToolTipText     =   "Browse..."
            Top             =   600
            Width           =   1335
         End
         Begin VB.Image imgPlay1 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   360
            Picture         =   "frmMultimedia.frx":562A
            Stretch         =   -1  'True
            ToolTipText     =   "Play/Pause"
            Top             =   600
            Width           =   1335
         End
         Begin VB.Image imgNext 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   2880
            Picture         =   "frmMultimedia.frx":6A8C
            Stretch         =   -1  'True
            ToolTipText     =   "Next"
            Top             =   120
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Image imgStop 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   3720
            Picture         =   "frmMultimedia.frx":82E2
            Stretch         =   -1  'True
            ToolTipText     =   "Stop"
            Top             =   600
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Image imgPlay 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   360
            Picture         =   "frmMultimedia.frx":9B38
            Stretch         =   -1  'True
            ToolTipText     =   "Play"
            Top             =   600
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Image imgBack 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   1200
            Picture         =   "frmMultimedia.frx":B38E
            Stretch         =   -1  'True
            ToolTipText     =   "Back"
            Top             =   120
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Image imgPause 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   2040
            Picture         =   "frmMultimedia.frx":CBE4
            Stretch         =   -1  'True
            ToolTipText     =   "Pause"
            Top             =   600
            Visible         =   0   'False
            Width           =   1335
         End
      End
      Begin VB.PictureBox picBar 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         FillColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   -74880
         ScaleHeight     =   495
         ScaleMode       =   0  'User
         ScaleWidth      =   104.336
         TabIndex        =   5
         Top             =   1560
         Width           =   5775
         Begin VB.PictureBox picProgress 
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   495
            Left            =   0
            Picture         =   "frmMultimedia.frx":E43A
            ScaleHeight     =   495
            ScaleLeft       =   100
            ScaleMode       =   0  'User
            ScaleWidth      =   100
            TabIndex        =   6
            Top             =   0
            Width           =   495
         End
      End
      Begin VB.PictureBox picLayar 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   -74760
         ScaleHeight     =   735
         ScaleWidth      =   5535
         TabIndex        =   1
         Top             =   600
         Width           =   5535
         Begin VB.Label lblStatus 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Stop"
            BeginProperty Font 
               Name            =   "System"
               Size            =   19.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H008080FF&
            Height          =   480
            Left            =   3120
            TabIndex        =   4
            Top             =   240
            Width           =   885
         End
         Begin VB.Label lblWaktu 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "00:00"
            BeginProperty Font 
               Name            =   "Bookman Old Style"
               Size            =   21.75
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   525
            Left            =   4200
            TabIndex        =   3
            Top             =   240
            Width           =   1305
         End
         Begin VB.Label lblJudul 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Copyright (2004) by Viansastra"
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
            TabIndex        =   2
            Top             =   0
            Width           =   3855
         End
      End
      Begin MSComDlg.CommonDialog dlgMulti 
         Left            =   -74880
         Top             =   2280
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin PicClip.PictureClip pclMultimedia 
         Left            =   3360
         Top             =   600
         _ExtentX        =   10478
         _ExtentY        =   4842
         _Version        =   393216
         Rows            =   3
         Cols            =   6
         Picture         =   "frmMultimedia.frx":E744
      End
      Begin VB.Image imgMultimedia 
         Height          =   1815
         Left            =   1200
         Stretch         =   -1  'True
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label lblAbout 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Created By Viansastra"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Top             =   2760
         Width           =   2715
      End
      Begin VB.Image imgOpen 
         Appearance      =   0  'Flat
         Height          =   435
         Left            =   -74760
         Picture         =   "frmMultimedia.frx":176CE
         Stretch         =   -1  'True
         ToolTipText     =   "Open"
         Top             =   2640
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.Image imgHapus 
         Appearance      =   0  'Flat
         Height          =   465
         Left            =   -73200
         Picture         =   "frmMultimedia.frx":18F24
         Stretch         =   -1  'True
         ToolTipText     =   "Hapus"
         Top             =   2640
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.Shape shaMulti 
         BorderColor     =   &H00000000&
         BorderWidth     =   7
         Height          =   975
         Left            =   -74880
         Shape           =   4  'Rounded Rectangle
         Top             =   480
         Width           =   5775
      End
      Begin VB.Image imgHapus1 
         Appearance      =   0  'Flat
         Height          =   465
         Left            =   -73200
         Picture         =   "frmMultimedia.frx":1A77A
         Stretch         =   -1  'True
         ToolTipText     =   "Play/Pause"
         Top             =   2640
         Width           =   1260
      End
      Begin VB.Image imgOpen1 
         Appearance      =   0  'Flat
         Height          =   435
         Left            =   -74760
         Picture         =   "frmMultimedia.frx":1BBDC
         Stretch         =   -1  'True
         ToolTipText     =   "Previous"
         Top             =   2640
         Width           =   1185
      End
   End
   Begin MCI.MMControl mmcMulti 
      Height          =   375
      Left            =   360
      TabIndex        =   10
      Top             =   1320
      Width           =   1980
      _ExtentX        =   3493
      _ExtentY        =   661
      _Version        =   393216
      Frames          =   2
      PrevEnabled     =   -1  'True
      NextEnabled     =   -1  'True
      PlayEnabled     =   -1  'True
      PauseEnabled    =   -1  'True
      BackEnabled     =   -1  'True
      StepEnabled     =   -1  'True
      StopEnabled     =   -1  'True
      RecordEnabled   =   -1  'True
      EjectEnabled    =   -1  'True
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Timer tmrJudul 
      Interval        =   100
      Left            =   2880
      Top             =   1320
   End
   Begin VB.Timer tmrWaktu 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2400
      Top             =   1320
   End
   Begin VB.Image imgTlBar 
      Appearance      =   0  'Flat
      Height          =   435
      Left            =   0
      Picture         =   "frmMultimedia.frx":1D03E
      Stretch         =   -1  'True
      ToolTipText     =   "Previous"
      Top             =   0
      Width           =   4905
   End
   Begin VB.Image imgMinimized 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   5400
      Picture         =   "frmMultimedia.frx":1D59B
      Stretch         =   -1  'True
      ToolTipText     =   "Minimized"
      Top             =   120
      Width           =   375
   End
   Begin VB.Image imgPower 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5880
      Picture         =   "frmMultimedia.frx":1D655
      Stretch         =   -1  'True
      ToolTipText     =   "Tutup"
      Top             =   120
      Width           =   375
   End
   Begin VB.Image imgLoad 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4920
      Picture         =   "frmMultimedia.frx":1DA97
      Stretch         =   -1  'True
      ToolTipText     =   "Buka CD.."
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "frmMultimedia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ReleaseCapture _
Lib "user32" () As Long

Private Declare Function SendMessage Lib "user32" _
Alias "SendMessageA" (ByVal Jend As Long, _
ByVal Psn As Long, ByVal Arg1 As Long, _
Arg2 As Any) As Long

Private Const HTCAPTION = 2
Private Const WM_NCLBUTTONDOWN = &HA1

Private Declare Function mciSendString Lib _
"winmm.dll" Alias "mciSendStringA" _
(ByVal Perintah As String, _
ByVal NKembali As String, _
ByVal PjKembali As Long, _
ByVal Panggil As Long) As Long

Private Judul$, Y, Ubah
Private Menit%, Detik%, Awal As Byte
Private Batas%, BukaCD As Byte

Private Sub Form_Activate()
    Judul = Me.lstLagu.Text
    Me.lblAbout.ForeColor = vbBlack
    Me.mmcMulti.Left = -5000
End Sub

Private Sub Form_Load()
    Ubah = 0
End Sub

Private Sub Form_Resize()
    If Me.WindowState <> vbNormal Then
        Me.Caption = "Viansastra Multimedia"
        Me.BorderStyle = 0
        Me.Icon = Me.Icon
    Else
        Me.Caption = ""
        Me.BorderStyle = 0
    End If
End Sub

Private Sub imgBack_Click()
Me.imgStop_Click
    With Me
        On Error Resume Next
        .lstLagu.ListIndex = _
        .lstLagu.ListIndex - 1
        If Err Then
        .lstLagu.ListIndex = 0
        End If
        Judul = Me.lstLagu.Text
        .mmcMulti.Command = "Close"
        .mmcMulti.FileName = Judul
        .mmcMulti.Command = "Open"
        .mmcMulti.Command = "Play"
        .lblJudul.Caption = Judul
        .lblStatus.Caption = "Play"
        .tmrWaktu.Enabled = True
    End With
End Sub

Private Sub imgBack1_MouseMove( _
Button As Integer, Shift As Integer, _
X As Single, Y As Single)
    Me.imgBack.Visible = True
    Me.imgBack1.Visible = False
End Sub

Private Sub imgHapus_Click()
Dim i As Long: On Error Resume Next
For i = 0 To lstLagu.ListCount - 1
    If lstLagu.Selected(i) Then
        lstLagu.RemoveItem i
    End If
Next
For i = 0 To lstLagu.ListCount - 1
    lstLagu.Selected(i) = False
Next
End Sub

Private Sub imgHapus1_MouseMove( _
Button As Integer, Shift As Integer, _
X As Single, Y As Single)
    Me.imgHapus.Visible = True
    Me.imgHapus1.Visible = False
End Sub

Private Sub imgLoad_MouseUp( _
Button As Integer, Shift As Integer, _
X As Single, Y As Single)
On Error Resume Next
    Dim o
    Me.imgLoad.BorderStyle = 0
    If BukaCD = 0 Then
        X = mciSendString( _
        "set CDAudio door open", o, 127, 0)
        BukaCD = 1
    Else: BukaCD = 0
        X = mciSendString( _
        "set CDAudio door closed", o, 127, 0)
    End If
End Sub

Private Sub imgMinimized_Click()
    Me.WindowState = vbMinimized
End Sub

Sub imgNext_Click(): On Error Resume Next
    Me.imgStop_Click
    With Me
        .lstLagu.ListIndex = _
        .lstLagu.ListIndex + 1
        
    If Not Err Then
        Judul = Me.lstLagu.Text
        .mmcMulti.Command = "Close"
        .mmcMulti.FileName = Judul
        .mmcMulti.Command = "Open"
        .mmcMulti.Command = "Play"
        .lblJudul.Caption = Judul
        .lblStatus.Caption = "Play"
        .tmrWaktu.Enabled = True
    End If
    End With
End Sub

Private Sub imgNext1_MouseMove( _
Button As Integer, Shift As Integer, _
X As Single, Y As Single)
    Me.imgNext.Visible = True
    Me.imgNext1.Visible = False
End Sub

Private Sub imgOpen1_MouseMove( _
Button As Integer, Shift As Integer, _
X As Single, Y As Single)
    Me.imgOpen.Visible = True
    Me.imgOpen1.Visible = False
End Sub

Private Sub imgPause_Click()
    Me.mmcMulti.Command = "Pause"
    Me.tmrWaktu.Enabled = False
    Me.lblStatus.Caption = "Pause"
End Sub

Private Sub imgPause1_MouseMove( _
Button As Integer, Shift As Integer, _
X As Single, Y As Single)
    imgPause.Visible = True
    imgPause1.Visible = False
End Sub

Sub imgPlay_Click(): On Error Resume Next
If Me.lstLagu.Text <> "" Then
        With Me.mmcMulti
        Judul = Me.lstLagu.Text
            .FileName = Judul
            Me.lblJudul.Caption = Judul
            .Command = "Open"
            .Command = "Play"
            Me.lblStatus.Caption = "Play"
            Me.tmrWaktu.Enabled = True
        End With
Else
    MsgBox "Tentukan alamat lagu", 48
End If
End Sub

Private Sub imgPlay_MouseUp( _
Button As Integer, Shift As Integer, _
X As Single, Y As Single)
    Me.imgPlay.BorderStyle = 0
End Sub

Private Sub imgPlay1_MouseMove( _
Button As Integer, Shift As Integer, _
X As Single, Y As Single)
    imgPlay.Visible = True
    imgPlay1.Visible = False
End Sub

Private Sub imgPower_Click()
    Unload Me
End Sub

Sub imgStop_Click()
With Me
    .mmcMulti.Command = "Close"
    .tmrWaktu.Enabled = False
    .lblWaktu.Caption = "00:00"
    .lblStatus.Caption = "Stop"
    .picProgress.Left = 0
    Detik = 0: Menit = 0: Awal = 0
End With
End Sub

Private Sub imgStop1_MouseMove( _
Button As Integer, Shift As Integer, _
X As Single, Y As Single)
    imgStop.Visible = True
    imgStop1.Visible = False
End Sub

Private Sub imgTlbar_MouseDown( _
Button As Integer, Shift As Integer, _
X As Single, Y As Single)
    On Error Resume Next
    Dim Aksi&, Ret&
        If Button = 1 Then
            Aksi = ReleaseCapture()
            Ret = SendMessage(hWnd, _
                  WM_NCLBUTTONDOWN, _
                  HTCAPTION, 0)
        End If
End Sub

Private Sub lstLagu_DblClick()
        With Me.mmcMulti
            .Command = "Close"
            Judul = Me.lstLagu.Text
            .FileName = Judul
            Me.lblJudul.Caption = Judul
            .Command = "Open"
            .Command = "Play"
            Me.lblStatus.Caption = "Play"
            Me.tmrWaktu.Enabled = True
        End With
End Sub

Private Sub picBar_MouseMove( _
Button As Integer, Shift As Integer, _
X As Single, Y As Single)
        With Me
        .imgBack.Visible = False
        .imgNext.Visible = False
        .imgStop.Visible = False
        .imgPause.Visible = False
        .imgPlay.Visible = False
        .imgBack1.Visible = True
        .imgNext1.Visible = True
        .imgStop1.Visible = True
        .imgPause1.Visible = True
        .imgPlay1.Visible = True
    End With
End Sub

Private Sub picButton_MouseMove( _
Button As Integer, Shift As Integer, _
X As Single, Y As Single)
    With Me
        .imgBack.Visible = False
        .imgNext.Visible = False
        .imgStop.Visible = False
        .imgPause.Visible = False
        .imgPlay.Visible = False
        .imgBack1.Visible = True
        .imgNext1.Visible = True
        .imgStop1.Visible = True
        .imgPause1.Visible = True
        .imgPlay1.Visible = True
    End With
End Sub

Private Sub picGroup_MouseMove( _
Index As Integer, Button As Integer, _
Shift As Integer, X As Single, Y As Single)
    Me.imgOpen.Visible = False
    Me.imgHapus.Visible = False
    Me.imgOpen1.Visible = True
    Me.imgHapus1.Visible = True
End Sub

Private Sub tmrAbout_Timer()
    Y = Y + 1: If Y = 18 Then Y = 0
    
    imgMultimedia.Picture = _
    pclMultimedia.GraphicCell(Y)
    
End Sub

Private Sub tmrJudul_Timer()
    Me.lblJudul.Left = Me.lblJudul.Left - 50
    If Me.lblJudul.Left < _
        0 - Me.lblJudul.Width Then
        Me.lblJudul.Left = Me.picLayar.Width
    End If
End Sub

Private Sub tmrLabel_Timer()
    Ubah = Ubah + 1
    Me.lblAbout.ForeColor = QBColor(Ubah)
    If Ubah = 15 Then Ubah = 1
End Sub

Private Sub tmrWaktu_Timer()
    Detik = Detik + 1
    If Detik > 59 Then
        Detik = 0
        Menit = Menit + 1
    End If
    Me.lblWaktu.Caption = Format(Menit, "00") & _
    ":" & Format(Detik, "00")
End Sub

Private Sub mmcMulti_StatusUpdate()
On Error Resume Next
Dim dLength, mLength, Posisi
    Batas = Int((Me.mmcMulti.Position / _
    Me.mmcMulti.Length) * 100)
    Me.picProgress.Left = Batas
    If Batas = 100 Then
        Batas = 0
        Me.imgNext_Click
    End If
End Sub

Function SelectMultipleFiles( _
CD As CommonDialog, Filter As String, _
Filenames() As String) As Boolean
    On Error Resume Next
    CD.Filter = "File Music (*.mp3;*.mid;*.wav)" & _
    "|*.mp3;*.mid;*.wav;*.CDA;*.DAT"
    CD.FilterIndex = 1
    CD.Flags = cdlOFNAllowMultiselect Or _
    cdlOFNFileMustExist Or _
    cdlOFNExplorer
    CD.DialogTitle = "Pilih lagu"
    CD.MaxFileSize = 10240
    CD.FileName = ""
    CD.CancelError = True
    CD.ShowOpen
    Filenames() = Split(CD.FileName, vbNullChar)
    SelectMultipleFiles = True
End Function

Private Sub imgOpen_Click(): Me.lstLagu.Clear
    Dim Filenames() As String, i As Integer
    If SelectMultipleFiles(dlgMulti, _
        "File Music (*.mp3;*.mid;*.wav)" & _
        "|*.mp3;*.mid;*.wav;*.CDA;*.DAT", Filenames()) Then
    
        If UBound(Filenames) <> 0 Then
            For i = 1 To UBound(Filenames)
                Me.lstLagu.AddItem Filenames(i)
            Next
        Else
        End If
    End If
End Sub

