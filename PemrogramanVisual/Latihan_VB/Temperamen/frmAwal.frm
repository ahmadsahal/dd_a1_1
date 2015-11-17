VERSION 5.00
Begin VB.Form frmAwal 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8100
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   8970
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAwal.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8100
   ScaleWidth      =   8970
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   8055
      Left            =   0
      ScaleHeight     =   8055
      ScaleWidth      =   9015
      TabIndex        =   5
      Top             =   0
      Width           =   9015
      Begin VB.CommandButton cmdOK 
         BackColor       =   &H00FFFFC0&
         Caption         =   "OK"
         Height          =   345
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   7440
         Width           =   1620
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   390
         Index           =   1
         Left            =   2160
         TabIndex        =   3
         Text            =   "INocHI"
         Top             =   6840
         Width           =   5775
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   390
         Index           =   0
         Left            =   2160
         TabIndex        =   1
         Text            =   "VIANSASTRA"
         Top             =   6240
         Width           =   3060
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   3
         Index           =   1
         X1              =   120
         X2              =   8760
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Nama Anda"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   330
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   6240
         Width           =   1650
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Alamat"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   330
         Index           =   1
         Left            =   180
         TabIndex        =   2
         Top             =   6840
         Width           =   1050
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF00FF&
         BorderWidth     =   3
         Index           =   0
         X1              =   120
         X2              =   8640
         Y1              =   6000
         Y2              =   6000
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dikreasikan Oleh: VIANSASTRA"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   4200
         TabIndex        =   8
         Top             =   7560
         Width           =   3960
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SELAMAT DATANG DI PROGRAM TEMPERAMENT TEST"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   21.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   1575
         Left            =   600
         TabIndex        =   7
         Top             =   0
         Width           =   7500
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   $"frmAwal.frx":030A
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   3630
         Left            =   0
         TabIndex        =   6
         Top             =   2040
         Width           =   8790
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "frmAwal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit: Dim A
Private Sub cmdOK_Click()
If Me.Text1(0).Text = "" Then
frmPesan.Lable = "Namamu siapa Nak?": frmPesan.Waktu = 3
Me.Text1(0).SetFocus
ElseIf Me.Text1(1).Text = "" Then
frmPesan.Lable = "Rumahmu di mana De'?": frmPesan.Waktu = 3
Me.Text1(1).SetFocus
Else: frm1.Show
frmPesan.lblPesan.ForeColor = vbBlack
frmPesan.Lable = "NAMAMU  " + Me.Text1(0).Text + "? NAMA YANG BAGUS!"
frmPesan.Waktu = 3
frmPesan.BackColor = vbCyan
NAMA = Me.Text1(0).Text
Alamat = Me.Text1(1).Text
Unload Me
End If:
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
frmPesan.lblPesan.ForeColor = vbWhite
frmPesan.Lable = "MAU KEMANA DE'? JANGAN NGERJAIN DONG!"
frmPesan.Waktu = 2
frmPesan.BackColor = vbBlack
Unload Me: End If
End Sub

Private Sub Form_Resize()
    Me.Move 0, 0, Screen.Width, Screen.Height
    Me.Picture1.Move (Me.ScaleWidth - Me.Picture1.ScaleWidth) \ 2, _
    (Me.ScaleHeight - Me.Picture1.ScaleHeight) / 2
End Sub

Private Sub Text1_Change(Index As Integer)
    Me.Text1(1).SelStart = Len(Me.Text1(1).Text)
    Me.Text1(1).Text = StrConv(Me.Text1(1).Text, 3)
    Me.Text1(0).SelStart = Len(Me.Text1(0).Text)
    Me.Text1(0).Text = StrConv(Me.Text1(0).Text, vbUpperCase)
End Sub
Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
If Me.Text1(0).Text = "" Then
frmPesan.Lable = "Namamu siapa Nak?": frmPesan.Waktu = 3
Me.Text1(0).SetFocus
ElseIf Me.Text1(1).Text = "" Then
frmPesan.Lable = "Rumahmu di mana De'?": frmPesan.Waktu = 3
Me.Text1(1).SetFocus
Else: frm1.Show
frmPesan.lblPesan.ForeColor = vbBlack
frmPesan.Lable = "NAMAMU  " + Me.Text1(0).Text + "? NAMA YANG BAGUS!"
frmPesan.Waktu = 5
frmPesan.BackColor = vbCyan
NAMA = Me.Text1(0).Text
Alamat = Me.Text1(1).Text
Unload Me
End If: End If
End Sub
Private Sub Text1_GotFocus(Index As Integer)
Me.Text1(0).SelStart = 0: Me.Text1(0).SelLength = Len(Me.Text1(0).Text)
Me.Text1(1).SelStart = 0: Me.Text1(1).SelLength = Len(Me.Text1(1).Text)
End Sub
