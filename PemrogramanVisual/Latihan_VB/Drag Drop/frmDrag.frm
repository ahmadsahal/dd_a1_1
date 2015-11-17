VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form frmDrag 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "Drag Drop"
   ClientHeight    =   4635
   ClientLeft      =   1500
   ClientTop       =   2040
   ClientWidth     =   7185
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "frmDrag.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4635
   ScaleWidth      =   7185
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   2880
      ScaleHeight     =   4455
      ScaleWidth      =   4215
      TabIndex        =   4
      Top             =   120
      Width           =   4215
      Begin VB.Image Image1 
         DragIcon        =   "frmDrag.frx":0442
         Height          =   4335
         Left            =   120
         Top             =   0
         Visible         =   0   'False
         Width           =   4095
      End
   End
   Begin MCI.MMControl MMControl1 
      Height          =   495
      Left            =   -480
      TabIndex        =   3
      Top             =   4200
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   873
      _Version        =   393216
      Frames          =   10000
      BorderStyle     =   0
      RecordMode      =   0
      Silent          =   -1  'True
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H000000C0&
      DragIcon        =   "frmDrag.frx":0884
      ForeColor       =   &H00FFFF00&
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1935
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      DragIcon        =   "frmDrag.frx":0CC6
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2190
      Left            =   120
      Pattern         =   "*.txt;*.bmp;*.jpg;*.mp3;*.wmf;*.mid;*.wav;*.jpeg;*.gif"
      TabIndex        =   1
      Top             =   2400
      Width           =   2655
   End
   Begin VB.DirListBox Dir1 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      DragIcon        =   "frmDrag.frx":1108
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1710
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   2655
   End
End
Attribute VB_Name = "frmDrag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim nTemplate$, Pindah$
Dim X&

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_MouseDown(Button _
As Integer, Shift As Integer, _
X As Single, Y As Single)
    File1.DragIcon = Drive1.DragIcon
    File1.Drag
End Sub

Private Sub Form_Load()
    MMControl1.Left = 0 - MMControl1.Width
    MsgBox "Pilih file berekstensi " & _
    "bmp, jpg, gif, txt atau mp3", 64
End Sub

Sub DragDanDrop()
    nTemplate = Right(File1.FileName, 3)
    If Mid(File1.Path, Len(File1.Path)) = "\" Then
      Pindah = File1.Path & File1.FileName
    Else
      Pindah = File1.Path & "\" & File1.FileName
    End If
    Image1.Picture = LoadPicture("")
    Select Case nTemplate
    Case "txt"
        X = Shell("Notepad " + Pindah, 1)
        Picture1.Cls
    Case "bmp", "wmf", "jpg", "jpeg", "gif"
        Me.Image1.Visible = True
        Image1.Picture = LoadPicture(Pindah)
        If Image1.Picture.Width > Picture1.Width Or _
            Image1.Picture.Height > Picture1.Height Then
            Image1.Width = Image1.Picture.Width \ 5
            Image1.Height = Image1.Picture.Height \ 5
            Image1.Stretch = True
        End If
        Image1.Move (Picture1.ScaleWidth - _
        Image1.Width) \ 2, (Picture1.ScaleHeight _
        - Image1.Height) \ 2
        Picture1.Cls
    Case "mp3"
        Me.Image1.Visible = False
        MMControl1.Command = "Close"
        MMControl1.FileName = Pindah
        MMControl1.Command = "Open"
        MMControl1.Command = "Play"
        Picture1.Print "Memainkan file Audio..."
    Case Else
        MsgBox "File tidak terbaca", 64, "Maaf"
    End Select
End Sub

Private Sub Image1_DragDrop(Source As Control, _
X As Single, Y As Single)
    Me.DragDanDrop
End Sub

Private Sub Picture1_DragDrop(Source As Control, _
X As Single, Y As Single)
    Me.DragDanDrop
End Sub
