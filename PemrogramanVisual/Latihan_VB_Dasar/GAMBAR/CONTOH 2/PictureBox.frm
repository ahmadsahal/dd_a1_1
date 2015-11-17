VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Memanggail gambar"
   ClientHeight    =   3468
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   3408
   LinkTopic       =   "Form1"
   ScaleHeight     =   3468
   ScaleWidth      =   3408
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2760
      Top             =   2880
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PANGGIL GAMBAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   720
      TabIndex        =   1
      Top             =   2880
      Width           =   1932
   End
   Begin VB.PictureBox Picture2 
      Height          =   2412
      Left            =   240
      ScaleHeight     =   2364
      ScaleWidth      =   2844
      TabIndex        =   0
      Top             =   240
      Width           =   2892
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    CommonDialog1.Filter = "All Files|*.*"
    CommonDialog1.ShowOpen
    Picture2.Picture = LoadPicture(CommonDialog1.FileName)
End Sub

