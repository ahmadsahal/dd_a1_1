VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3468
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   3168
   LinkTopic       =   "Form1"
   ScaleHeight     =   3468
   ScaleWidth      =   3168
   StartUpPosition =   3  'Windows Default
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
      Left            =   480
      TabIndex        =   1
      Top             =   2880
      Width           =   1932
   End
   Begin VB.PictureBox Picture2 
      Height          =   2412
      Left            =   240
      ScaleHeight     =   2364
      ScaleWidth      =   2604
      TabIndex        =   0
      Top             =   240
      Width           =   2652
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Lokasi = InputBox("Ketik lokasi dan nama gambar. Contoh: C:\GAMBAR.JPG: ", "Memanggil ganbar")
    Picture2.Picture = LoadPicture(Lokasi)
End Sub

