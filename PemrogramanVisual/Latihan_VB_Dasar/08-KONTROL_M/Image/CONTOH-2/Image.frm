VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3012
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   3744
   LinkTopic       =   "Form1"
   ScaleHeight     =   3012
   ScaleWidth      =   3744
   StartUpPosition =   3  'Windows Default
   Begin VB.Image Image1 
      Height          =   492
      Left            =   120
      Top             =   120
      Width           =   972
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'Memasukkan gambar dengan program
    On Error GoTo salah
    'Ganti lokasi gambar dengan gambar Anda
    'Contoh: "C:\GAMBAR\GAMBARKU.JPG"
    Image1.Picture = LoadPicture("Shancai.jpg")
    Exit Sub
salah:
    MsgBox "Gambar tidak ditemukan, periksa lokasi gambar!"
    End
End Sub
