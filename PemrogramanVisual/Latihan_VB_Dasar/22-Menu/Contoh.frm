VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Database 1.0"
   ClientHeight    =   2400
   ClientLeft      =   132
   ClientTop       =   816
   ClientWidth     =   3744
   LinkTopic       =   "Form1"
   ScaleHeight     =   6420
   ScaleWidth      =   9504
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu MenuDatabase 
      Caption         =   "Database"
      Begin VB.Menu MenuTambah 
         Caption         =   "Tambah Data"
      End
      Begin VB.Menu MenuCari 
         Caption         =   "Cari Data"
      End
      Begin VB.Menu MenuGanti 
         Caption         =   "Ganti Data"
      End
      Begin VB.Menu MenuHapus 
         Caption         =   "Hapus Data"
      End
   End
   Begin VB.Menu MenuKeluar 
      Caption         =   "Keluar"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MenuCari_Click()
    Form3.Show
End Sub

Private Sub MenuGanti_Click()
    Form4.Show
End Sub

Private Sub MenuHapus_Click()
    Form5.Show
End Sub

Private Sub MenuKeluar_Click()
    End
End Sub

Private Sub MenuTambah_Click()
    Form2.Show
End Sub
