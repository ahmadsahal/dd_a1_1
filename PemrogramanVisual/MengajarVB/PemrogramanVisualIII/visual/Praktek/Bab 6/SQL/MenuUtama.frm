VERSION 5.00
Begin VB.Form Menu 
   Caption         =   "Menu Program Pembelian dan  Penjualan Barang"
   ClientHeight    =   2970
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7005
   LinkTopic       =   "Form1"
   ScaleHeight     =   2970
   ScaleWidth      =   7005
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuInput 
      Caption         =   "&Input"
      Index           =   1
      Begin VB.Menu mnuPemasok 
         Caption         =   "Data Pemasok"
      End
      Begin VB.Menu mnBarang 
         Caption         =   "Data Barang"
      End
      Begin VB.Menu mnPelanggan 
         Caption         =   "Data Pelanggan"
      End
   End
   Begin VB.Menu mnucari 
      Caption         =   "&Cari Data"
      Begin VB.Menu mnucaribarang 
         Caption         =   "Barang"
      End
   End
   Begin VB.Menu mnTransaksi 
      Caption         =   "&Transaksi"
      Begin VB.Menu mnuBeli 
         Caption         =   "Pembelian"
      End
      Begin VB.Menu mnuJual 
         Caption         =   "Penjualan"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit Data"
      Begin VB.Menu mnuEditBarang 
         Caption         =   "Barang"
      End
      Begin VB.Menu mnuEditPemasok 
         Caption         =   "Pemasok"
      End
      Begin VB.Menu mnuPelanggan 
         Caption         =   "Pelanggan"
      End
   End
   Begin VB.Menu mnCetak 
      Caption         =   "&Cetak"
      Begin VB.Menu mnuCetakPemasok 
         Caption         =   "Pemasok"
      End
      Begin VB.Menu mnuCetakBarang 
         Caption         =   "Barang"
      End
      Begin VB.Menu mnuCetakPelanggan 
         Caption         =   "Pelanggan"
      End
      Begin VB.Menu mnuCetakBeli 
         Caption         =   "Pembelian"
      End
      Begin VB.Menu mnuCetakJual 
         Caption         =   "Penjualan"
      End
   End
   Begin VB.Menu mnTutup 
      Caption         =   "Tutup &Program"
   End
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub mnBarang_Click()
    frmBarang.Show
End Sub

Private Sub mnPelanggan_Click()
    frmPelanggan.Show
End Sub

Private Sub mnTutup_Click()
    End
End Sub

Private Sub mnuBeli_Click()
     frmInBeli.Show
End Sub


Private Sub mnucaribarang_Click()
    frmCariBarang.Show
End Sub

Private Sub mnuCetakBarang_Click()
    frmCetakBarang.Show
End Sub
Private Sub mnuCetakBeli_Click()
    frmCetakBeli.Show
End Sub

Private Sub mnuCetakJual_Click()
    frmCetakJual.Show
End Sub

Private Sub mnuCetakPelanggan_Click()
    frmCetakPelanggan.Show
End Sub

Private Sub mnuCetakPemasok_Click()
    frmCetakPemasok.Show
End Sub

Private Sub mnuJual_Click()
    frmInJual.Show
End Sub

Private Sub mnuPemasok_Click()
    frmPemasok.Show
End Sub
Private Sub mnuEditBarang_Click()
    frmEditBarang.Show
End Sub

Private Sub mnuEditPemasok_Click()
    frmEditPemasok.Show
End Sub

Private Sub mnuPelanggan_Click()
    frmEditPelanggan.Show
End Sub
