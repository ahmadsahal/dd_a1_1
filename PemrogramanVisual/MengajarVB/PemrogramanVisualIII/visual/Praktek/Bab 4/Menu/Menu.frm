VERSION 5.00
Begin VB.Form frmMenu 
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
      Begin VB.Menu mnPemasok 
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
      Begin VB.Menu mnBeli 
         Caption         =   "Pembelian"
      End
      Begin VB.Menu mnJual 
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
      Begin VB.Menu mnCetakPemasok 
         Caption         =   "Pemasok"
      End
      Begin VB.Menu mnCetakBarang 
         Caption         =   "Barang"
      End
      Begin VB.Menu mnCetakPelanggan 
         Caption         =   "Pelanggan"
      End
   End
   Begin VB.Menu mnTutup 
      Caption         =   "Tutup &Program"
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub mnTutup_Click()
    Unload Me
End Sub
