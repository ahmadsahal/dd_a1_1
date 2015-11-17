VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Cetak Barang"
   ClientHeight    =   795
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   3540
   LinkTopic       =   "Form1"
   ScaleHeight     =   795
   ScaleWidth      =   3540
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu MnuFileReport 
      Caption         =   "&Report"
      Begin VB.Menu MnuFileReportTampilkan 
         Caption         =   "&Tampilkan"
      End
      Begin VB.Menu MnuFileReportCetak 
         Caption         =   "&Cetak"
      End
   End
   Begin VB.Menu mnuFileKeluar 
      Caption         =   "&Keluar"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuFileKeluar_Click()
    Unload ReportBarang
    Unload Me
End Sub

Private Sub MnuFileReportCetak_Click()
    ReportBarang.PrintReport
End Sub

Private Sub MnuFileReportTampilkan_Click()
    ReportBarang.Show
End Sub
