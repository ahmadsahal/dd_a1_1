VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   ScaleHeight     =   2400
   ScaleWidth      =   3750
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Input: Membuka file untuk dibaca

Private Sub Form_Activate()
    Dim NamaBarang, Jumlah, HargaSatuan As String
    Dim TotalHarga As Currency
    On Error GoTo Salah
    Open App.Path & "\DATA.DAT" For Input As #1
        Input #1, NamaBarang, Jumlah, HargaSatuan
    Close
    Print "Nama barang: "; NamaBarang
    Print "Jumlah : "; Jumlah
    Print "Harga Satuan = "; HargaSatuan
    TotalHarga = Jumlah * HargaSatuan
    Print "Total harga = "; Format(TotalHarga, "Currency")
    Print
    Print "Ukuran file: "; FileLen("C:\VB6\DATA.DAT"); " byte"
    Exit Sub
Salah:
    MsgBox "File belum dibuat, buat dahulu dengan Open!"
    End
End Sub

