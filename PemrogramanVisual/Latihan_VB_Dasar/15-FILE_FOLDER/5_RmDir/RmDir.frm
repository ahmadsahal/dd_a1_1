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
'RmDir: Menghapus Directory

Private Sub Form_Load()
On Error Resume Next
    MkDir "C:\Program" 'Buat folder Program
    MsgBox "File C:\PROGRAM sudah dibuat"
    RmDir "C:\Program" 'Hapus kembali
    MsgBox "File C:\PROGRAM dihapus kembali"
    End
End Sub

