VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2400
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   3744
   LinkTopic       =   "Form1"
   ScaleHeight     =   2400
   ScaleWidth      =   3744
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'GetAttr: Untuk menampilkan atribut sebuah file

Private Sub Form_Load()
    Dim NamaFile, Atribut As String
    On Error GoTo Salah
    NamaFile = InputBox("Ketik nama file: ")
    Atribut = GetAttr(NamaFile)
    Select Case Atribut
        Case 0: MsgBox "Normal"
        Case 1: MsgBox "ReadOnly"
        Case 2: MsgBox "Hidden"
        Case 4: MsgBox "System"
        Case 8: MsgBox "Volume"
        Case 16: MsgBox "Directory"
        Case 32: MsgBox "Archive"
    End Select
    End
Salah:
    MsgBox "File tidak ditemukan!"
    End
End Sub
