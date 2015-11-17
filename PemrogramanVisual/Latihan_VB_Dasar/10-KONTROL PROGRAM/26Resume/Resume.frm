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
'Resume: Untuk kembali ke baris
'yang membuat terjadi error

Private Sub Form_Load()
    Dim Jawab As Integer
    On Error GoTo salah
    MkDir "A:\PROGRAM" 'Membuat folder pada drive A
    MsgBox "Folder sukses dibuat!"
    End
    Exit Sub
salah:
    Jawab = MsgBox("Drive tidak acu atau nama folder sudah ada!", vbExclamation + vbRetryCancel)
    If Jawab = vbRetry Then Resume
    End
End Sub

