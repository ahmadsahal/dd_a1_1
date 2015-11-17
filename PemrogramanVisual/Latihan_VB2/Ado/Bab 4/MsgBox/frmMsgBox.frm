VERSION 5.00
Begin VB.Form frmDataEnv 
   Caption         =   "Kotak Pesan"
   ClientHeight    =   765
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2280
   LinkTopic       =   "Form1"
   ScaleHeight     =   765
   ScaleWidth      =   2280
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "Simpan"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmDataEnv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSimpan_Click()
'Meletakkan komponen kotak pesan
    Dim Hasil As Byte
    JudulWindow = "MsgBox Demo"
    Pesan = "File belum disimpan." & "Apakah anda ingin menyimpan file?"
    tipe = vbYesNo + vbQuestion + vbDefaultButton2
'Dapatkan respon pemakai
    Hasil = MsgBox(Pesan, tipe, JudulWindow)
'Evaluasi hasil
    If Hasil = vbYes Then
        msg = "File sukses Disimpan"
    Else
        msg = "File Batal Disimpan"
    End If
    MsgBox msg
End Sub
