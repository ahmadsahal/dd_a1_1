VERSION 5.00
Begin VB.Form frmCetakPelanggan 
   Caption         =   "Cetak Pelanggan"
   ClientHeight    =   735
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   735
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBatal 
      Caption         =   "Batal"
      Height          =   492
      Left            =   2640
      TabIndex        =   1
      Top             =   120
      Width           =   1644
   End
   Begin VB.CommandButton cmdCetak 
      Caption         =   "Cetak"
      Height          =   492
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1644
   End
End
Attribute VB_Name = "frmCetakPelanggan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBatal_Click()
    Unload Me
End Sub

Private Sub cmdCetak_Click()
    CetakPelanggan.Show
End Sub
